"""Isolated native Microsoft Excel automation with bounded COM busy retries.

This module imports pywin32 only when a session starts. It never attaches to
the user's existing Excel instance and never saves workbooks during cleanup.

Sources:
https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open
https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefullrebuild
https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity
"""
from __future__ import annotations

import os
from pathlib import Path
import time
from typing import Callable, TypeVar

XL_CALCULATION_MANUAL = -4135
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
RPC_E_CALL_REJECTED = 0x80010001
RPC_E_SERVERCALL_RETRYLATER = 0x8001010A
_BUSY_HRESULTS = {RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER}
T = TypeVar("T")


class ExcelSessionError(RuntimeError):
    """Excel could not complete an operation or shut down cleanly."""


class ExcelUnavailableError(ExcelSessionError):
    """Native Excel or its Python COM dependency is unavailable."""


def _load_pywin32():
    if os.name != "nt":
        raise ExcelUnavailableError("Native Excel automation requires Windows and installed Microsoft Excel desktop.")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise ExcelUnavailableError("pywin32 is required. Install this project's dependencies with "
                                    "python -m pip install -r requirements.txt (or python -m pip install pywin32).") from exc
    return pythoncom, win32com.client


def _is_excel_busy(error: Exception) -> bool:
    value = getattr(error, "hresult", None)
    if value is None and error.args and isinstance(error.args[0], int):
        value = error.args[0]
    return isinstance(value, int) and value & 0xFFFFFFFF in _BUSY_HRESULTS


class ExcelSession:
    """Context manager owning one hidden ``DispatchEx('Excel.Application')``.

    ``call(lambda: ...)`` retries only COM's two explicit call-rejection codes.
    Wrapping access in a lambda also retries COM property/method resolution.
    Other errors are raised immediately; operations are never silently skipped.

    ``calculate(source)`` must run while source is the only open workbook.
    Excel's full dependency rebuild is application-wide; this restriction keeps
    template workbooks out of that calculation. No RefreshAll/UpdateLink call
    is made. Keep ordinary workbook changes in manual calculation afterwards.

    Use ``track_workbook`` for workbooks created by the caller and
    ``close_workbook`` when closing one before the context exits.
    Call ``restore_calculation(destination)`` before SaveAs so the output keeps
    the calculation mode observed when its source input was first opened.
    """

    def __init__(self, *, max_attempts: int = 6, retry_delay: float = 0.2,
                 max_retry_delay: float = 1.0):
        if isinstance(max_attempts, bool) or not isinstance(max_attempts, int) or max_attempts < 1:
            raise ValueError("max_attempts must be a positive integer.")
        if not 0 <= retry_delay <= max_retry_delay <= 5:
            raise ValueError("Retry delays must satisfy 0 <= retry_delay <= max_retry_delay <= 5 seconds.")
        self.max_attempts = max_attempts
        self.retry_delay = retry_delay
        self.max_retry_delay = max_retry_delay
        self.app = None
        self.cleanup_errors: list[str] = []
        self._pythoncom = None
        self._initialized = False
        self._entered = False
        self._workbooks = []
        self._calculation_settings: dict[int, tuple[int, bool]] = {}

    def __enter__(self):
        if self._entered:
            raise ExcelSessionError("An ExcelSession cannot be entered more than once.")
        self._entered = True
        pythoncom, client = _load_pywin32()
        self._pythoncom = pythoncom
        try:
            pythoncom.CoInitialize()
            self._initialized = True
            try:
                self.app = self.call(lambda: client.DispatchEx("Excel.Application"))
            except Exception as exc:
                raise ExcelUnavailableError("Could not start an isolated Microsoft Excel instance. "
                                            "Install Excel desktop and check that it opens normally.") from exc
            for attribute, value in (("Visible", False), ("ScreenUpdating", False),
                                     ("DisplayAlerts", False), ("EnableEvents", False),
                                     ("AskToUpdateLinks", False),
                                     ("AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE)):
                self.call(lambda attribute=attribute, value=value: setattr(self.app, attribute, value))
            return self
        except BaseException as primary:
            self._cleanup(primary)
            raise

    def call(self, operation: Callable[..., T], *args, **kwargs) -> T:
        """Execute one COM operation, retrying only rejected/not-yet-executed calls."""
        delay = self.retry_delay
        for attempt in range(self.max_attempts):
            try:
                return operation(*args, **kwargs)
            except Exception as exc:
                if not _is_excel_busy(exc) or attempt + 1 == self.max_attempts:
                    raise
                if self._pythoncom is not None:
                    self._pythoncom.PumpWaitingMessages()
                time.sleep(delay)
                delay = min(delay * 2, self.max_retry_delay)
        raise AssertionError("Unreachable retry loop exit")

    def _require_app(self):
        if self.app is None:
            raise ExcelSessionError("ExcelSession is not running; use it inside a with block.")

    def copy_after(self, selection, anchor):
        """Copy a worksheet/group after an explicit sheet in the same instance.

        Dynamic pywin32 wrappers use Missing for omitted optional parameters.
        InvokeTypes stops at the first Missing, so Copy(After=anchor) can lose
        After and create a new workbook. Empty preserves the Before position,
        matching pywin32's generated (makepy) wrappers.
        """
        self._require_app()
        return self.call(lambda: selection.Copy(self._pythoncom.Empty, anchor))

    def move_after(self, selection, anchor):
        """Move a worksheet/group after a sheet without dropping After in COM."""
        self._require_app()
        return self.call(lambda: selection.Move(self._pythoncom.Empty, anchor))

    def add_sheet_after(self, workbook, anchor):
        """Append one ordinary worksheet using explicit COM argument positions."""
        self._require_app()
        return self.call(lambda: workbook.Worksheets.Add(self._pythoncom.Empty, anchor, 1, -4167))

    def track_workbook(self, workbook):
        """Register a workbook created in this session for unsaved cleanup."""
        self._require_app()
        if self.call(lambda: int(workbook.Application.Hwnd)) != self.call(lambda: int(self.app.Hwnd)):
            raise ExcelSessionError("Cannot register a workbook from another Excel instance.")
        if not any(item is workbook for item in self._workbooks):
            self._workbooks.append(workbook)
            self._capture_calculation(workbook)
        return workbook

    def _capture_calculation(self, workbook):
        self._calculation_settings[id(workbook)] = self.call(
            lambda: (int(self.app.Calculation), bool(self.app.CalculateBeforeSave)))

    def open_workbook(self, path: str | Path, *, read_only: bool = True):
        """Open without updating links, write prompts, automatic repairs, or MRU changes."""
        self._require_app()
        filename = Path(path).expanduser().resolve()
        if not filename.is_file():
            raise FileNotFoundError(f"Excel input workbook does not exist: {filename}")
        self.call(lambda: setattr(self.app, "AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE))
        workbook = self.call(lambda: self.app.Workbooks.Open(
            Filename=str(filename), UpdateLinks=0, ReadOnly=read_only,
            Format=self._pythoncom.Empty, Password="", WriteResPassword="",
            IgnoreReadOnlyRecommended=True, Origin=self._pythoncom.Empty,
            Delimiter=self._pythoncom.Empty, Editable=self._pythoncom.Empty,
            Notify=False, Converter=self._pythoncom.Empty, AddToMru=False,
            Local=False, CorruptLoad=0))
        self._workbooks.append(workbook)  # our Open returned it; track before another COM call can fail
        self._capture_calculation(workbook)
        # Some Excel versions reject setting Calculation before any workbook is
        # open. Set it immediately after each open and keep it manual thereafter.
        self.call(lambda: setattr(self.app, "Calculation", XL_CALCULATION_MANUAL))
        self.call(lambda: setattr(self.app, "CalculateBeforeSave", False))
        self.call(lambda: setattr(self.app, "AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE))
        return workbook

    def close_workbook(self, workbook):
        """Close a workbook opened/registered here, explicitly discarding changes."""
        if not any(item is workbook for item in self._workbooks):
            raise ExcelSessionError("Cannot close a workbook that is not owned by this ExcelSession.")
        self.call(lambda: workbook.Close(SaveChanges=False))
        self._workbooks = [item for item in self._workbooks if item is not workbook]
        self._calculation_settings.pop(id(workbook), None)

    def calculate(self, workbook):
        """Rebuild source formula dependencies before any other workbook is opened."""
        self._require_app()
        if len(self._workbooks) != 1 or self._workbooks[0] is not workbook:
            raise ExcelSessionError("Calculate the source while it is the only open workbook, before opening the PPT template.")
        if self.call(lambda: int(self.app.Workbooks.Count)) != 1:
            raise ExcelSessionError("Excel has additional open workbooks. A full rebuild would also calculate them; "
                                    "the source-only calculation was not started.")
        self.call(lambda: self.app.CalculateFullRebuild())
        self.call(lambda: setattr(self.app, "Calculation", XL_CALCULATION_MANUAL))
        self.call(lambda: setattr(self.app, "CalculateBeforeSave", False))

    def restore_calculation(self, workbook):
        """Restore the source's initial application calculation settings before save.

        Excel's calculation mode is application-wide. Close the template before
        calling this method if it must not participate in automatic calculation.
        Untracked startup workbooks prevent the mode change for the same reason.
        """
        self._require_app()
        if not any(item is workbook for item in self._workbooks):
            raise ExcelSessionError("Cannot restore calculation for a workbook not owned by this ExcelSession.")
        if self.call(lambda: int(self.app.Workbooks.Count)) != len(self._workbooks):
            raise ExcelSessionError("Excel has additional or unexpectedly closed workbooks; calculation settings were not restored.")
        settings = self._calculation_settings.get(id(workbook))
        if settings is None:
            raise ExcelSessionError("The workbook's initial calculation settings were not captured.")
        mode, before_save = settings
        self.call(lambda: setattr(self.app, "CalculateBeforeSave", before_save))
        self.call(lambda: setattr(self.app, "Calculation", mode))

    def _cleanup(self, primary=None):
        errors = []
        for workbook in reversed(self._workbooks):
            try:
                self.call(lambda workbook=workbook: workbook.Close(SaveChanges=False))
            except Exception as exc:
                errors.append(f"Could not close a session workbook without saving: {exc}")
        self._workbooks.clear()
        self._calculation_settings.clear()
        if self.app is not None:
            try:
                self.call(lambda: self.app.Quit())
            except Exception as exc:
                errors.append(f"Could not quit the isolated Excel instance: {exc}")
            finally:
                self.app = None
        if self._initialized:
            try:
                self._pythoncom.CoUninitialize()
            except Exception as exc:
                errors.append(f"COM session cleanup failed: {exc}")
            finally:
                self._initialized = False
        self.cleanup_errors.extend(errors)
        if errors and primary is not None and hasattr(primary, "add_note"):
            primary.add_note("Excel cleanup: " + "; ".join(errors))
        return errors

    def __exit__(self, exception_type, exception, traceback):
        errors = self._cleanup(exception)
        if errors and exception is None:
            raise ExcelSessionError("Excel cleanup did not finish cleanly: " + "; ".join(errors))
        return False
