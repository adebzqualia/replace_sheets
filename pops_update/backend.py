"""Conservative native-Excel refresh for existing POPS country worksheets."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import json
import os
import shutil
import tempfile
import time

from .config import CountryUpdate, UpdateConfig
from .excel_session import ExcelSession
from .inspection import WorkbookInfo, check_layout, inspect_workbook
from .references import ReferenceError, canonical_formula, validate_formula


@dataclass(frozen=True)
class CountryPlan:
    update: CountryUpdate
    source_info: WorkbookInfo
    sheets: tuple


@dataclass(frozen=True)
class UpdatePlan:
    config: UpdateConfig
    final_info: WorkbookInfo
    countries: tuple[CountryPlan, ...]


def _sheet_names(wb) -> set[str]:
    return {str(wb.Worksheets(i).Name) for i in range(1, int(wb.Worksheets.Count) + 1)}


def _formula2(cell):
    try:
        return cell.Formula2
    except Exception:
        return cell.Formula


def _scalar(value):
    if isinstance(value, tuple):
        return value[0] if value else None
    return value


def _is_formula(value) -> bool:
    return isinstance(value, str) and value.startswith('=')


def _used_range(ws):
    used = ws.UsedRange
    return int(used.Row), int(used.Column), int(used.Rows.Count), int(used.Columns.Count)


def _matrix_value(value, row, col):
    if isinstance(value, tuple):
        return value[row][col]
    return value if row == 0 and col == 0 else None


def _target_sources(info: WorkbookInfo, update: CountryUpdate):
    for sheet in update.sheets:
        if sheet.source_name not in info.sheets:
            raise ValueError(f"{update.country}: source worksheet {sheet.source_name!r} is missing")


def build_plan(config: UpdateConfig) -> UpdatePlan:
    """Read-only preflight of all XML workbook layouts before Excel is opened."""
    final_info = inspect_workbook(config.final_workbook)
    countries = []
    destination_names = set(final_info.sheet_names)
    for update in config.updates:
        source_info = inspect_workbook(update.file)
        _target_sources(source_info, update)
        planned = []
        for sheet in update.sheets:
            if sheet.target_name not in destination_names:
                raise ValueError(f"{update.country}: destination worksheet {sheet.target_name!r} is missing from final workbook")
            check_layout(source_info.sheets[sheet.source_name], final_info.sheets[sheet.target_name], config.check_text_layout)
            planned.append(sheet)
        countries.append(CountryPlan(update, source_info, tuple(planned)))
    return UpdatePlan(config, final_info, tuple(countries))


def _rename_source_sheets(session, source, mapping, prefix):
    """Rename every mapped source tab present, through temporary names."""
    present = {str(source.Worksheets(i).Name): source.Worksheets(i)
               for i in range(1, int(source.Worksheets.Count) + 1)}
    active = []
    for source_name, suffix in mapping:
        if source_name in present:
            active.append((present[source_name], f"__POPS_UPDATE_{len(active)+1}__"))
    for ws, temporary in active:
        if len(temporary) > 31:
            raise ValueError(f"Temporary Excel worksheet name is too long for {ws.Name!r}")
        session.call(lambda ws=ws, temporary=temporary: setattr(ws, 'Name', temporary))
    renamed = {}
    for index, (ws, temporary) in enumerate(active):
        original = next(name for name, value in present.items() if value is ws)
        suffix = dict(mapping)[original]
        target = f"{prefix}_{suffix}"
        session.call(lambda ws=ws, target=target: setattr(ws, 'Name', target))
        renamed[original] = target
    return renamed


def _defined_names(wb):
    values = {}
    try:
        names = wb.Names
        for i in range(1, int(names.Count) + 1):
            item = names.Item(i)
            values[str(item.Name)] = str(item.RefersTo)
    except Exception:
        pass
    return values


def _validate_source_formulas(source, destination, destination_info: WorkbookInfo, renamed):
    allowed_sheets = _sheet_names(destination)
    allowed_tables = {table.name for sheet in destination_info.sheets.values() for table in sheet.tables}
    source_names = _defined_names(source)
    dest_names = _defined_names(destination)
    for i in range(1, int(source.Worksheets.Count) + 1):
        ws = source.Worksheets(i)
        row, col, rows, cols = _used_range(ws)
        if rows <= 0 or cols <= 0:
            continue
        formulas = ws.Range(ws.Cells(row, col), ws.Cells(row + rows - 1, col + cols - 1)).Formula2
        for r in range(rows):
            for c in range(cols):
                formula = _matrix_value(formulas, r, c)
                if not _is_formula(formula):
                    continue
                if '[' in formula or '#REF!' in formula.upper():
                    raise ValueError(f"{ws.Name}: unsupported external or broken formula at {ws.Cells(row+r, col+c).Address}: {formula}")
                try:
                    validate_formula(formula, sheet_name=str(ws.Name), allowed_sheets=allowed_sheets,
                                     allowed_tables=allowed_tables, source_names=source_names,
                                     destination_names=dest_names)
                except ReferenceError as exc:
                    # Unqualified identifiers are commonly Excel names. If the
                    # destination has no names, reject instead of guessing.
                    raise ValueError(f"{ws.Name}: {exc}") from exc


def _clear_and_transfer(session, source_ws, target_ws, password='a'):
    source_row, source_col, source_rows, source_cols = _used_range(source_ws)
    was_protected = bool(target_ws.ProtectContents)
    if was_protected:
        target_ws.Unprotect(Password=password)
    target_ws.UsedRange.ClearContents()
    if source_rows <= 0 or source_cols <= 0:
        return
    source_range = source_ws.Range(source_ws.Cells(source_row, source_col), source_ws.Cells(source_row + source_rows - 1, source_col + source_cols - 1))
    formulas = source_range.Formula2
    values = source_range.Value2
    target_row, target_col, _, _ = _used_range(target_ws)
    # Keep the source origin. The layout check ensures labels and merges remain
    # at the same coordinates; this prevents silent shifts in fixed links.
    target_row = source_row
    target_col = source_col
    target_range = target_ws.Range(target_ws.Cells(target_row, target_col), target_ws.Cells(target_row + source_rows - 1, target_col + source_cols - 1))
    for r in range(source_rows):
        for c in range(source_cols):
            src_cell = source_ws.Cells(source_row + r, source_col + c)
            dst_cell = target_ws.Cells(target_row + r, target_col + c)
            try:
                if bool(src_cell.MergeCells) and str(src_cell.Address) != str(src_cell.MergeArea.Cells(1, 1).Address):
                    continue
            except Exception:
                pass
            formula = _matrix_value(formulas, r, c)
            value = _matrix_value(values, r, c)
            session.call(lambda dst_cell=dst_cell, formula=formula, value=value: setattr(dst_cell, 'Formula2', formula if _is_formula(formula) else value))
    if was_protected:
        target_ws.Protect(Password=password)


def _formula_snapshot(wb):
    result = {}
    for i in range(1, int(wb.Worksheets.Count) + 1):
        ws = wb.Worksheets(i)
        row, col, rows, cols = _used_range(ws)
        formulas = ws.Range(ws.Cells(row, col), ws.Cells(row + rows - 1, col + cols - 1)).Formula2 if rows and cols else None
        for r in range(rows):
            for c in range(cols):
                value = _matrix_value(formulas, r, c)
                if _is_formula(value):
                    result[(str(ws.Name), ws.Cells(row+r, col+c).Address)] = canonical_formula(value)
    return result


def _check_ref_errors(wb):
    errors = []
    for i in range(1, int(wb.Worksheets.Count) + 1):
        ws = wb.Worksheets(i)
        row, col, rows, cols = _used_range(ws)
        if not rows or not cols:
            continue
        area = ws.Range(ws.Cells(row, col), ws.Cells(row + rows - 1, col + cols - 1))
        values = area.Value2
        formulas = area.Formula2
        for r in range(rows):
            for c in range(cols):
                value = _matrix_value(values, r, c)
                formula = _matrix_value(formulas, r, c)
                if ((isinstance(value, str) and '#REF!' in value.upper()) or
                        (isinstance(formula, str) and ('#REF!' in formula.upper() or '[' in formula))):
                    errors.append(f"{ws.Name}!{ws.Cells(row+r, col+c).Address}: {formula or value}")
    return errors


def refresh(config: UpdateConfig, plan: UpdatePlan, progress=None):
    """Create and verify a separate workbook containing the corrected sheets."""
    output = config.output_workbook
    output.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output.with_name(f".{output.stem}.candidate-{os.getpid()}.xlsx")
    if temp_path.exists():
        temp_path.unlink()
    shutil.copy2(config.final_workbook, temp_path)
    report = output.with_name(f"{output.stem}.update_report.json")
    warnings = []
    try:
        with ExcelSession() as session:
            destination = session.open_workbook(temp_path, read_only=False)
            baseline = _formula_snapshot(destination)
            final_info = inspect_workbook(temp_path)
            for country in plan.countries:
                source = session.open_workbook(country.update.file, read_only=True)
                try:
                    _rename_source_sheets(session, source, config.sheet_mapping, country.update.prefix)
                    _validate_source_formulas(source, destination, final_info, {})
                    for sheet in country.sheets:
                        source_ws = source.Worksheets(sheet.target_name)
                        target_ws = destination.Worksheets(sheet.target_name)
                        if progress:
                            progress(f"Updating {country.update.country}: {sheet.target_name}")
                        _clear_and_transfer(session, source_ws, target_ws, config.sheet_password)
                finally:
                    session.close_workbook(source)
            session.call(lambda: session.app.CalculateFullRebuild())
            session.call(lambda: destination.Save())
            errors = _check_ref_errors(destination)
            if errors:
                raise ValueError('Broken or external references found after recalculation: ' + '; '.join(errors[:10]))
            updated = _formula_snapshot(destination)
            for key, formula in baseline.items():
                if key[0] not in {sheet.target_name for country in plan.countries for sheet in country.sheets} and updated.get(key) != formula:
                    raise ValueError(f"Unrelated formula changed: {key[0]}!{key[1]}")
            session.close_workbook(destination)
            verified = session.open_workbook(temp_path, read_only=True)
            try:
                errors = _check_ref_errors(verified)
                if errors:
                    raise ValueError('Broken or external references found after reopening: ' + '; '.join(errors[:10]))
            finally:
                session.close_workbook(verified)
        if output.exists() and not config.overwrite_output:
            raise FileExistsError(f"Output already exists: {output}")
        os.replace(temp_path, output)
        warnings.append('FILLING_RATE, VALIDATION and PPT_VALIDATION are generation snapshots; they were not refreshed.')
        result = {'output_workbook': str(output), 'report_file': str(report), 'warnings': warnings,
                  'countries': [country.update.country for country in plan.countries]}
        report.write_text(json.dumps(result, indent=2), encoding='utf-8')
        return result
    except Exception:
        if temp_path.exists():
            temp_path.unlink()
        raise
