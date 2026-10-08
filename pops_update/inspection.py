"""Read workbook structure without opening Excel or rewriting the package.

The only openpyxl utility used is the shared-formula reference translator; no
openpyxl workbook is loaded or saved. Inspection does not calculate formulas.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import posixpath
import re
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

from openpyxl.formula.translate import Translator, TranslatorError


@dataclass(frozen=True)
class TableInfo:
    name: str
    ref: str
    columns: tuple[str, ...]


@dataclass(frozen=True)
class SheetInfo:
    title: str
    text_labels: dict[str, str] = field(default_factory=dict)
    formulas: dict[str, str] = field(default_factory=dict)
    merged_ranges: tuple[str, ...] = ()
    array_ranges: tuple[str, ...] = ()
    tables: tuple[TableInfo, ...] = ()
    error_cells: dict[str, str] = field(default_factory=dict)


@dataclass(frozen=True)
class WorkbookInfo:
    sheet_names: tuple[str, ...]
    sheets: dict[str, SheetInfo]
    names: dict[str, str] = field(default_factory=dict)


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _child(element: ET.Element, name: str) -> ET.Element | None:
    return next((item for item in element if _local(item.tag) == name), None)


def _xml(archive: ZipFile, part: str) -> ET.Element:
    try:
        return ET.fromstring(archive.read(part))
    except KeyError as exc:
        raise ValueError(f"Required workbook part is missing: {part}") from exc
    except ET.ParseError as exc:
        raise ValueError(f"Malformed XML in workbook part {part}: {exc}") from exc


def _resolve_part(source: str, target: str) -> str:
    """Resolve an internal OPC relationship without allowing package escape."""
    if not target or "\\" in target:
        raise ValueError(f"Invalid relationship target: {target!r}")
    parsed = urlsplit(target)
    if parsed.scheme or parsed.netloc or parsed.query or parsed.fragment:
        raise ValueError(f"Expected an internal workbook relationship: {target!r}")
    decoded = unquote(parsed.path)
    if "\\" in decoded or "\x00" in decoded:
        raise ValueError(f"Invalid relationship target: {target!r}")
    base = "" if decoded.startswith("/") else posixpath.dirname(source)
    result: list[str] = []
    for component in (base + "/" + decoded.lstrip("/")).split("/"):
        if component in ("", "."):
            continue
        if component == "..":
            if not result:
                raise ValueError(f"Relationship escapes the workbook package: {target!r}")
            result.pop()
        else:
            result.append(component)
    if not result:
        raise ValueError(f"Relationship does not identify a workbook part: {target!r}")
    return "/".join(result)


def _relationships(archive: ZipFile, part: str) -> dict[str, dict[str, str]]:
    rel_part = (posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
                if part else "_rels/.rels")
    root = _xml(archive, rel_part)
    if _local(root.tag) != "Relationships":
        raise ValueError(f"Invalid relationships root in {rel_part}")
    result: dict[str, dict[str, str]] = {}
    for node in root:
        if _local(node.tag) != "Relationship":
            continue
        identifier = node.get("Id", "")
        if not identifier or identifier in result:
            raise ValueError(f"Missing or duplicate relationship Id in {rel_part}")
        if not node.get("Type") or not node.get("Target"):
            raise ValueError(f"Incomplete relationship {identifier!r} in {rel_part}")
        result[identifier] = dict(node.attrib)
    return result


def _relationship_part(source: str, relationship: dict[str, str], expected: str) -> str:
    if not relationship["Type"].endswith("/" + expected):
        raise ValueError(f"Expected a {expected} relationship in {source or 'package root'}")
    if relationship.get("TargetMode", "Internal") != "Internal":
        raise ValueError(f"External {expected} relationships are unsupported")
    return _resolve_part(source, relationship["Target"])


def _relationship_id(node: ET.Element) -> str:
    return next((value for key, value in node.attrib.items()
                 if key.startswith("{") and _local(key) == "id"), "")


def _address(value: str, context: str) -> str:
    match = re.fullmatch(r"([A-Za-z]{1,3})([1-9][0-9]*)", value)
    if match is None:
        raise ValueError(f"Invalid cell address {value!r} in {context}")
    column = 0
    for character in match.group(1).upper():
        column = column * 26 + ord(character) - ord("A") + 1
    if column > 16384 or int(match.group(2)) > 1048576:
        raise ValueError(f"Cell address exceeds Excel limits: {value!r} in {context}")
    return value.upper()


def _position(address: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]+)([0-9]+)", address)
    assert match is not None
    column = 0
    for character in match.group(1):
        column = column * 26 + ord(character) - ord("A") + 1
    return int(match.group(2)), column


def _range(value: str, context: str) -> str:
    parts = value.split(":")
    if len(parts) not in (1, 2):
        raise ValueError(f"Invalid cell range {value!r} in {context}")
    points = [_address(part, context) for part in parts]
    first, last = _position(points[0]), _position(points[-1])
    if first[0] > last[0] or first[1] > last[1]:
        raise ValueError(f"Reversed cell range {value!r} in {context}")
    return ":".join(points)


def _rich_text(node: ET.Element) -> str:
    """Read plain/rich string text, excluding phonetic annotation runs."""
    fragments: list[str] = []
    for child in node:
        if _local(child.tag) == "t":
            fragments.append(child.text or "")
        elif _local(child.tag) == "r":
            text = _child(child, "t")
            if text is not None:
                fragments.append(text.text or "")
    return "".join(fragments)


def _formula(text: str) -> str:
    return text if text.startswith("=") else "=" + text


def _table(archive: ZipFile, part: str) -> TableInfo:
    root = _xml(archive, part)
    if _local(root.tag) != "table":
        raise ValueError(f"Invalid table root in {part}")
    name = root.get("displayName") or root.get("name")
    if not name:
        raise ValueError(f"Missing table name in {part}")
    reference = _range(root.get("ref", ""), part)
    container = _child(root, "tableColumns")
    if container is None:
        raise ValueError(f"Missing table columns in {part}")
    columns = tuple(column.get("name", "") for column in container
                    if _local(column.tag) == "tableColumn")
    if not columns or any(not column for column in columns):
        raise ValueError(f"Missing table column names in {part}")
    points = reference.split(":")
    width = _position(points[-1])[1] - _position(points[0])[1] + 1
    if len(columns) != width:
        raise ValueError(f"Table column count does not match {reference} in {part}")
    return TableInfo(name=name, ref=reference, columns=columns)


def _sheet(archive: ZipFile, title: str, part: str, strings: list[str]) -> SheetInfo:
    root = _xml(archive, part)
    if _local(root.tag) != "worksheet":
        raise ValueError(f"Expected a worksheet XML part for {title!r}")
    labels: dict[str, str] = {}
    formulas: dict[str, str] = {}
    errors: dict[str, str] = {}
    arrays: set[str] = set()
    shared_masters: dict[str, tuple[str, str, str | None]] = {}
    shared_followers: list[tuple[str, str]] = []
    seen_cells: set[str] = set()
    sheet_data = _child(root, "sheetData")
    if sheet_data is None:
        raise ValueError(f"Missing sheetData in worksheet {title!r}")
    for row in sheet_data:
        if _local(row.tag) != "row":
            continue
        for cell in row:
            if _local(cell.tag) != "c":
                continue
            address = _address(cell.get("r", ""), title)
            if address in seen_cells:
                raise ValueError(f"Duplicate cell {title}!{address}")
            seen_cells.add(address)
            value_node, formula_node = _child(cell, "v"), _child(cell, "f")
            value = value_node.text if value_node is not None else None
            if cell.get("t") == "e":
                errors[address] = value or ""
            if formula_node is not None:
                kind = formula_node.get("t", "normal")
                content = formula_node.text or ""
                if kind == "array":
                    arrays.add(_range(formula_node.get("ref", address), title))
                elif kind == "dataTable":
                    raise ValueError(f"What-if data-table formulas are unsupported: {title}!{address}")
                elif kind not in ("normal", "shared"):
                    raise ValueError(f"Unsupported formula kind {kind!r}: {title}!{address}")
                if kind == "shared":
                    shared_id = formula_node.get("si", "")
                    if not re.fullmatch(r"[0-9]+", shared_id):
                        raise ValueError(f"Missing or invalid shared-formula index at {title}!{address}")
                    shared_id = str(int(shared_id))
                    if content:
                        if shared_id in shared_masters:
                            raise ValueError(f"Duplicate shared-formula master in {title}: {shared_id}")
                        reference = formula_node.get("ref")
                        shared_masters[shared_id] = (address, _formula(content),
                                                     _range(reference, title) if reference else None)
                        formulas[address] = _formula(content)
                    else:
                        shared_followers.append((address, shared_id))
                elif content:
                    formulas[address] = _formula(content)
                else:
                    raise ValueError(f"Missing formula text at {title}!{address}")
                # A cached string result is not a constant layout label.
                continue
            kind = cell.get("t")
            if kind == "s":
                if value is None or re.fullmatch(r"[0-9]+", value) is None or int(value) >= len(strings):
                    raise ValueError(f"Invalid shared-string index at {title}!{address}")
                labels[address] = strings[int(value)]
            elif kind == "inlineStr":
                inline = _child(cell, "is")
                if inline is None:
                    raise ValueError(f"Missing inline string at {title}!{address}")
                labels[address] = _rich_text(inline)
            elif kind == "str":
                labels[address] = value or ""
    for address, shared_id in shared_followers:
        if shared_id not in shared_masters:
            raise ValueError(f"Shared formula has no master at {title}!{address}")
        origin, master, reference = shared_masters[shared_id]
        if reference:
            points = reference.split(":")
            first, last, current = _position(points[0]), _position(points[-1]), _position(address)
            if not (first[0] <= current[0] <= last[0] and first[1] <= current[1] <= last[1]):
                raise ValueError(f"Shared-formula follower lies outside its range: {title}!{address}")
        try:
            formulas[address] = Translator(master, origin=origin).translate_formula(address)
        except (TranslatorError, ValueError, IndexError) as exc:
            raise ValueError(f"Cannot expand shared formula at {title}!{address}: {exc}") from exc
    merges_node = _child(root, "mergeCells")
    merges = tuple(sorted(_range(item.get("ref", ""), title) for item in merges_node
                          if _local(item.tag) == "mergeCell")) if merges_node is not None else ()
    tables: list[TableInfo] = []
    table_parts = _child(root, "tableParts")
    if table_parts is not None and len(table_parts):
        relationships = _relationships(archive, part)
        for table_part in table_parts:
            if _local(table_part.tag) != "tablePart":
                continue
            identifier = _relationship_id(table_part)
            if identifier not in relationships:
                raise ValueError(f"Missing table relationship {identifier!r} for {title!r}")
            target = _relationship_part(part, relationships[identifier], "table")
            tables.append(_table(archive, target))
    return SheetInfo(title, labels, formulas, merges, tuple(sorted(arrays)),
                     tuple(sorted(tables, key=lambda item: (item.ref, item.name))), errors)


def inspect_workbook(path: str | Path) -> WorkbookInfo:
    """Inspect an unencrypted .xlsx package; never calculate or modify files.

    Shared formula followers are expanded to ordinary A1 formula strings.
    Array ranges (including dynamic-array anchors) are recorded so that layout
    validation can explicitly reject them before a content refresh.
    """
    path = Path(path)
    if path.suffix.casefold() != ".xlsx":
        raise ValueError(f"Workbook inspection requires an .xlsx file: {path}")
    try:
        with ZipFile(path) as archive:
            parts = archive.namelist()
            if len(parts) != len(set(parts)):
                raise ValueError("Workbook package contains duplicate part names")
            roots = [relation for relation in _relationships(archive, "").values()
                     if relation["Type"].endswith("/officeDocument")]
            if len(roots) != 1:
                raise ValueError("Workbook package must have one officeDocument relationship")
            workbook_part = _relationship_part("", roots[0], "officeDocument")
            root = _xml(archive, workbook_part)
            if _local(root.tag) != "workbook":
                raise ValueError("The officeDocument relationship does not identify an Excel workbook")
            relationships = _relationships(archive, workbook_part)
            strings: list[str] = []
            string_relations = [relation for relation in relationships.values()
                                if relation["Type"].endswith("/sharedStrings")]
            if len(string_relations) > 1:
                raise ValueError("Workbook has multiple shared-string relationships")
            if string_relations:
                shared_root = _xml(archive, _relationship_part(workbook_part, string_relations[0], "sharedStrings"))
                if _local(shared_root.tag) != "sst":
                    raise ValueError("Invalid shared-string table root")
                strings = [_rich_text(item) for item in shared_root if _local(item.tag) == "si"]
            sheet_container = _child(root, "sheets")
            if sheet_container is None or not len(sheet_container):
                raise ValueError("Workbook does not contain any worksheets")
            sheets: dict[str, SheetInfo] = {}
            casefold_names: set[str] = set()
            worksheet_parts: set[str] = set()
            for entry in sheet_container:
                if _local(entry.tag) != "sheet":
                    continue
                name = entry.get("name", "")
                if not name or name.casefold() in casefold_names:
                    raise ValueError(f"Missing or duplicate worksheet name: {name!r}")
                identifier = _relationship_id(entry)
                if identifier not in relationships:
                    raise ValueError(f"Missing worksheet relationship for {name!r}")
                part = _relationship_part(workbook_part, relationships[identifier], "worksheet")
                if part in worksheet_parts:
                    raise ValueError(f"Several worksheet names use the same XML part: {part}")
                sheets[name] = _sheet(archive, name, part, strings)
                casefold_names.add(name.casefold())
                worksheet_parts.add(part)
            if not sheets:
                raise ValueError("Workbook does not contain any worksheets")
            sheet_names = tuple(sheets)
            names: dict[str, str] = {}
            defined_names = _child(root, "definedNames")
            if defined_names is not None:
                for node in defined_names:
                    if _local(node.tag) != "definedName":
                        continue
                    name = node.get("name", "")
                    if not name:
                        raise ValueError("Defined name has no name attribute")
                    local = node.get("localSheetId")
                    if local is not None:
                        if not re.fullmatch(r"[0-9]+", local) or int(local) >= len(sheet_names):
                            raise ValueError(f"Invalid worksheet scope for defined name {name!r}")
                        name = sheet_names[int(local)] + "!" + name
                    if name in names:
                        raise ValueError(f"Duplicate defined name: {name}")
                    names[name] = _formula(node.text or "")
            return WorkbookInfo(sheet_names, sheets, names)
    except (BadZipFile, OSError, RuntimeError) as exc:
        raise ValueError(f"Cannot inspect workbook {path}: {exc}") from exc


def check_layout(original: SheetInfo, updated: SheetInfo, check_text_layout: bool = True) -> None:
    """Require compatible cell positions, merges and existing table schemas.

    Titles and table names may differ because consolidation renames them.
    Text checking compares exact constant text and coordinates. Disable it only
    when that layout assurance is supplied by another explicit process.
    """
    if original.array_ranges or updated.array_ranges:
        ranges = sorted(set(original.array_ranges) | set(updated.array_ranges))
        raise ValueError(f"Array/spill formulas are unsupported on {original.title!r}: {', '.join(ranges[:10])}")
    if sorted(original.merged_ranges) != sorted(updated.merged_ranges):
        raise ValueError(f"Merged-cell layout differs on {original.title!r}")
    old_tables = sorted((table.ref, table.columns) for table in original.tables)
    new_tables = sorted((table.ref, table.columns) for table in updated.tables)
    if old_tables != new_tables:
        raise ValueError(f"Table layout or column schema differs on {original.title!r}")
    if check_text_layout and original.text_labels != updated.text_labels:
        missing = object()
        changed = sorted((address for address in original.text_labels.keys() | updated.text_labels.keys()
                          if original.text_labels.get(address, missing) != updated.text_labels.get(address, missing)),
                         key=_position)
        def describe(value: object) -> str:
            if value is missing:
                return "<absent>"
            text = repr(value)
            return text if len(text) <= 100 else text[:97] + "..."
        details = [f"{address}: {describe(original.text_labels.get(address, missing))} -> "
                   f"{describe(updated.text_labels.get(address, missing))}" for address in changed[:10]]
        remainder = f"; and {len(changed) - 10} more" if len(changed) > 10 else ""
        raise ValueError(f"Text layout differs on {original.title!r} ({len(changed)} cells): "
                         + "; ".join(details) + remainder)
