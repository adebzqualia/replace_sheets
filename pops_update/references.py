"""Check formula bindings without calculating or rewriting business formulas."""
from __future__ import annotations

import re
from openpyxl.formula import Tokenizer


class ReferenceError(ValueError):
    """A country formula cannot safely bind to the retained consolidated sheets."""


_CELL = r'\$?[A-Za-z]{1,3}\$?[1-9][0-9]*'
_ADDRESS = re.compile(rf'^(?:{_CELL}(?::{_CELL})?|\$?[A-Za-z]{{1,3}}:\$?[A-Za-z]{{1,3}}|\$?[1-9][0-9]*:\$?[1-9][0-9]*)$')


def unquote_sheet(value: str) -> str:
    if value.startswith("'") and value.endswith("'"):
        return value[1:-1].replace("''", "'")
    return value


def canonical_formula(formula: str) -> tuple:
    """Ignore harmless Excel function prefixes, whitespace and reference casing."""
    result = []
    for token in Tokenizer(formula).items:
        if token.type in {'WSPACE', 'WHITE-SPACE'}:
            continue
        value = token.value
        if token.type == 'FUNC':
            value = value.replace('_xlfn.', '').replace('_xlws.', '').upper()
        elif token.type == 'OPERAND' and token.subtype == 'RANGE':
            if '!' in value:
                sheet, address = value.rsplit('!', 1)
                value = unquote_sheet(sheet).casefold() + '!' + address.casefold()
            else:
                value = value.casefold()
        result.append((token.type, token.subtype, value))
    return tuple(result)


def validate_formula(formula: str, *, sheet_name: str, allowed_sheets: set[str],
                     allowed_tables: set[str], source_names: dict[str, str],
                     destination_names: dict[str, str], in_table: bool = False,
                     _seen: frozenset[str] = frozenset()) -> None:
    """Require local, unambiguous bindings after Excel has renamed source sheets.

    Native Excel performs sheet/table reference changes on the unsaved source.
    This function only validates the resulting expression; it does not edit it.
    Name keys are casefolded, with Sheet!Name keys for worksheet-local names.
    """
    try:
        tokens = Tokenizer(formula).items
    except Exception as exc:
        raise ReferenceError(f'Cannot inspect formula {formula!r}: {exc}') from exc
    sheets = {s.casefold() for s in allowed_sheets}
    tables = {s.casefold() for s in allowed_tables}

    def check_name(name: str, scope: str) -> None:
        local = (scope + '!' + name).casefold()
        key = local if local in source_names else name.casefold()
        if key not in source_names:
            raise ReferenceError(f'Unknown name or reference {name!r} in {formula}')
        if key not in destination_names:
            raise ReferenceError(f'Destination is missing defined name {key!r}')
        if canonical_formula(source_names[key]) != canonical_formula(destination_names[key]):
            raise ReferenceError(f'Defined name {key!r} has a different destination definition; regenerate or align names first')
        if key in _seen:
            raise ReferenceError(f'Recursive defined name {key!r} is not supported')
        validate_formula(source_names[key], sheet_name=scope, allowed_sheets=allowed_sheets,
                         allowed_tables=allowed_tables, source_names=source_names,
                         destination_names=destination_names, in_table=in_table,
                         _seen=_seen | {key})

    for token in tokens:
        if token.type == 'FUNC' and token.subtype == 'OPEN':
            function = token.value.rstrip('(').replace('_xlfn.', '').upper()
            if function == 'INDIRECT':
                raise ReferenceError('INDIRECT is unsupported: text-based references do not follow native sheet renaming')
        if token.type == 'OPERAND' and token.subtype == 'ERROR' and token.value == '#REF!':
            raise ReferenceError(f'Broken reference in imported formula: {formula}')
        if token.type != 'OPERAND' or token.subtype != 'RANGE':
            continue
        ref = token.value
        scope = sheet_name
        if '!' in ref:
            prefix, ref = ref.rsplit('!', 1)
            scope = unquote_sheet(prefix)
            if '[' in scope or ']' in scope:
                raise ReferenceError(f'External workbook reference is unsupported: {token.value}')
            if ':' in scope:
                raise ReferenceError(f'3-D sheet references are unsupported: {token.value}')
            if scope.casefold() not in sheets:
                raise ReferenceError(f'Formula references an unmapped country/helper sheet: {scope!r}')
        if _ADDRESS.fullmatch(ref):
            continue
        if '[' in ref:
            table, _ = ref.split('[', 1)
            if table:
                if table.casefold() not in tables:
                    raise ReferenceError(f'Formula references an unmapped table: {table!r}')
            elif not in_table or not ref.endswith(']'):
                raise ReferenceError(f'Unresolved structured/external reference: {ref!r}')
            continue
        check_name(ref, scope)


def value_for_excel(value):
    """Prevent literal text from becoming a formula or numeric value on assignment."""
    if isinstance(value, str):
        return "'" + value
    if isinstance(value, (tuple, list)):
        return tuple(value_for_excel(item) for item in value)
    return value
