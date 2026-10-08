"""Strict configuration for selective country-sheet updates.

Loading configuration never opens Excel or modifies a workbook. Paths are
resolved against the configuration file's directory, independently of the
process working directory.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml
from yaml.constructor import ConstructorError


class ConfigError(ValueError):
    """The update configuration is invalid or its required files are unavailable."""


@dataclass(frozen=True)
class SheetUpdate:
    source_name: str
    target_name: str


@dataclass(frozen=True)
class CountryUpdate:
    country: str
    prefix: str
    file: Path
    sheets: tuple[SheetUpdate, ...]


@dataclass(frozen=True)
class UpdateConfig:
    final_workbook: Path
    output_workbook: Path
    overwrite_output: bool
    check_text_layout: bool
    sheet_password: str
    updates: tuple[CountryUpdate, ...]
    sheet_mapping: tuple[tuple[str, str], ...]


class _UniqueKeyLoader(yaml.SafeLoader):
    """Reject duplicate mapping keys, including nested or merged mappings."""

    def construct_mapping(self, node: yaml.MappingNode, deep: bool = False) -> dict:
        if not isinstance(node, yaml.MappingNode):
            raise ConstructorError(
                None, None, "Expected a YAML mapping", node.start_mark
            )
        self.flatten_mapping(node)
        mapping = {}
        for key_node, value_node in node.value:
            key = self.construct_object(key_node, deep=deep)
            try:
                duplicate = key in mapping
            except TypeError as exc:
                raise ConstructorError(
                    None, None, "YAML mapping keys must be scalar", key_node.start_mark
                ) from exc
            if duplicate:
                raise ConstructorError(
                    None,
                    None,
                    f"Duplicate YAML key {key!r}",
                    key_node.start_mark,
                )
            mapping[key] = self.construct_object(value_node, deep=deep)
        return mapping


_TOP_KEYS = {
    "version", "final_workbook", "output_workbook", "overwrite_output",
    "check_text_layout", "sheet_password", "country_prefixes", "sheet_mapping",
    "updates",
}
_UPDATE_KEYS = {"country", "file", "sheets"}
_INVALID_SHEET_CHARS = set("\\/*?:[]")


def _mapping(value: Any, field: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ConfigError(f"{field} must be a mapping")
    for key in value:
        if not isinstance(key, str) or not key.strip():
            raise ConfigError(f"{field} must contain nonempty string keys")
        if key != key.strip():
            raise ConfigError(f"{field} keys must not have surrounding whitespace")
    return value


def _keys(value: dict[str, Any], allowed: set[str], field: str) -> None:
    unknown = set(value) - allowed
    if unknown:
        raise ConfigError(f"Unknown {field} key(s): {', '.join(sorted(unknown))}")


def _text(value: Any, field: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ConfigError(f"{field} must be a nonempty string")
    if value != value.strip():
        raise ConfigError(f"{field} must not have surrounding whitespace")
    return value


def _boolean(value: Any, field: str) -> bool:
    if type(value) is not bool:
        raise ConfigError(f"{field} must be true or false")
    return value


def _sheet_name(value: Any, field: str) -> str:
    name = _text(value, field)
    if len(name) > 31:
        raise ConfigError(f"{field} exceeds Excel's 31-character worksheet-name limit")
    if any(char in _INVALID_SHEET_CHARS or ord(char) < 32 for char in name):
        raise ConfigError(f"{field} contains an invalid worksheet-name character")
    if name.startswith("'") or name.endswith("'"):
        raise ConfigError(f"{field} must not begin or end with an apostrophe")
    return name


def _workbook_path(value: Any, field: str, directory: Path) -> Path:
    raw = _text(value, field)
    try:
        path = Path(raw).expanduser()
        if not path.is_absolute():
            path = directory / path
        path = path.resolve()
    except (OSError, RuntimeError, ValueError) as exc:
        raise ConfigError(f"{field} is not a valid path") from exc
    if path.suffix.lower() != ".xlsx":
        raise ConfigError(f"{field} must name an .xlsx workbook")
    return path


def _path_key(path: Path) -> str:
    return os.path.normcase(str(path.resolve()))


def load_config(path: str | Path, check_files: bool = False) -> UpdateConfig:
    """Load and validate one YAML update plan.

    Schema, worksheet names, selected destinations, extensions and resolved
    path collisions are always checked. ``check_files=True`` additionally
    requires all input workbooks to exist as files and enforces the output
    overwrite setting. No directories or workbooks are created by this loader.

    Optional settings default to ``overwrite_output=False``,
    ``check_text_layout=True`` and ``sheet_password='a'``. Omitting an update's
    ``sheets`` selects every source in ``sheet_mapping`` in mapping order.
    """
    config_path = Path(path).expanduser().resolve()
    try:
        with config_path.open("r", encoding="utf-8-sig") as handle:
            raw = yaml.load(handle, Loader=_UniqueKeyLoader)
    except (OSError, UnicodeError, yaml.YAMLError) as exc:
        raise ConfigError(f"Cannot load configuration {config_path}: {exc}") from exc

    data = _mapping(raw, "configuration")
    _keys(data, _TOP_KEYS, "configuration")
    if type(data.get("version")) is not int or data["version"] != 1:
        raise ConfigError("version must be the integer 1")

    final_workbook = _workbook_path(
        data.get("final_workbook"), "final_workbook", config_path.parent
    )
    output_workbook = _workbook_path(
        data.get("output_workbook"), "output_workbook", config_path.parent
    )
    overwrite_output = _boolean(data.get("overwrite_output", False), "overwrite_output")
    check_text_layout = _boolean(data.get("check_text_layout", True), "check_text_layout")
    sheet_password = data.get("sheet_password", "a")
    if not isinstance(sheet_password, str):
        raise ConfigError("sheet_password must be a string")

    countries_raw = _mapping(data.get("country_prefixes"), "country_prefixes")
    if not countries_raw:
        raise ConfigError("country_prefixes must not be empty")
    countries = {
        country: _sheet_name(prefix, f"country_prefixes[{country!r}]")
        for country, prefix in countries_raw.items()
    }
    mapping_raw = _mapping(data.get("sheet_mapping"), "sheet_mapping")
    if not mapping_raw:
        raise ConfigError("sheet_mapping must not be empty")
    sheet_mapping = {}
    source_keys = set()
    for source, suffix in mapping_raw.items():
        _sheet_name(source, f"sheet_mapping source {source!r}")
        if source.casefold() in source_keys:
            raise ConfigError("sheet_mapping source names must be unique ignoring case")
        source_keys.add(source.casefold())
        sheet_mapping[source] = _sheet_name(suffix, f"sheet_mapping[{source!r}]")

    updates_raw = data.get("updates")
    if not isinstance(updates_raw, list) or not updates_raw:
        raise ConfigError("updates must be a nonempty list")
    final_key, output_key = _path_key(final_workbook), _path_key(output_workbook)
    if final_key == output_key:
        raise ConfigError("output_workbook must differ from final_workbook")
    prefixes, source_files, destinations = set(), set(), set()
    updates = []
    for index, raw_update in enumerate(updates_raw, start=1):
        field = f"updates[{index}]"
        update = _mapping(raw_update, field)
        _keys(update, _UPDATE_KEYS, field)
        country = _text(update.get("country"), f"{field}.country")
        if country not in countries:
            raise ConfigError(f"{field}.country is not in country_prefixes: {country!r}")
        prefix = countries[country]
        if prefix.casefold() in prefixes:
            raise ConfigError(f"Duplicate update country prefix: {prefix!r}")
        prefixes.add(prefix.casefold())

        source_file = _workbook_path(update.get("file"), f"{field}.file", config_path.parent)
        source_key = _path_key(source_file)
        if source_key in {final_key, output_key}:
            raise ConfigError(f"{field}.file must differ from final_workbook and output_workbook")
        if source_key in source_files:
            raise ConfigError(f"{field}.file duplicates another update source file")
        source_files.add(source_key)

        selected = update.get("sheets", list(sheet_mapping))
        if not isinstance(selected, list) or not selected:
            raise ConfigError(f"{field}.sheets must be a nonempty list when supplied")
        selected_keys = set()
        sheets = []
        for sheet in selected:
            source_name = _text(sheet, f"{field}.sheets entry")
            if source_name not in sheet_mapping:
                raise ConfigError(f"Unknown {field}.sheets entry: {source_name!r}")
            if source_name.casefold() in selected_keys:
                raise ConfigError(f"Duplicate {field}.sheets entry: {source_name!r}")
            selected_keys.add(source_name.casefold())
            target_name = _sheet_name(
                f"{prefix}_{sheet_mapping[source_name]}", f"destination for {country!r}/{source_name!r}"
            )
            if target_name.casefold() in destinations:
                raise ConfigError(f"Duplicate destination worksheet: {target_name!r}")
            destinations.add(target_name.casefold())
            sheets.append(SheetUpdate(source_name=source_name, target_name=target_name))
        updates.append(CountryUpdate(country, prefix, source_file, tuple(sheets)))

    if check_files:
        for field, workbook in [("final_workbook", final_workbook)] + [
            (f"source for {update.country!r}", update.file) for update in updates
        ]:
            if not workbook.is_file():
                raise ConfigError(f"{field} does not exist as a file: {workbook}")
        if output_workbook.exists():
            if not output_workbook.is_file():
                raise ConfigError(f"output_workbook is not a file: {output_workbook}")
            if not overwrite_output:
                raise ConfigError(
                    "output_workbook already exists; choose another path or set overwrite_output: true"
                )

    return UpdateConfig(
        final_workbook, output_workbook, overwrite_output, check_text_layout,
        sheet_password, tuple(updates), tuple(sheet_mapping.items()),
    )
