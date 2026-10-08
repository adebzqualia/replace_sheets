"""Portable configuration tests; no Excel installation or workbooks required."""

from __future__ import annotations

import copy
import os
import tempfile
import unittest
from dataclasses import FrozenInstanceError
from pathlib import Path

import yaml

from pops_update.config import ConfigError, load_config


class ConfigTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.path = self.root / "config" / "update.yaml"
        self.path.parent.mkdir()
        self.data = {
            "version": 1,
            "final_workbook": "../data/final.xlsx",
            "output_workbook": "../output/updated.xlsx",
            "overwrite_output": False,
            "check_text_layout": True,
            "sheet_password": "a",
            "country_prefixes": {
                "Germany": "Germany", "Belgium": "Belgium",
                "Findomestic": "Italy", "Italy": "Italy",
            },
            "sheet_mapping": {"ID Card": "ID Card", "FTE": "FTE"},
            "updates": [
                {"country": "Germany", "file": "../data/germany.xlsx", "sheets": ["FTE"]},
            ],
        }

    def write(self, data=None):
        self.path.write_text(
            yaml.safe_dump(self.data if data is None else data, sort_keys=False),
            encoding="utf-8",
        )
        return self.path

    def load(self, data=None, *, check_files=False):
        return load_config(self.write(data), check_files=check_files)

    def create_inputs(self):
        config = self.load()
        for path in [config.final_workbook] + [update.file for update in config.updates]:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(b"existence-only fixture")
        return config

    def add_belgium(self, **overrides):
        update = {"country": "Belgium", "file": "../data/belgium.xlsx"}
        update.update(overrides)
        self.data["updates"].append(update)

    def test_relative_paths_selection_order_and_immutable_results(self):
        self.data["updates"][0]["sheets"] = ["FTE", "ID Card"]
        config = self.load()
        self.assertEqual(config.final_workbook, (self.root / "data/final.xlsx").resolve())
        self.assertEqual(config.output_workbook, (self.root / "output/updated.xlsx").resolve())
        update = config.updates[0]
        self.assertEqual(update.file, (self.root / "data/germany.xlsx").resolve())
        self.assertEqual(update.country, "Germany")
        self.assertEqual(update.prefix, "Germany")
        self.assertEqual(
            [(sheet.source_name, sheet.target_name) for sheet in update.sheets],
            [("FTE", "Germany_FTE"), ("ID Card", "Germany_ID Card")],
        )
        self.assertIsInstance(config.updates, tuple)
        self.assertIsInstance(update.sheets, tuple)
        with self.assertRaises(FrozenInstanceError):
            update.prefix = "Other"
        with self.assertRaises(FrozenInstanceError):
            update.sheets[0].target_name = "Other"
        with self.assertRaises(FrozenInstanceError):
            config.overwrite_output = True

    def test_omitted_selection_uses_global_mapping_order(self):
        del self.data["updates"][0]["sheets"]
        config = self.load()
        self.assertEqual(
            [sheet.source_name for sheet in config.updates[0].sheets], ["ID Card", "FTE"]
        )

    def test_global_sheet_mapping_retains_unselected_sources_and_order(self):
        self.data["sheet_mapping"]["GRANTING 2027-2030"] = "GRANTING"
        config = self.load()
        self.assertEqual(
            config.sheet_mapping,
            (("ID Card", "ID Card"), ("FTE", "FTE"), ("GRANTING 2027-2030", "GRANTING")),
        )
        self.assertEqual([sheet.source_name for sheet in config.updates[0].sheets], ["FTE"])
        self.assertIsInstance(config.sheet_mapping, tuple)
        self.assertTrue(all(isinstance(pair, tuple) for pair in config.sheet_mapping))
        with self.assertRaises(FrozenInstanceError):
            config.sheet_mapping = ()

    def test_optional_settings_defaults_and_empty_password(self):
        for key in ("overwrite_output", "check_text_layout", "sheet_password"):
            del self.data[key]
        config = self.load()
        self.assertFalse(config.overwrite_output)
        self.assertTrue(config.check_text_layout)
        self.assertEqual(config.sheet_password, "a")
        self.data["sheet_password"] = ""
        self.assertEqual(self.load().sheet_password, "")

    def test_paths_resolve_from_config_location_not_working_directory(self):
        self.write()
        previous = Path.cwd()
        other = self.root / "elsewhere"
        other.mkdir()
        try:
            os.chdir(other)
            config = load_config(self.path)
        finally:
            os.chdir(previous)
        self.assertEqual(config.final_workbook, (self.root / "data/final.xlsx").resolve())

    def test_absolute_paths_are_retained(self):
        absolute = (self.root / "elsewhere/final.xlsx").resolve()
        self.data["final_workbook"] = str(absolute)
        self.assertEqual(self.load().final_workbook, absolute)

    def test_version_rejects_bool_float_string_and_other_values(self):
        for value in (True, False, 1.0, "1", 2, None):
            with self.subTest(value=value):
                self.data["version"] = value
                with self.assertRaisesRegex(ConfigError, "version"):
                    self.load()

    def test_boolean_settings_are_strict(self):
        for key in ("overwrite_output", "check_text_layout"):
            for value in (1, 0, "true", "false", None, []):
                with self.subTest(key=key, value=value):
                    data = copy.deepcopy(self.data)
                    data[key] = value
                    with self.assertRaisesRegex(ConfigError, key):
                        self.load(data)

    def test_password_must_be_string(self):
        for value in (None, False, 123, [], {}):
            with self.subTest(value=value):
                self.data["sheet_password"] = value
                with self.assertRaisesRegex(ConfigError, "sheet_password"):
                    self.load()

    def test_unknown_keys_rejected_at_top_and_update_levels(self):
        self.data["unexpected"] = True
        with self.assertRaisesRegex(ConfigError, "Unknown configuration"):
            self.load()
        del self.data["unexpected"]
        self.data["updates"][0]["worksheet"] = "FTE"
        with self.assertRaisesRegex(ConfigError, "Unknown updates"):
            self.load()

    def test_yaml_duplicate_keys_rejected_including_nested_maps(self):
        valid = yaml.safe_dump(self.data, sort_keys=False)
        cases = {
            "root": valid + "version: 1\n",
            "countries": valid.replace("country_prefixes:\n", "country_prefixes:\n  Germany: Germany\n"),
            "sheet_mapping": valid.replace("sheet_mapping:\n", "sheet_mapping:\n  FTE: FTE\n"),
            "update": valid.replace("- country: Germany\n", "- country: Germany\n  country: Belgium\n"),
        }
        for name, text in cases.items():
            with self.subTest(name=name):
                self.path.write_text(text, encoding="utf-8")
                with self.assertRaisesRegex(ConfigError, "Duplicate YAML key"):
                    load_config(self.path)

    def test_merged_mapping_cannot_hide_duplicate_keys(self):
        text = yaml.safe_dump(self.data, sort_keys=False)
        text = text.replace(
            "- country: Germany\n",
            "- <<: {country: Germany}\n  country: Belgium\n",
        )
        self.path.write_text(text, encoding="utf-8")
        with self.assertRaisesRegex(ConfigError, "Duplicate YAML key"):
            load_config(self.path)

    def test_invalid_yaml_and_missing_configuration_are_config_errors(self):
        with self.assertRaises(ConfigError):
            load_config(self.path)
        for text in ("{broken", "[]", "", "!!python/object:object {}"):
            with self.subTest(text=text):
                self.path.write_text(text, encoding="utf-8")
                with self.assertRaises(ConfigError):
                    load_config(self.path)

    def test_required_structure_cannot_be_missing_or_empty(self):
        for key in ("version", "final_workbook", "output_workbook", "country_prefixes", "sheet_mapping", "updates"):
            with self.subTest(key=key):
                data = copy.deepcopy(self.data)
                del data[key]
                with self.assertRaises(ConfigError):
                    self.load(data)
        for key, value in (("country_prefixes", {}), ("sheet_mapping", {}), ("updates", []), ("updates", {})):
            with self.subTest(key=key, value=value):
                data = copy.deepcopy(self.data)
                data[key] = value
                with self.assertRaises(ConfigError):
                    self.load(data)

    def test_invalid_selections_rejected(self):
        for selection in ([], None, "FTE", ["Unknown"], ["FTE", "FTE"], [True], [""]):
            with self.subTest(selection=selection):
                self.data["updates"][0]["sheets"] = selection
                with self.assertRaises(ConfigError):
                    self.load()

    def test_unknown_country_rejected(self):
        self.data["updates"][0]["country"] = "Unknown"
        with self.assertRaisesRegex(ConfigError, "country_prefixes"):
            self.load()

    def test_aliases_can_exist_but_cannot_update_same_prefix_twice(self):
        self.data["updates"][0]["country"] = "Findomestic"
        self.assertEqual(self.load().updates[0].prefix, "Italy")
        self.data["updates"].append({"country": "Italy", "file": "../data/italy.xlsx"})
        with self.assertRaisesRegex(ConfigError, "Duplicate update country prefix"):
            self.load()

    def test_prefix_uniqueness_ignores_case(self):
        self.data["country_prefixes"]["Belgium"] = "germany"
        self.add_belgium()
        with self.assertRaisesRegex(ConfigError, "Duplicate update country prefix"):
            self.load()

    def test_duplicate_source_files_rejected_after_path_resolution(self):
        self.add_belgium(file="../data/sub/../germany.xlsx")
        with self.assertRaisesRegex(ConfigError, "duplicates another update source"):
            self.load()

    def test_input_output_and_source_paths_must_be_distinct_even_without_file_checks(self):
        self.data["overwrite_output"] = True
        cases = (
            ("output_workbook", "../data/sub/../final.xlsx"),
            ("source", "../data/final.xlsx"),
            ("source", "../output/updated.xlsx"),
        )
        for key, value in cases:
            with self.subTest(key=key, value=value):
                data = copy.deepcopy(self.data)
                if key == "source":
                    data["updates"][0]["file"] = value
                else:
                    data[key] = value
                with self.assertRaisesRegex(ConfigError, "must differ"):
                    self.load(data)

    @unittest.skipUnless(os.path.normcase("A") == os.path.normcase("a"), "Host paths are case-sensitive")
    def test_path_collisions_follow_host_normcase(self):
        self.data["output_workbook"] = "../DATA/FINAL.XLSX"
        with self.assertRaisesRegex(ConfigError, "must differ"):
            self.load()

    def test_all_workbook_paths_require_xlsx_regardless_of_file_checks(self):
        for field in ("final_workbook", "output_workbook", "source"):
            with self.subTest(field=field):
                data = copy.deepcopy(self.data)
                if field == "source":
                    data["updates"][0]["file"] = "../data/source.xlsm"
                else:
                    data[field] = "../data/unsupported.xlsm"
                with self.assertRaisesRegex(ConfigError, "xlsx"):
                    self.load(data)
        self.data["final_workbook"] = "../data/final.XLSX"
        self.assertEqual(self.load().final_workbook.suffix, ".XLSX")

    def test_duplicate_destination_names_rejected_ignoring_case(self):
        self.data["sheet_mapping"]["ID Card"] = "fte"
        del self.data["updates"][0]["sheets"]
        with self.assertRaisesRegex(ConfigError, "Duplicate destination"):
            self.load()

    def test_source_sheet_names_unique_ignoring_case(self):
        self.data["sheet_mapping"]["fte"] = "Alternative"
        with self.assertRaisesRegex(ConfigError, "unique ignoring case"):
            self.load()

    def test_excel_name_length_checks_full_destination(self):
        self.data["country_prefixes"]["Germany"] = "G"
        self.data["sheet_mapping"]["FTE"] = "x" * 29
        self.assertEqual(len(self.load().updates[0].sheets[0].target_name), 31)
        self.data["sheet_mapping"]["FTE"] = "x" * 30
        with self.assertRaisesRegex(ConfigError, "31-character"):
            self.load()
        self.data["sheet_mapping"] = {"s" * 32: "FTE"}
        with self.assertRaisesRegex(ConfigError, "31-character"):
            self.load()

    def test_invalid_sheet_characters_and_blank_names_rejected(self):
        for name in ("", " ", "FTE/Other", "FTE\\Other", "FTE?", "FTE*", "FTE:", "[FTE]", "FTE\x01", "'FTE", "FTE'"):
            with self.subTest(name=name):
                data = copy.deepcopy(self.data)
                data["sheet_mapping"]["FTE"] = name
                with self.assertRaises(ConfigError):
                    self.load(data)

    def test_nonstring_mapping_keys_and_values_rejected(self):
        for field in ("country_prefixes", "sheet_mapping"):
            for mapping in ({1: "FTE"}, {"FTE": True}, {"": "FTE"}):
                with self.subTest(field=field, mapping=mapping):
                    data = copy.deepcopy(self.data)
                    data[field] = mapping
                    with self.assertRaises(ConfigError):
                        self.load(data)

    def test_file_checks_are_opt_in_and_do_not_create_directories(self):
        self.load()
        self.assertFalse((self.root / "data").exists())
        self.assertFalse((self.root / "output").exists())
        with self.assertRaisesRegex(ConfigError, "final_workbook does not exist"):
            self.load(check_files=True)
        config = self.create_inputs()
        self.assertEqual(self.load(check_files=True).final_workbook, config.final_workbook)
        config.updates[0].file.unlink()
        with self.assertRaisesRegex(ConfigError, "source for .* does not exist"):
            self.load(check_files=True)

    def test_existing_output_requires_explicit_overwrite(self):
        config = self.create_inputs()
        config.output_workbook.parent.mkdir()
        config.output_workbook.write_bytes(b"keep output intact")
        self.load(check_files=False)
        with self.assertRaisesRegex(ConfigError, "already exists"):
            self.load(check_files=True)
        self.data["overwrite_output"] = True
        self.load(check_files=True)
        self.assertEqual(config.output_workbook.read_bytes(), b"keep output intact")

    def test_directories_are_not_accepted_as_workbook_files(self):
        config = self.create_inputs()
        config.output_workbook.mkdir(parents=True)
        self.data["overwrite_output"] = True
        with self.assertRaisesRegex(ConfigError, "output_workbook is not a file"):
            self.load(check_files=True)
        config.output_workbook.rmdir()
        config.final_workbook.unlink()
        config.final_workbook.mkdir()
        with self.assertRaisesRegex(ConfigError, "final_workbook does not exist as a file"):
            self.load(check_files=True)

    def test_shipped_configuration_has_original_map_and_explicit_selections(self):
        example = Path(__file__).resolve().parents[1] / "config/update.yaml"
        config = load_config(example)
        self.assertEqual([update.country for update in config.updates], ["Germany", "Belgium"])
        self.assertTrue(all([sheet.source_name for sheet in update.sheets] == ["ID Card", "FTE"] for update in config.updates))
        raw = yaml.safe_load(example.read_text(encoding="utf-8"))
        self.assertEqual(len(raw["sheet_mapping"]), 13)
        self.assertEqual(config.sheet_mapping, tuple(raw["sheet_mapping"].items()))
        self.assertEqual(raw["country_prefixes"]["Findomestic"], "Italy")
        self.assertEqual(raw["country_prefixes"]["Italy"], "Italy")
        self.assertEqual(raw["country_prefixes"]["France"], "PFF")


if __name__ == "__main__":
    unittest.main()
