"""Read-only inspection and layout checks using small synthetic Open XML files."""
from __future__ import annotations

from dataclasses import replace
from hashlib import sha256
from pathlib import Path
import tempfile
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from pops_update.inspection import (
    SheetInfo, TableInfo, check_layout, inspect_workbook,
)


S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"


def worksheet(content: str = "", extra: str = "") -> str:
    return f'<worksheet xmlns="{S}" xmlns:r="{R}"><sheetData>{content}</sheetData>{extra}</worksheet>'


def relationships(entries: list[dict[str, str]]) -> str:
    root = ET.Element("Relationships", xmlns=P)
    for entry in entries:
        ET.SubElement(root, "Relationship", entry)
    return ET.tostring(root, encoding="unicode")


def base_parts(sheets: tuple[tuple[str, str], ...] | None = None) -> dict[str, str]:
    sheets = sheets or (("Country", worksheet()),)
    workbook = ET.Element("workbook", xmlns=S)
    sheet_list = ET.SubElement(workbook, "sheets")
    entries = []
    parts = {"_rels/.rels": relationships([
        {"Id": "doc", "Type": R + "/officeDocument", "Target": "xl/workbook.xml"},
    ])}
    for number, (title, content) in enumerate(sheets, 1):
        ET.SubElement(sheet_list, "sheet", name=title, sheetId=str(number),
                      attrib={f"{{{R}}}id": f"sheet{number}"})
        entries.append({"Id": f"sheet{number}", "Type": R + "/worksheet",
                        "Target": f"worksheets/sheet{number}.xml"})
        parts[f"xl/worksheets/sheet{number}.xml"] = content
    parts["xl/workbook.xml"] = ET.tostring(workbook, encoding="unicode")
    parts["xl/_rels/workbook.xml.rels"] = relationships(entries)
    return parts


class InspectionTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "fixture.xlsx"

    def write(self, parts: dict[str, str]) -> Path:
        with ZipFile(self.path, "w") as archive:
            for name, content in sorted(parts.items(), reverse=True):
                archive.writestr(name, content)
        return self.path

    def test_text_formula_error_maps_and_workbook_order_without_mutation(self):
        cells = (
            '<row r="1"><c r="A1" t="s"><v>0</v></c>'
            '<c r="B1" t="inlineStr"><is><r><t>Inline </t></r><r><t>label</t></r>'
            '<rPh sb="0" eb="1"><t>not-label-text</t></rPh></is></c>'
            '<c r="C1" t="str"><v>Text constant</v></c>'
            '<c r="D1" t="str"><f>IF(A1="","","cached")</f><v>cached</v></c>'
            '<c r="E1"><v>42</v></c><c r="F1" t="b"><v>1</v></c>'
            '<c r="G1" t="e"><v>#VALUE!</v></c></row>'
            '<row r="2"><c r="A2" t="e"><f>#REF!+1</f><v>#REF!</v></c>'
            '<c r="B2" t="inlineStr"><is><t></t></is></c></row>'
        )
        parts = base_parts((("Z country", worksheet(cells)), ("A report", worksheet())))
        relations = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
        ET.SubElement(relations, f"{{{P}}}Relationship", Id="strings", Type=R + "/sharedStrings",
                      Target="sharedStrings.xml")
        parts["xl/_rels/workbook.xml.rels"] = ET.tostring(relations, encoding="unicode")
        parts["xl/sharedStrings.xml"] = (
            f'<sst xmlns="{S}"><si><r><t>Shared </t></r><r><t>label</t></r>'
            '<rPh sb="0" eb="1"><t>phonetic</t></rPh></si></sst>'
        )
        path = self.write(parts)
        before_hash = sha256(path.read_bytes()).hexdigest()
        before_files = sorted(Path(self.directory.name).iterdir())
        info = inspect_workbook(path)
        self.assertEqual(info.sheet_names, ("Z country", "A report"))
        self.assertEqual(tuple(info.sheets), info.sheet_names)
        source = info.sheets["Z country"]
        self.assertEqual(source.text_labels, {
            "A1": "Shared label", "B1": "Inline label", "C1": "Text constant", "B2": "",
        })
        self.assertEqual(source.formulas, {"D1": '=IF(A1="","","cached")', "A2": "=#REF!+1"})
        self.assertEqual(source.error_cells, {"G1": "#VALUE!", "A2": "#REF!"})
        self.assertEqual(sha256(path.read_bytes()).hexdigest(), before_hash)
        self.assertEqual(sorted(Path(self.directory.name).iterdir()), before_files)

    def test_shared_formulas_expand_relative_references_not_string_literals(self):
        cells = (
            '<row r="3"><c r="D3"><f t="shared" si="4" ref="D3:D4">'
            '$A3+B$1+\'Some Sheet\'!C3+"A3"</f><v>1</v></c></row>'
            '<row r="4"><c r="D4"><f t="shared" si="4"/><v>2</v></c></row>'
        )
        info = inspect_workbook(self.write(base_parts((("Country", worksheet(cells)),))))
        self.assertEqual(info.sheets["Country"].formulas["D4"], '=$A4+B$1+\'Some Sheet\'!C4+"A3"')

    def test_shared_formula_follower_can_precede_master_in_xml(self):
        cells = (
            '<row r="2"><c r="B2"><f t="shared" si="0"/></c></row>'
            '<row r="1"><c r="B1"><f t="shared" si="0" ref="B1:B2">A1*2</f></c></row>'
        )
        info = inspect_workbook(self.write(base_parts((("Country", worksheet(cells)),))))
        self.assertEqual(info.sheets["Country"].formulas, {"B1": "=A1*2", "B2": "=A2*2"})

    def test_merges_arrays_tables_and_scoped_names(self):
        cells = '<row r="1"><c r="D1"><f t="array" ref="D1:D3">_xlfn.SEQUENCE(3)</f><v>1</v></c></row>'
        extras = (
            '<mergeCells><mergeCell ref="C8:D8"/><mergeCell ref="A1:B1"/></mergeCells>'
            '<tableParts count="1"><tablePart r:id="table1"/></tableParts>'
        )
        parts = base_parts((("Country", worksheet(cells, extras)),))
        parts["xl/worksheets/_rels/sheet1.xml.rels"] = relationships([
            {"Id": "table1", "Type": R + "/table", "Target": "../tables/table1.xml"},
        ])
        parts["xl/tables/table1.xml"] = (
            f'<table xmlns="{S}" name="KPITable" displayName="KPITable" ref="A4:B6">'
            '<tableColumns count="2"><tableColumn id="1" name="KPI"/>'
            '<tableColumn id="2" name="2026"/></tableColumns></table>'
        )
        book = ET.fromstring(parts["xl/workbook.xml"])
        names = ET.SubElement(book, f"{{{S}}}definedNames")
        ET.SubElement(names, f"{{{S}}}definedName", name="Global").text = "Country!$A$1"
        ET.SubElement(names, f"{{{S}}}definedName", name="Local", localSheetId="0").text = "Country!$B$1"
        parts["xl/workbook.xml"] = ET.tostring(book, encoding="unicode")
        info = inspect_workbook(self.write(parts))
        sheet = info.sheets["Country"]
        self.assertEqual(sheet.merged_ranges, ("A1:B1", "C8:D8"))
        self.assertEqual(sheet.array_ranges, ("D1:D3",))
        self.assertEqual(sheet.tables, (TableInfo("KPITable", "A4:B6", ("KPI", "2026")),))
        self.assertEqual(info.names, {"Global": "=Country!$A$1", "Country!Local": "=Country!$B$1"})

    def test_absolute_package_relationship_is_supported(self):
        parts = base_parts()
        parts["xl/_rels/workbook.xml.rels"] = relationships([
            {"Id": "sheet1", "Type": R + "/worksheet", "Target": "/xl/worksheets/sheet1.xml"},
        ])
        self.assertEqual(inspect_workbook(self.write(parts)).sheet_names, ("Country",))

    def test_non_xlsx_and_invalid_zip_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "requires an .xlsx"):
            inspect_workbook(self.path.with_suffix(".xlsm"))
        self.path.write_text("not a zip package", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "Cannot inspect"):
            inspect_workbook(self.path)

    def test_missing_required_parts_and_malformed_xml_are_rejected(self):
        for part, content in (("_rels/.rels", None), ("xl/workbook.xml", None),
                              ("xl/worksheets/sheet1.xml", "<not-closed>")):
            with self.subTest(part=part):
                parts = base_parts()
                if content is None:
                    parts.pop(part)
                else:
                    parts[part] = content
                with self.assertRaises(ValueError):
                    inspect_workbook(self.write(parts))

    def test_unsafe_or_external_worksheet_relationship_is_rejected(self):
        for target, mode in (("../../outside.xml", "Internal"),
                             ("https://example.com/sheet.xml", "External"),
                             ("worksheets\\sheet1.xml", "Internal"),
                             ("worksheets/sheet1.xml#fragment", "Internal")):
            with self.subTest(target=target):
                parts = base_parts()
                parts["xl/_rels/workbook.xml.rels"] = relationships([
                    {"Id": "sheet1", "Type": R + "/worksheet", "Target": target, "TargetMode": mode},
                ])
                with self.assertRaises(ValueError):
                    inspect_workbook(self.write(parts))

    def test_invalid_shared_string_and_shared_formula_are_rejected(self):
        cases = (
            '<c r="A1" t="s"><v>0</v></c>',
            '<c r="A1"><f t="shared" si="1"/></c>',
            '<c r="A1"><f t="shared" si="1" ref="A1:A2">B1</f></c>'
            '<c r="A3"><f t="shared" si="1"/></c>',
        )
        for cells in cases:
            with self.subTest(cells=cells):
                parts = base_parts((("Country", worksheet('<row r="1">' + cells + '</row>')),))
                with self.assertRaises(ValueError):
                    inspect_workbook(self.write(parts))

    def test_duplicate_cells_and_sheet_names_are_rejected(self):
        parts = base_parts((("Country", worksheet('<row r="1"><c r="A1"/><c r="A1"/></row>')),))
        with self.assertRaisesRegex(ValueError, "Duplicate cell"):
            inspect_workbook(self.write(parts))
        parts = base_parts((("Country", worksheet()), ("COUNTRY", worksheet())))
        with self.assertRaisesRegex(ValueError, "duplicate worksheet name"):
            inspect_workbook(self.write(parts))

    def test_missing_table_relationship_is_rejected(self):
        parts = base_parts((("Country", worksheet(extra='<tableParts><tablePart r:id="missing"/></tableParts>')),))
        parts["xl/worksheets/_rels/sheet1.xml.rels"] = relationships([])
        with self.assertRaisesRegex(ValueError, "Missing table relationship"):
            inspect_workbook(self.write(parts))

    def test_unsupported_chart_sheet_and_what_if_data_table_fail_explicitly(self):
        parts = base_parts()
        parts["xl/_rels/workbook.xml.rels"] = relationships([
            {"Id": "sheet1", "Type": R + "/chartsheet", "Target": "worksheets/sheet1.xml"},
        ])
        with self.assertRaisesRegex(ValueError, "Expected a worksheet relationship"):
            inspect_workbook(self.write(parts))
        parts = base_parts((("Country", worksheet('<row r="1"><c r="A1"><f t="dataTable" ref="A1:B2"/></c></row>')),))
        with self.assertRaisesRegex(ValueError, "What-if data-table formulas are unsupported"):
            inspect_workbook(self.write(parts))


class LayoutTests(unittest.TestCase):
    def setUp(self):
        self.original = SheetInfo(
            title="Belgium_KPI",
            text_labels={"A1": "KPI", "B1": "2026", "A2": "Revenue"},
            formulas={"B2": "=1+2"},
            merged_ranges=("A4:B4",),
            tables=(TableInfo("BelgiumTable", "A1:B3", ("KPI", "2026")),),
        )

    def test_renamed_sheet_table_and_changed_formulas_do_not_change_layout(self):
        updated = replace(self.original, title="KPI", formulas={"B2": "=5+6"},
                          tables=(TableInfo("SourceTable", "A1:B3", ("KPI", "2026")),))
        self.assertIsNone(check_layout(self.original, updated))

    def test_moved_or_changed_labels_are_rejected(self):
        for labels in ({"A1": "KPI", "C1": "2026", "A2": "Revenue"},
                       {"A1": "KPI", "B1": "2027", "A2": "Revenue"}):
            with self.subTest(labels=labels):
                with self.assertRaisesRegex(ValueError, "Text layout differs"):
                    check_layout(self.original, replace(self.original, text_labels=labels))

    def test_text_check_can_be_disabled_without_disabling_structure_checks(self):
        changed = replace(self.original, text_labels={"Z99": "New text"})
        check_layout(self.original, changed, check_text_layout=False)
        with self.assertRaisesRegex(ValueError, "Merged-cell layout differs"):
            check_layout(self.original, replace(changed, merged_ranges=()), check_text_layout=False)

    def test_changed_table_extent_or_columns_are_rejected(self):
        for table in (TableInfo("NewName", "A1:B4", ("KPI", "2026")),
                      TableInfo("NewName", "A1:B3", ("KPI", "2027"))):
            with self.subTest(table=table):
                with self.assertRaisesRegex(ValueError, "Table layout or column schema differs"):
                    check_layout(self.original, replace(self.original, tables=(table,)))

    def test_arrays_are_rejected_in_either_workbook_even_with_text_check_disabled(self):
        array = replace(self.original, array_ranges=("D1:D3",))
        for old, new in ((array, self.original), (self.original, array)):
            with self.assertRaisesRegex(ValueError, "Array/spill formulas are unsupported"):
                check_layout(old, new, check_text_layout=False)

    def test_text_differences_are_limited_to_ten_examples(self):
        old = SheetInfo("Country", text_labels={f"A{row}": "old" for row in range(1, 13)})
        new = SheetInfo("Country", text_labels={f"A{row}": "new" for row in range(1, 13)})
        with self.assertRaises(ValueError) as caught:
            check_layout(old, new)
        message = str(caught.exception)
        self.assertIn("(12 cells)", message)
        self.assertIn("A10:", message)
        self.assertNotIn("A11:", message)
        self.assertIn("and 2 more", message)


if __name__ == "__main__":
    unittest.main()
