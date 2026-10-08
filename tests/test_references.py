"""Portable formula-reference checks used before native Excel writes."""
import unittest

from pops_update.references import ReferenceError, canonical_formula, validate_formula


class ReferenceTests(unittest.TestCase):
    def setUp(self):
        self.kwargs = dict(
            sheet_name='Germany_ID Card',
            allowed_sheets={'Germany_ID Card', 'Germany_FTE', 'ID_CARD'},
            allowed_tables=set(),
            source_names={},
            destination_names={},
        )

    def test_local_sheet_range_is_allowed(self):
        validate_formula("='Germany_FTE'!$B$2", **self.kwargs)
        validate_formula("=SUM('Germany_ID Card'!B2:B4)", **self.kwargs)

    def test_external_and_indirect_references_are_rejected(self):
        for formula in ("='[country.xlsx]ID Card'!B2", '=INDIRECT("A1")', '=A1+#REF!'):
            with self.assertRaises(ReferenceError):
                validate_formula(formula, **self.kwargs)

    def test_unmapped_sheet_is_rejected(self):
        with self.assertRaises(ReferenceError):
            validate_formula("='Other Country'!B2", **self.kwargs)

    def test_formula_canonicalisation_ignores_case_and_spaces(self):
        self.assertEqual(canonical_formula('=SUM(A1:A2)'), canonical_formula('=sum( A1:A2 )'))


if __name__ == '__main__':
    unittest.main()
