"""
Integration / smoke tests for the full Excel generation pipeline.
These tests call the engine end-to-end (no GUI, no mocks) and validate
the output workbook structure.
"""
import os
import tempfile
import unittest

import pandas as pd
from openpyxl import load_workbook

import runner


def _make_df():
    """Minimal valid DataFrame matching runner.validasi_data expectations."""
    rows = []
    for division, region, channel, account, cat in [
        ('Nutrition', 'WEST', 'MT', 'ALPHA', 'INSTANT NOODLE'),
        ('Indulgence', 'EAST', 'GT', 'BETA', 'SNACK'),
    ]:
        rows.append({
            'Source Division': division,
            'Region': region,
            'Channel': channel,
            'Account': account,
            'Category Channel': cat,
            'Brand': 'INDOMIE',
            'Produsen': 'INDOFOOD',
            'Period': 'Jan 25',
            'Store Code': f'S00{len(rows) + 1}',
            'Product Code': 'SKU001',
            'Facing': 30,
        })
        rows.append({
            'Source Division': division,
            'Region': region,
            'Channel': channel,
            'Account': account,
            'Category Channel': cat,
            'Brand': 'COMPETITOR',
            'Produsen': 'COMPETITOR',
            'Period': 'Jan 25',
            'Store Code': f'S00{len(rows) + 1}',
            'Product Code': 'COMPETITOR SKU',
            'Facing': 10,
        })
    return pd.DataFrame(rows)


class SmokeTest(unittest.TestCase):
    """Full pipeline: load → validate → enrich → export → validate workbook."""

    def setUp(self):
        self._orig_sleep = runner.time.sleep
        runner.time.sleep = lambda _: None

    def tearDown(self):
        runner.time.sleep = self._orig_sleep

    def test_full_pipeline_produces_valid_workbook(self):
        df = _make_df()
        with tempfile.TemporaryDirectory() as tmp:
            df_clean, removed = runner.validasi_data(df)
            self.assertGreater(len(df_clean), 0)
            self.assertEqual(len(removed), 0)

            target_rows = [['Division', 'Dimension', 'Name', 'Target']]
            target_rows = runner._enrich_targets_with_df_values(df_clean, target_rows)
            targets = runner._target_rows_to_dict(target_rows)

            output_path = runner.export_summary_excel(
                df_clean, targets, output_dir=tmp,
                cluster_name='IntegrationTest', target_rows=target_rows,
            )

            # File exists and has non-zero size
            self.assertTrue(os.path.exists(output_path))
            self.assertGreater(os.path.getsize(output_path), 0)

            wb = load_workbook(output_path, read_only=True)
            sheets = wb.sheetnames
            wb.close()

            # Both divisions present
            self.assertIn('Nutrition', sheets)
            self.assertIn('Indulgence', sheets)

            # TARGETS sheet present
            self.assertIn('TARGETS', sheets)

    def test_workbook_has_data_in_division_sheets(self):
        df = _make_df()
        with tempfile.TemporaryDirectory() as tmp:
            df_clean, _ = runner.validasi_data(df)
            target_rows = runner._enrich_targets_with_df_values(df_clean, None)
            targets = runner._target_rows_to_dict(target_rows)
            output_path = runner.export_summary_excel(
                df_clean, targets, output_dir=tmp,
                cluster_name='DataCheck', target_rows=target_rows,
            )
            wb = load_workbook(output_path, read_only=True)
            for sheet_name in ['Nutrition', 'Indulgence']:
                ws = wb[sheet_name]
                all_values = [row for row in ws.iter_rows(values_only=True) if any(c for c in row)]
                self.assertGreater(len(all_values), 0, f'Sheet {sheet_name} is empty')
            wb.close()

    def test_targets_enrichment_appends_only_new_rows(self):
        df = _make_df()
        base = [
            ['Division', 'Dimension', 'Name', 'Target'],
            ['Nutrition', 'REGION', 'DEFAULT', 65.0],
        ]
        enriched = runner._enrich_targets_with_df_values(df, base)
        # Base rows preserved
        self.assertEqual(enriched[0], base[0])
        self.assertIn(base[1], enriched)
        # New rows appended (Indulgence not in base)
        divisions_in_enriched = {r[0] for r in enriched[1:] if r[0] is not None}
        self.assertIn('Indulgence', divisions_in_enriched)

    def test_division_targets_isolated_per_division(self):
        target_rows = [
            ['Division', 'Dimension', 'Name', 'Target'],
            ['Nutrition', 'REGION', 'WEST', 80.0],
            ['Indulgence', 'REGION', 'WEST', 60.0],
        ]
        nutrition_targets = runner._build_division_targets(target_rows, 'Nutrition')
        indulgence_targets = runner._build_division_targets(target_rows, 'Indulgence')
        self.assertEqual(nutrition_targets.get(('REGION', 'WEST')), 80.0)
        self.assertEqual(indulgence_targets.get(('REGION', 'WEST')), 60.0)

    def test_proses_data_full_flow(self):
        orig_cwd = os.getcwd()
        orig_baca = runner.baca_semua_csv
        with tempfile.TemporaryDirectory() as tmp:
            try:
                os.chdir(tmp)
                runner.baca_semua_csv = _make_df
                runner.proses_data()
                outputs = [f for f in os.listdir(tmp) if f.startswith('Summary SOS_') and f.endswith('.xlsx')]
                self.assertEqual(len(outputs), 1)
                wb = load_workbook(os.path.join(tmp, outputs[0]), read_only=True)
                self.assertIn('TARGETS', wb.sheetnames)
                wb.close()
            finally:
                runner.baca_semua_csv = orig_baca
                os.chdir(orig_cwd)

    def test_validate_workbook_helper_detects_missing_targets(self):
        from gui.utils import validate_workbook
        with tempfile.TemporaryDirectory() as tmp:
            from openpyxl import Workbook as WB
            wb = WB()
            wb.active.title = 'Data'
            path = os.path.join(tmp, 'test.xlsx')
            wb.save(path)
            err = validate_workbook(path)
            self.assertIsNotNone(err)
            self.assertIn('TARGETS', err)

    def test_validate_workbook_helper_accepts_valid_workbook(self):
        from gui.utils import validate_workbook
        df = _make_df()
        with tempfile.TemporaryDirectory() as tmp:
            df_clean, _ = runner.validasi_data(df)
            target_rows = runner._enrich_targets_with_df_values(df_clean, None)
            targets = runner._target_rows_to_dict(target_rows)
            output_path = runner.export_summary_excel(
                df_clean, targets, output_dir=tmp,
                cluster_name='Valid', target_rows=target_rows,
            )
            err = validate_workbook(output_path)
            self.assertIsNone(err, f'Expected valid workbook but got: {err}')


if __name__ == '__main__':
    unittest.main(verbosity=2)
