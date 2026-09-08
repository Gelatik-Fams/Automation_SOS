import unittest
import os
import tempfile

from openpyxl import Workbook, load_workbook
import pandas as pd

import runner


def dataframe_with_produsen(rows):
    df = pd.DataFrame(rows)
    if "Product Code" in df.columns and "Produsen" not in df.columns:
        is_competitor = df["Product Code"].astype(str).str.contains(
            "COMPETITOR", case=False, na=False
        )
        df["Produsen"] = is_competitor.map({True: "COMPETITOR", False: "INDOFOOD"})
    return df


class SourceDivisionTests(unittest.TestCase):
    def test_extracts_source_division_from_last_filename_segment(self):
        self.assertEqual(
            runner.extract_source_division_from_filename(
                "Report Product - Mei 25 - Oil & Fat.csv"
            ),
            "Oil & Fat",
        )

    def test_old_filename_without_division_returns_unknown(self):
        self.assertEqual(
            runner.extract_source_division_from_filename(
                "Report Product - Januari 2024.csv"
            ),
            runner.UNKNOWN_SOURCE_DIVISION,
        )

    def test_extracts_source_division_from_raw_divisi_column(self):
        df = pd.DataFrame({"Divisi": ["Pasta", "Pasta", " Pasta "]})

        self.assertEqual(
            runner.extract_source_division_from_raw_data(df, "Report Product - Jan 25.csv"),
            "Pasta",
        )

    def test_extract_source_division_requires_divisi_column(self):
        df = pd.DataFrame({"Visit Date": ["2025-01-02"]})

        with self.assertRaisesRegex(ValueError, "kolom Divisi tidak ditemukan"):
            runner.extract_source_division_from_raw_data(df, "Report Product - Jan 25.csv")

    def test_extract_source_division_rejects_mixed_division_file(self):
        df = pd.DataFrame({"Divisi": ["Pasta", "Noodle"]})

        with self.assertRaisesRegex(ValueError, "lebih dari satu Divisi"):
            runner.extract_source_division_from_raw_data(df, "Report Product - Jan 25.csv")


class FileDiscoveryTests(unittest.TestCase):
    def test_discovers_report_product_csv_in_root_and_division_folders(self):
        with tempfile.TemporaryDirectory() as tmp:
            paths = [
                os.path.join(tmp, "Report Product - Jan 25.csv"),
                os.path.join(tmp, "SIMP", "Report Product - Jan 25 - SIMP.csv"),
                os.path.join(tmp, "NICI", "Report Product - Jan 25 - NICI.csv"),
                os.path.join(tmp, "IFM", "Report Product - Jan 25 - IFM.csv"),
                os.path.join(tmp, "Noodle", "Report Product - Jan 25 - Noodle.csv"),
            ]
            for path in paths:
                os.makedirs(os.path.dirname(path), exist_ok=True)
                with open(path, "w", encoding="utf-8") as f:
                    f.write("Visit Date,Product Code,Facing\n")
            with open(os.path.join(tmp, "Other.csv"), "w", encoding="utf-8") as f:
                f.write("Visit Date,Product Code,Facing\n")

            found = [
                os.path.relpath(p, tmp).replace(os.sep, "/")
                for p in runner.discover_report_product_files(tmp)
            ]

        self.assertEqual(found, [
            "IFM/Report Product - Jan 25 - IFM.csv",
            "NICI/Report Product - Jan 25 - NICI.csv",
            "Noodle/Report Product - Jan 25 - Noodle.csv",
            "Report Product - Jan 25.csv",
            "SIMP/Report Product - Jan 25 - SIMP.csv",
        ])

    def test_counts_files_per_source_division_from_filenames(self):
        files = [
            "SIMP/Report Product - Jan 25 - SIMP.csv",
            "SIMP/Report Product - Feb 25 - SIMP.csv",
            "NICI/Report Product - Jan 25 - NICI.csv",
            "Report Product - Jan 25.csv",
        ]

        self.assertEqual(runner.count_files_by_source_division(files), {
            "NICI": 1,
            "SIMP": 2,
            runner.UNKNOWN_SOURCE_DIVISION: 1,
        })

    def test_baca_semua_csv_keeps_physical_line_count_for_user_log(self):
        original_cwd = os.getcwd()
        with tempfile.TemporaryDirectory() as tmp:
            try:
                os.chdir(tmp)
                with open("Report Product - Jan 25 - Pasta.csv", "w", encoding="utf-8", newline="") as f:
                    f.write("Visit Date,Product Code,Produsen,Divisi,Facing\n")
                    f.write('"2025-01-02","SKU\n001","INDOFOOD","Pasta",10\n')
                    f.write('"2025-01-02","COMPETITOR SKU","COMPETITOR","Pasta",5\n')

                df = runner.baca_semua_csv()

                self.assertEqual(df.attrs["physical_csv_lines"], 4)
                self.assertEqual(len(df), 2)
                self.assertEqual(set(df["Source Division"]), {"Pasta"})
            finally:
                os.chdir(original_cwd)

    def test_baca_semua_csv_uses_divisi_column_even_when_filename_has_no_division(self):
        original_cwd = os.getcwd()
        with tempfile.TemporaryDirectory() as tmp:
            try:
                os.chdir(tmp)
                with open("Report Product - Jan 25.csv", "w", encoding="utf-8", newline="") as f:
                    f.write("Visit Date,Product Code,Produsen,Divisi,Facing\n")
                    f.write('"2025-01-02","SKU001","INDOFOOD","Pasta",10\n')

                df = runner.baca_semua_csv()

                self.assertEqual(df["Source Division"].tolist(), ["Pasta"])
            finally:
                os.chdir(original_cwd)

    def test_baca_semua_csv_reads_each_month_file_and_detects_same_division(self):
        original_cwd = os.getcwd()
        with tempfile.TemporaryDirectory() as tmp:
            try:
                os.chdir(tmp)
                for month in ["Jan", "Feb", "Mar"]:
                    with open(f"Report Product - {month} 25.csv", "w", encoding="utf-8", newline="") as f:
                        f.write("Visit Date,Product Code,Produsen,Divisi,Facing\n")
                        f.write(f'"2025-01-02","SKU-{month}","INDOFOOD","Pasta",10\n')

                df = runner.baca_semua_csv()

                self.assertEqual(len(df), 3)
                self.assertEqual(set(df["Source Division"]), {"Pasta"})
            finally:
                os.chdir(original_cwd)

    def test_baca_semua_csv_fails_when_divisi_column_missing(self):
        original_cwd = os.getcwd()
        with tempfile.TemporaryDirectory() as tmp:
            try:
                os.chdir(tmp)
                with open("Report Product - Jan 25.csv", "w", encoding="utf-8", newline="") as f:
                    f.write("Visit Date,Product Code,Produsen,Facing\n")
                    f.write('"2025-01-02","SKU001","INDOFOOD",10\n')

                with self.assertRaisesRegex(ValueError, "kolom Divisi tidak ditemukan"):
                    runner.baca_semua_csv()
            finally:
                os.chdir(original_cwd)

    def test_baca_semua_csv_fails_when_one_file_has_multiple_divisions(self):
        original_cwd = os.getcwd()
        with tempfile.TemporaryDirectory() as tmp:
            try:
                os.chdir(tmp)
                with open("Report Product - Jan 25.csv", "w", encoding="utf-8", newline="") as f:
                    f.write("Visit Date,Product Code,Produsen,Divisi,Facing\n")
                    f.write('"2025-01-02","SKU001","INDOFOOD","Pasta",10\n')
                    f.write('"2025-01-02","SKU002","INDOFOOD","Noodle",10\n')

                with self.assertRaisesRegex(ValueError, "lebih dari satu Divisi"):
                    runner.baca_semua_csv()
            finally:
                os.chdir(original_cwd)


class AppsScriptRemovalTests(unittest.TestCase):
    def test_apps_script_file_is_not_required_for_excel_output(self):
        self.assertFalse(os.path.exists("dashboard_division_filter.gs"))


class ValidationFreezeTests(unittest.TestCase):
    def test_sos_classifies_competitor_from_produsen_not_product_code(self):
        df = dataframe_with_produsen([
            {
                "Region": "BANDUNG",
                "Period": "Jan 25",
                "Produsen": "INDOFOOD",
                "Product Code": "SKU001",
                "Facing": 30,
            },
            {
                "Region": "BANDUNG",
                "Period": "Jan 25",
                "Produsen": "COMPETITOR",
                "Product Code": "'I401",
                "Facing": 20,
            },
        ])

        result = runner.calc_sos(df, ["Region"])

        self.assertEqual(result.loc[0, "fi"], 30)
        self.assertEqual(result.loc[0, "fk"], 20)
        self.assertEqual(result.loc[0, "SOS%"], 60.0)

    def test_dashboard_does_not_keep_unused_compliance_dead_code(self):
        with open("runner.py", encoding="utf-8") as f:
            source = f.read()

        for token in ("calc_" + "compliance", "compliance_" + "map"):
            self.assertEqual(source.count(token), 0, token)

    def test_source_division_metadata_does_not_change_duplicate_detection(self):
        df = dataframe_with_produsen([
            {
                "_source_file": "Report Product - Jan 25 - Beverage",
                "Source Division": "Beverage",
                "Visit Date": "2025-01-02",
                "Store Code": "S001",
                "Product Code": "SKU001",
                "Facing": 10,
            },
            {
                "_source_file": "Report Product - Jan 25 - Pasta",
                "Source Division": "Pasta",
                "Visit Date": "2025-01-02",
                "Store Code": "S001",
                "Product Code": "SKU001",
                "Facing": 10,
            },
        ])

        cleaned, removed = runner.validasi_data(df)

        self.assertEqual(len(cleaned), 1)
        self.assertEqual(len(removed), 1)


class LocalTargetsFileTests(unittest.TestCase):
    def test_default_target_rows_use_embedded_four_column_structure(self):
        rows = runner.default_target_rows(["Noodle", "Snack"])

        self.assertEqual(rows[0], ["Division", "Dimension", "Name", "Target"])
        self.assertIn(["Noodle", "REGION", "DEFAULT", 65.0], rows)
        self.assertIn(["Snack", "REGION", "DEFAULT", 65.0], rows)

    def test_missing_summary_uses_default_targets_without_creating_targets_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "Summary SOS_Test.xlsx")

            targets, target_rows = runner.load_or_create_summary_targets(path)

            self.assertFalse(os.path.exists(path))
            self.assertFalse(os.path.exists(os.path.join(tmp, "TARGETS.xlsx")))
            self.assertEqual(targets[("REGION", "DEFAULT")], 65.0)
            self.assertIsNone(target_rows)

    def test_existing_summary_targets_sheet_is_read_without_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "Summary SOS_Test.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.title = "DASHBOARD"
            targets_ws = wb.create_sheet("TARGETS")
            targets_ws.append(["Division", "Dimension", "Name", "Target"])
            targets_ws.append(["Noodle", "REGION", "WEST", 80])
            targets_ws.append(["Snack", "CHANNEL", "MT", "70,5"])
            wb.save(path)

            targets, target_rows = runner.load_or_create_summary_targets(path)

            self.assertEqual(targets[("REGION", "WEST")], 80.0)
            self.assertEqual(targets[("CHANNEL", "MT")], 70.5)
            self.assertEqual(target_rows[1], ["Noodle", "REGION", "WEST", 80])

            wb_after = load_workbook(path)
            self.assertEqual(wb_after.sheetnames, ["DASHBOARD", "TARGETS"])

    def test_invalid_summary_targets_sheet_raises_clear_error_without_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "Summary SOS_Test.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.title = "DASHBOARD"
            targets_ws = wb.create_sheet("TARGETS")
            targets_ws.append(["Bad", "Header"])
            wb.save(path)

            with self.assertRaises(ValueError) as ctx:
                runner.load_or_create_summary_targets(path)

            self.assertIn("TARGETS", str(ctx.exception))
            wb_after = load_workbook(path)
            self.assertEqual(wb_after.sheetnames, ["DASHBOARD", "TARGETS"])

class ExcelExportTests(unittest.TestCase):
    def make_df(self):
        return DashboardDivisionAwareTests().make_dashboard_df()

    def test_export_summary_excel_creates_only_dashboard_and_targets(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_path = runner.export_summary_excel(
                self.make_df(),
                {("REGION", "WEST"): 70},
                output_dir=tmp,
                cluster_name="Indulgence",
            )

            self.assertEqual(
                os.path.basename(output_path),
                "Summary SOS_Indulgence.xlsx",
            )
            wb = load_workbook(output_path)
            self.assertEqual(wb.sheetnames, ["Noodle", "Snack", "TARGETS", "VALIDATION REPORT"])
            targets_ws = wb["TARGETS"]
            target_divisions = {
                targets_ws.cell(row=row_number, column=1).value
                for row_number in range(2, targets_ws.max_row + 1)
            }
            self.assertEqual(target_divisions, {"Noodle", "Snack"})
            for sheet_name in ["Noodle", "Snack"]:
                ws = wb[sheet_name]
                self.assertEqual(ws["A1"].value, "DIVISION SUMMARY")
                self.assertNotEqual(ws["B1"].value, "ALL")
                self.assertFalse(ws.data_validations.dataValidation)
                self.assertFalse(any(ws.row_dimensions[row].hidden for row in range(1, ws.max_row + 1)))
                values = [row[0] for row in ws.iter_rows(values_only=True) if row and row[0]]
                self.assertFalse(any(str(value).startswith("CHART SOURCE ") for value in values))

    def test_export_summary_excel_creates_one_sheet_per_source_division(self):
        df = self.make_df()
        extra = []
        for division in ["Nici", "Oil & Fat"]:
            row = df.iloc[0].copy()
            row["Source Division"] = division
            row["Region"] = division
            extra.append(row)
            comp = df.iloc[1].copy()
            comp["Source Division"] = division
            comp["Region"] = division
            extra.append(comp)
        df = pd.concat([df, dataframe_with_produsen(extra)], ignore_index=True)

        with tempfile.TemporaryDirectory() as tmp:
            output_path = runner.export_summary_excel(
                df,
                {("REGION", "WEST"): 70},
                output_dir=tmp,
                cluster_name="Indulgence",
            )
            wb = load_workbook(output_path)

        self.assertEqual(wb.sheetnames, ["Nici", "Noodle", "Oil & Fat", "Snack", "TARGETS", "VALIDATION REPORT"])

    def test_excel_dashboard_removes_facing_columns_and_keeps_sos_values(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_path = runner.export_summary_excel(
                self.make_df(),
                {("REGION", "WEST"): 70},
                output_dir=tmp,
                cluster_name="Indulgence",
            )
            ws = load_workbook(output_path)["Noodle"]
            rows = list(ws.iter_rows(values_only=True))

        title_row = next(i for i, row in enumerate(rows) if row[0] == "SOS% BY REGION")
        header1 = list(rows[title_row + 1])
        header2 = list(rows[title_row + 2])
        noodle_row = next(row for row in rows if row[:2] == ("Noodle", "WEST"))

        self.assertIn("TARGET", header1)
        self.assertIn("Jan 25", header1)
        self.assertIn("SOS%", header2)
        self.assertNotIn("Indofood", header2)
        self.assertNotIn("Kompetitor", header2)
        self.assertNotIn("Total", header2)
        self.assertEqual(list(noodle_row[:4]), ["Noodle", "WEST", 70, 75.0])

    def test_excel_conditional_format_ranges_only_target_detail_sos_cells(self):
        payload = runner.dashboard_payload_sos_only(
            runner.build_dashboard_payload(
                self.make_df(),
                {("REGION", "WEST"): 70},
            )
        )

        ranges = runner.get_excel_conditional_format_ranges(payload)

        self.assertTrue(ranges)
        self.assertFalse(any(r["chart_source"] for r in ranges))
        for cf_range in ranges:
            self.assertLess(cf_range["end_row"], cf_range["grand_row"])
            for subtotal_row in cf_range["subtotal_rows"]:
                self.assertFalse(
                    cf_range["start_row"] <= subtotal_row <= cf_range["end_row"],
                    cf_range,
                )

    def test_excel_conditional_formatting_uses_green_pass_fill_and_same_row_target(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_path = runner.export_summary_excel(
                self.make_df(),
                {("REGION", "WEST"): 70},
                output_dir=tmp,
                cluster_name="Indulgence",
            )
            ws = load_workbook(output_path)["Noodle"]

        rules = []
        for cf_range, cf_rules in ws.conditional_formatting._cf_rules.items():
            for rule in cf_rules:
                color = getattr(rule.dxf.fill.fgColor, "rgb", "") if rule.dxf and rule.dxf.fill else ""
                rules.append((str(cf_range), rule.formula[0], color))

        pass_rules = [rule for rule in rules if ">=$" in rule[1]]
        fail_rules = [rule for rule in rules if "<$" in rule[1]]

        self.assertTrue(pass_rules)
        self.assertTrue(fail_rules)
        self.assertTrue(any(str(color).endswith("C4D79B") for _, _, color in pass_rules))
        self.assertTrue(any(str(color).endswith("E6B8B7") for _, _, color in fail_rules))
        cf_fonts = [
            rule.dxf.font
            for _cf_range, cf_rules in ws.conditional_formatting._cf_rules.items()
            for rule in cf_rules
            if rule.dxf
        ]
        self.assertTrue(all(font is None for font in cf_fonts))
        for _cell_range, formula, _color in rules:
            left, right = formula.split(">=") if ">=" in formula else formula.split("<")
            self.assertEqual(
                "".join(ch for ch in left if ch.isdigit()),
                "".join(ch for ch in right if ch.isdigit()),
            )

    def test_export_summary_excel_preserves_existing_embedded_target_rows(self):
        target_rows = [
            ["Division", "Dimension", "Name", "Target"],
            ["Noodle", "REGION", "WEST", 80],
            ["Snack", "CHANNEL", "GT", 72],
        ]
        targets = runner._target_rows_to_dict(target_rows)

        with tempfile.TemporaryDirectory() as tmp:
            output_path = runner.export_summary_excel(
                self.make_df(),
                targets,
                output_dir=tmp,
                cluster_name="Indulgence",
                target_rows=target_rows,
            )
            ws = load_workbook(output_path)["TARGETS"]
            rows = [
                [ws.cell(row=row_number, column=col).value for col in range(1, 5)]
                for row_number in range(1, ws.max_row + 1)
            ]

        self.assertEqual(rows[:len(target_rows)], target_rows)

    def test_export_summary_excel_applies_targets_by_division_dimension_and_name(self):
        target_rows = [
            ["Division", "Dimension", "Name", "Target"],
            ["Noodle", "REGION", "WEST", 80],
            ["Snack", "REGION", "EAST", 55],
        ]
        targets = runner._target_rows_to_dict(target_rows)

        with tempfile.TemporaryDirectory() as tmp:
            output_path = runner.export_summary_excel(
                self.make_df(),
                targets,
                output_dir=tmp,
                cluster_name="Indulgence",
                target_rows=target_rows,
            )
            wb = load_workbook(output_path)
            noodle_rows = list(wb["Noodle"].iter_rows(values_only=True))
            snack_rows = list(wb["Snack"].iter_rows(values_only=True))

        noodle_region = next(row for row in noodle_rows if row[:2] == ("Noodle", "WEST"))
        snack_region = next(row for row in snack_rows if row[:2] == ("Snack", "EAST"))
        self.assertEqual(noodle_region[2], 80)
        self.assertEqual(snack_region[2], 55)

    def test_excel_category_charts_have_data_labels_and_non_overlapping_category_rows(self):
        df = pd.concat([
            self.make_df(),
            dataframe_with_produsen([
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "CUP NOODLE",
                    "Brand": "POP MIE",
                    "Period": "Jan 25",
                    "Store Code": "S003",
                    "Product Code": "SKU003",
                    "Facing": 12,
                },
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "CUP NOODLE",
                    "Brand": "COMP",
                    "Period": "Jan 25",
                    "Store Code": "S003",
                    "Product Code": "COMPETITOR SKU",
                    "Facing": 8,
                },
            ]),
        ], ignore_index=True)

        with tempfile.TemporaryDirectory() as tmp:
            output_path = runner.export_summary_excel(
                df,
                {("REGION", "WEST"): 70},
                output_dir=tmp,
                cluster_name="Indulgence",
            )
            ws = load_workbook(output_path)["Noodle"]

        category_charts = [chart for chart in ws._charts if " - " in str(chart.title)]
        anchors = sorted(chart.anchor._from.row for chart in category_charts)
        self.assertGreaterEqual(len(category_charts), 2)
        self.assertGreaterEqual(anchors[1] - anchors[0], runner.EXCEL_CATEGORY_CHART_ROW_STEP - 1)
        self.assertTrue(all(chart.dLbls and chart.dLbls.showVal for chart in category_charts))
        self.assertTrue(all(chart.x_axis.delete is False for chart in category_charts))
        self.assertTrue(all(chart.x_axis.tickLblPos == "low" for chart in category_charts))
        self.assertTrue(all(chart.series and chart.series[0].cat is not None for chart in category_charts))
        self.assertTrue(all(chart.series[0].cat.strRef is not None for chart in category_charts))

    def test_excel_category_charts_adaptive_width_and_position(self):
        df = pd.concat([
            self.make_df(),
            dataframe_with_produsen([
                # CUP NOODLE - Jan 25 (2 brands)
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "CUP NOODLE",
                    "Brand": "POP MIE",
                    "Period": "Jan 25",
                    "Store Code": "S003",
                    "Product Code": "SKU003",
                    "Facing": 12,
                },
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "CUP NOODLE",
                    "Brand": "COMP1",
                    "Period": "Jan 25",
                    "Store Code": "S003",
                    "Product Code": "COMPETITOR SKU 1",
                    "Facing": 8,
                },
                # CUP NOODLE - Feb 25 (same 2 brands)
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "CUP NOODLE",
                    "Brand": "POP MIE",
                    "Period": "Feb 25",
                    "Store Code": "S003",
                    "Product Code": "SKU003",
                    "Facing": 15,
                },
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "CUP NOODLE",
                    "Brand": "COMP1",
                    "Period": "Feb 25",
                    "Store Code": "S003",
                    "Product Code": "COMPETITOR SKU 1",
                    "Facing": 5,
                },
                # BAG NOODLE - Jan 25 (6 brands)
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "BAG NOODLE",
                    "Brand": "INDOMIE",
                    "Period": "Jan 25",
                    "Store Code": "S003",
                    "Product Code": "SKU001",
                    "Facing": 12,
                },
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "BAG NOODLE",
                    "Brand": "COMP1",
                    "Period": "Jan 25",
                    "Store Code": "S003",
                    "Product Code": "COMPETITOR SKU 1",
                    "Facing": 8,
                },
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "BAG NOODLE",
                    "Brand": "COMP2",
                    "Period": "Jan 25",
                    "Store Code": "S003",
                    "Product Code": "COMPETITOR SKU 2",
                    "Facing": 5,
                },
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "BAG NOODLE",
                    "Brand": "COMP3",
                    "Period": "Jan 25",
                    "Store Code": "S003",
                    "Product Code": "COMPETITOR SKU 3",
                    "Facing": 5,
                },
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "BAG NOODLE",
                    "Brand": "COMP4",
                    "Period": "Jan 25",
                    "Store Code": "S003",
                    "Product Code": "COMPETITOR SKU 4",
                    "Facing": 5,
                },
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "BAG NOODLE",
                    "Brand": "COMP5",
                    "Period": "Jan 25",
                    "Store Code": "S003",
                    "Product Code": "COMPETITOR SKU 5",
                    "Facing": 5,
                },
                # BAG NOODLE - Feb 25 (same 6 brands)
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "BAG NOODLE",
                    "Brand": "INDOMIE",
                    "Period": "Feb 25",
                    "Store Code": "S003",
                    "Product Code": "SKU001",
                    "Facing": 12,
                },
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "BAG NOODLE",
                    "Brand": "COMP1",
                    "Period": "Feb 25",
                    "Store Code": "S003",
                    "Product Code": "COMPETITOR SKU 1",
                    "Facing": 8,
                },
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "BAG NOODLE",
                    "Brand": "COMP2",
                    "Period": "Feb 25",
                    "Store Code": "S003",
                    "Product Code": "COMPETITOR SKU 2",
                    "Facing": 5,
                },
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "BAG NOODLE",
                    "Brand": "COMP3",
                    "Period": "Feb 25",
                    "Store Code": "S003",
                    "Product Code": "COMPETITOR SKU 3",
                    "Facing": 5,
                },
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "BAG NOODLE",
                    "Brand": "COMP4",
                    "Period": "Feb 25",
                    "Store Code": "S003",
                    "Product Code": "COMPETITOR SKU 4",
                    "Facing": 5,
                },
                {
                    "Source Division": "Noodle",
                    "Region": "WEST",
                    "Channel": "MT",
                    "Account": "ALPHA",
                    "Category Channel": "BAG NOODLE",
                    "Brand": "COMP5",
                    "Period": "Feb 25",
                    "Store Code": "S003",
                    "Product Code": "COMPETITOR SKU 5",
                    "Facing": 5,
                },
            ]),
        ], ignore_index=True)

        saved_wb = None
        orig_save = Workbook.save
        def mock_save(self, filename):
            nonlocal saved_wb
            saved_wb = self
            orig_save(self, filename)

        from unittest.mock import patch
        with patch('openpyxl.Workbook.save', mock_save):
            with tempfile.TemporaryDirectory() as tmp:
                runner.export_summary_excel(
                    df,
                    {("REGION", "WEST"): 70},
                    output_dir=tmp,
                    cluster_name="Indulgence",
                )
        
        ws = saved_wb["Noodle"]

        # Filter category charts
        category_charts = [chart for chart in ws._charts if " - " in str(chart.title)]
        
        import re
        from openpyxl.utils import column_index_from_string
        def get_chart_col_idx(c):
            col_letter = re.match(r"^([A-Z]+)", c.anchor).group(1)
            return column_index_from_string(col_letter)

        # Verify charts are scaled adaptively
        cup_noodle_charts = sorted([c for c in category_charts if "CUP NOODLE" in str(c.title)], key=get_chart_col_idx)
        bag_noodle_charts = sorted([c for c in category_charts if "BAG NOODLE" in str(c.title)], key=get_chart_col_idx)
        
        self.assertEqual(len(cup_noodle_charts), 2)
        self.assertEqual(len(bag_noodle_charts), 2)
        
        # Both categories should have uniform dimensions based on the max brand count in the division (7 brands)
        # 6.0 + 7 * 1.5 = 16.5 cm width
        self.assertAlmostEqual(cup_noodle_charts[0].width, 16.5)
        self.assertAlmostEqual(bag_noodle_charts[0].width, 16.5)
        
        # Spacing test: anchor column is uniform for all categories on the sheet
        # Max brands = 7 -> width = 16.5 -> chart_width_cols = int(16.5 / 2.7) + 1 = 7
        # Since base_col is 4 (D), first period chart (Jan 25) starts at:
        # Cup noodle: D (4)
        # Bag noodle: D (4)
        # Second period chart (Feb 25) anchor column is at col_letter(4 - 1 + 1*7) = col_letter(10) = K
        # This keeps the months vertically aligned across all categories!
        
        self.assertTrue(cup_noodle_charts[0].anchor.startswith('D'))
        self.assertTrue(cup_noodle_charts[1].anchor.startswith('K'))
        
        self.assertTrue(bag_noodle_charts[0].anchor.startswith('D'))
        self.assertTrue(bag_noodle_charts[1].anchor.startswith('K'))

    def test_export_summary_excel_has_no_vba_dependency_or_macro_output(self):
        with open("runner.py", encoding="utf-8") as f:
            source = f.read()

        self.assertNotIn("win32com", source)
        self.assertNotIn("DispatchEx", source)
        self.assertNotIn("Worksheet_Change", source)
        self.assertNotIn("Workbook_Open", source)
        self.assertNotIn("add_vba_to_workbook", source)
        self.assertNotIn("enable_vba", source)

class StoreDetailExportTests(unittest.TestCase):
    def make_df(self):
        return dataframe_with_produsen([
            {
                "Source Division": "Noodle",
                "Region": "WEST",
                "Area": "BANDUNG",
                "Channel": "GT",
                "Account": "RETAIL",
                "Store Code": "1001",
                "Store Name": "TOKO A",
                "Period": "Jan 25",
                "Product Code": "SKU_INDOFOOD",
                "Facing": 10,
            },
            {
                "Source Division": "Noodle",
                "Region": "WEST",
                "Area": "BANDUNG",
                "Channel": "GT",
                "Account": "RETAIL",
                "Store Code": "1001",
                "Store Name": "TOKO A",
                "Period": "Jan 25",
                "Product Code": "SKU_COMPETITOR",
                "Facing": 30,
            },
            {
                "Source Division": "Noodle",
                "Region": "WEST",
                "Area": "BANDUNG",
                "Channel": "GT",
                "Account": "RETAIL",
                "Store Code": "1001",
                "Store Name": "TOKO A",
                "Period": "Feb 25",
                "Product Code": "SKU_INDOFOOD",
                "Facing": 20,
            },
        ])

    def test_export_store_detail_creates_correct_structure(self):
        with tempfile.TemporaryDirectory() as tmp:
            df = self.make_df()
            output_path = runner.export_store_detail_excel(df, output_dir=tmp, cluster_name="TestCluster")
            
            self.assertTrue(os.path.exists(output_path))
            self.assertTrue(output_path.endswith("Store Detail_TestCluster.xlsx"))
            
            wb = load_workbook(output_path)
            self.assertIn("Noodle", wb.sheetnames)
            ws = wb["Noodle"]
            
            rows = list(ws.iter_rows(values_only=True))
            
            # Header 1
            self.assertEqual(
                list(rows[0]),
                ['REGION', 'AREA', 'CHANNEL', 'ACCOUNT', 'STORE CODE', 'STORE NAME', 'Jan 25', 'Feb 25']
            )
            # Header 2
            self.assertEqual(
                list(rows[1]),
                [None, None, None, None, None, None, 'SOS%', 'SOS%']
            )
            # Data row
            # Jan 25: 10 Indofood, 30 Competitor = 25.0% SOS
            # Feb 25: 20 Indofood, 0 Competitor = 100.0% SOS
            self.assertEqual(
                list(rows[2]),
                ['WEST', 'BANDUNG', 'GT', 'RETAIL', '1001', 'TOKO A', 0.25, 1.0]
            )

class DashboardDivisionAwareTests(unittest.TestCase):
    def make_dashboard_df(self):
        return dataframe_with_produsen([
            {
                "Source Division": "Noodle",
                "Region": "WEST",
                "Channel": "MT",
                "Account": "ALPHA",
                "Category Channel": "BAG NOODLE",
                "Brand": "INDOMIE",
                "Period": "Jan 25",
                "Store Code": "S001",
                "Product Code": "SKU001",
                "Facing": 30,
            },
            {
                "Source Division": "Noodle",
                "Region": "WEST",
                "Channel": "MT",
                "Account": "ALPHA",
                "Category Channel": "BAG NOODLE",
                "Brand": "COMP",
                "Period": "Jan 25",
                "Store Code": "S001",
                "Product Code": "COMPETITOR SKU",
                "Facing": 10,
            },
            {
                "Source Division": "Snack",
                "Region": "EAST",
                "Channel": "GT",
                "Account": "BETA",
                "Category Channel": "OTHER",
                "Brand": "CHIKI",
                "Period": "Jan 25",
                "Store Code": "S002",
                "Product Code": "SKU002",
                "Facing": 10,
            },
            {
                "Source Division": "Snack",
                "Region": "EAST",
                "Channel": "GT",
                "Account": "BETA",
                "Category Channel": "OTHER",
                "Brand": "COMP",
                "Period": "Jan 25",
                "Store Code": "S002",
                "Product Code": "COMPETITOR SKU",
                "Facing": 30,
            },
        ])

    def test_monthly_sos_table_can_group_by_source_division_and_region(self):
        df = dataframe_with_produsen([
            {
                "Source Division": "Noodle",
                "Region": "WEST",
                "Period": "Jan 25",
                "Product Code": "SKU001",
                "Facing": 30,
            },
            {
                "Source Division": "Noodle",
                "Region": "WEST",
                "Period": "Jan 25",
                "Product Code": "COMPETITOR SKU",
                "Facing": 10,
            },
            {
                "Source Division": "Snack",
                "Region": "WEST",
                "Period": "Jan 25",
                "Product Code": "SKU002",
                "Facing": 10,
            },
            {
                "Source Division": "Snack",
                "Region": "WEST",
                "Period": "Jan 25",
                "Product Code": "COMPETITOR SKU",
                "Facing": 30,
            },
        ])

        rows, meta = runner.buat_tabel_sos_monthly(
            df,
            ["Source Division", "Region"],
            "REGION",
            ["Jan 25"],
            {("REGION", "WEST"): 70},
        )

        self.assertEqual(rows[0][:3], ["DIVISION", "Region", "TARGET"])
        self.assertIn(["Noodle", "WEST", 70, 30, 10, 40, 75.0], rows)
        self.assertIn(["Snack", "WEST", 70, 10, 30, 40, 25.0], rows)
        self.assertIn(["Noodle TOTAL", "", 65.0, 30, 10, 40, 75.0], rows)
        self.assertIn(["Snack TOTAL", "", 65.0, 10, 30, 40, 25.0], rows)
        self.assertEqual(meta["target_col_idx"], 2)
        self.assertEqual(meta["sos_col_indices"], [6])

    def test_channel_account_table_keeps_source_divisions_separate(self):
        df = dataframe_with_produsen([
            {
                "Source Division": "Noodle",
                "Channel": "MT",
                "Account": "ALPHA",
                "Period": "Jan 25",
                "Store Code": "S001",
                "Product Code": "SKU001",
                "Facing": 8,
            },
            {
                "Source Division": "Noodle",
                "Channel": "MT",
                "Account": "ALPHA",
                "Period": "Jan 25",
                "Store Code": "S001",
                "Product Code": "COMPETITOR SKU",
                "Facing": 2,
            },
            {
                "Source Division": "Snack",
                "Channel": "MT",
                "Account": "ALPHA",
                "Period": "Jan 25",
                "Store Code": "S002",
                "Product Code": "SKU002",
                "Facing": 3,
            },
            {
                "Source Division": "Snack",
                "Channel": "MT",
                "Account": "ALPHA",
                "Period": "Jan 25",
                "Store Code": "S002",
                "Product Code": "COMPETITOR SKU",
                "Facing": 7,
            },
        ])

        rows, meta = runner.buat_tabel_channel_account(
            df,
            ["Jan 25"],
            {("CHANNEL-ACCOUNT", "MT - ALPHA"): 65},
        )

        self.assertEqual(rows[0][:4], ["DIVISION", "CHANNEL", "ACCOUNT", "TARGET"])
        self.assertIn(["Noodle", "MT", "ALPHA", 65, 8, 2, 10, 80.0], rows)
        self.assertIn(["Snack", "MT", "ALPHA", 65, 3, 7, 10, 30.0], rows)
        self.assertIn(["Noodle", "MT TOTAL", "", "", 8, 2, 10, 80.0], rows)
        self.assertIn(["Snack", "MT TOTAL", "", "", 3, 7, 10, 30.0], rows)
        self.assertEqual(meta["target_col_idx"], 3)
        self.assertEqual(meta["sos_col_indices"], [7])

    def test_single_division_skips_division_subtotal_row(self):
        """When data has only 1 division, [Divisi] TOTAL row should NOT appear."""
        df = dataframe_with_produsen([
            {
                "Source Division": "Noodle",
                "Region": "WEST",
                "Period": "Jan 25",
                "Product Code": "SKU001",
                "Facing": 30,
            },
            {
                "Source Division": "Noodle",
                "Region": "WEST",
                "Period": "Jan 25",
                "Product Code": "COMPETITOR SKU",
                "Facing": 10,
            },
        ])

        rows, meta = runner.buat_tabel_sos_monthly(
            df,
            ["Source Division", "Region"],
            "REGION",
            ["Jan 25"],
            {("REGION", "WEST"): 70},
        )

        first_cells = [row[0] for row in rows]
        self.assertNotIn("Noodle TOTAL", first_cells)
        self.assertIn("GRAND TOTAL", first_cells)

    def test_multi_division_keeps_division_subtotal_rows(self):
        """When data has >1 divisions, [Divisi] TOTAL rows should still appear."""
        df = dataframe_with_produsen([
            {
                "Source Division": "Noodle",
                "Region": "WEST",
                "Period": "Jan 25",
                "Product Code": "SKU001",
                "Facing": 30,
            },
            {
                "Source Division": "Noodle",
                "Region": "WEST",
                "Period": "Jan 25",
                "Product Code": "COMPETITOR SKU",
                "Facing": 10,
            },
            {
                "Source Division": "Snack",
                "Region": "WEST",
                "Period": "Jan 25",
                "Product Code": "SKU002",
                "Facing": 10,
            },
            {
                "Source Division": "Snack",
                "Region": "WEST",
                "Period": "Jan 25",
                "Product Code": "COMPETITOR SKU",
                "Facing": 30,
            },
        ])

        rows, meta = runner.buat_tabel_sos_monthly(
            df,
            ["Source Division", "Region"],
            "REGION",
            ["Jan 25"],
            {("REGION", "WEST"): 70},
        )

        first_cells = [row[0] for row in rows]
        self.assertIn("Noodle TOTAL", first_cells)
        self.assertIn("Snack TOTAL", first_cells)
        self.assertIn("GRAND TOTAL", first_cells)

    def test_excel_export_no_division_total_row(self):
        """Per-division Excel sheets should not contain [Divisi] TOTAL rows."""
        df = self.make_dashboard_df()
        with tempfile.TemporaryDirectory() as tmp:
            output_path = runner.export_summary_excel(
                df,
                {("REGION", "WEST"): 70},
                output_dir=tmp,
                cluster_name="TestCluster",
            )
            wb = load_workbook(output_path)

        for sheet_name in [s for s in wb.sheetnames if s != "TARGETS"]:
            ws = wb[sheet_name]
            first_col_values = [
                ws.cell(row=r, column=1).value
                for r in range(1, ws.max_row + 1)
                if ws.cell(row=r, column=1).value
            ]
            division_total_rows = [
                v for v in first_col_values
                if str(v).endswith(" TOTAL") and v != "GRAND TOTAL"
            ]
            self.assertEqual(
                division_total_rows, [],
                f"Sheet '{sheet_name}' should not have [Divisi] TOTAL rows, found: {division_total_rows}"
            )

    def test_enrich_targets_adds_grand_total_per_dimension(self):
        """_enrich_targets_with_df_values should add GRAND TOTAL for each division × dimension."""
        df = self.make_dashboard_df()
        header = ['Division', 'Dimension', 'Name', 'Target']
        initial_rows = [header, ['Noodle', 'REGION', 'DEFAULT', 65.0]]

        enriched = runner._enrich_targets_with_df_values(df, initial_rows)

        enriched_names = [
            (str(row[0]), str(row[1]), str(row[2]))
            for row in enriched[1:]
        ]
        # GRAND TOTAL should exist for each division × each dimension
        for division in ['Noodle', 'Snack']:
            for dim in ['REGION', 'CHANNEL', 'ACCOUNT GELATIK', 'CATEGORY CHANNEL', 'CHANNEL-ACCOUNT']:
                self.assertIn(
                    (division, dim, 'GRAND TOTAL'),
                    enriched_names,
                    f"Missing GRAND TOTAL for {division} × {dim}"
                )

    def test_targets_sheet_contains_grand_total_entries(self):
        """TARGETS sheet in Excel output should have GRAND TOTAL rows."""
        df = self.make_dashboard_df()
        with tempfile.TemporaryDirectory() as tmp:
            output_path = runner.export_summary_excel(
                df,
                {("REGION", "WEST"): 70},
                output_dir=tmp,
                cluster_name="TestGT",
            )
            wb = load_workbook(output_path)
            ws = wb["TARGETS"]

        target_entries = []
        for r in range(2, ws.max_row + 1):
            div = ws.cell(row=r, column=1).value
            dim = ws.cell(row=r, column=2).value
            name = ws.cell(row=r, column=3).value
            if div and dim and name:
                target_entries.append((str(div), str(dim), str(name)))

        for division in ['Noodle', 'Snack']:
            for dim in ['REGION', 'CHANNEL', 'ACCOUNT GELATIK', 'CATEGORY CHANNEL']:
                self.assertIn(
                    (division, dim, 'GRAND TOTAL'),
                    target_entries,
                    f"TARGETS sheet missing GRAND TOTAL for {division} × {dim}"
                )


if __name__ == "__main__":
    unittest.main()
