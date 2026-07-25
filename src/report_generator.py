from pathlib import Path

from openpyxl import load_workbook


def generate_report(kpis: dict):
    project_root = Path(__file__).resolve().parent.parent

    template_file = project_root / "templates" / "report_template.xlsx"
    output_file = project_root / "data" / "output" / "Report.xlsx"

    workbook = load_workbook(template_file)
    sheet = workbook.active

    sheet["B3"] = kpis["Total Quantity"]
    sheet["B4"] = kpis["Total Revenue"]
    sheet["B5"] = kpis["Total Cost"]
    sheet["B6"] = kpis["Total Profit"]
    sheet["B7"] = kpis["Profit Margin"] / 100

    sheet["B4"].number_format = "₺#,##0"
    sheet["B5"].number_format = "₺#,##0"
    sheet["B6"].number_format = "₺#,##0"
    sheet["B7"].number_format = "0.00%"

    workbook.save(output_file)