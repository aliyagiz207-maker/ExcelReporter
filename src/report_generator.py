from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference


def generate_report(kpis, region_summary, product_summary, df):
    project_root = Path(__file__).resolve().parent.parent

    template_file = project_root / "templates" / "report_template.xlsx"
    output_file = project_root / "data" / "output" / "Report.xlsx"

    workbook = load_workbook(template_file)
    sheet = workbook["Dashboard"]

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Dashboard başlığı
    sheet["A1"] = "Excel Reporter Dashboard"
    sheet["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    sheet["A1"].fill = PatternFill(
        fill_type="solid",
        start_color="1F4E78",
        end_color="1F4E78",
    )
    sheet["A1"].alignment = Alignment(horizontal="center")

# KPI
    row = 3

    for key, value in kpis.items():
        sheet[f"A{row}"] = key

        if isinstance(value, float):
            if "Margin" in key:
                sheet[f"B{row}"] = value / 100
                sheet[f"B{row}"].number_format = "0.00%"
            else:
                sheet[f"B{row}"] = value
                sheet[f"B{row}"].number_format = "₺#,##0.00"
        else:
            sheet[f"B{row}"] = value

            if "Quantity" not in key:
                sheet[f"B{row}"].number_format = "₺#,##0"

        sheet[f"A{row}"].border = thin_border
        sheet[f"B{row}"].border = thin_border

        row += 1

    sheet.column_dimensions["A"].width = 24
    sheet.column_dimensions["B"].width = 18

# Region Summary
    sheet["D2"] = "Revenue by Region"
    sheet["D2"].font = Font(size=14, bold=True)

    sheet["D3"] = "Region"
    sheet["E3"] = "Revenue"

    sheet["D3"].font = Font(bold=True)
    sheet["E3"].font = Font(bold=True)

    start_row = 4

    for _, row_data in region_summary.iterrows():
        sheet[f"D{start_row}"] = row_data["Region"]
        sheet[f"E{start_row}"] = row_data["Revenue"]
        sheet[f"E{start_row}"].number_format = "₺#,##0"

        start_row += 1

    sheet.column_dimensions["D"].width = 18
    sheet.column_dimensions["E"].width = 18

# Grafik
    chart = BarChart()

    data = Reference(
        sheet,
        min_col=5,
        min_row=3,
        max_row=start_row - 1,
    )

    categories = Reference(
        sheet,
        min_col=4,
        min_row=4,
        max_row=start_row - 1,
    )

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    chart.title = "Revenue by Region"
    chart.y_axis.title = "Revenue"
    chart.x_axis.title = "Region"

    chart.width = 12
    chart.height = 8

    sheet.add_chart(chart, "G3")

    # Top Products
    sheet["D11"] = "Top Products"
    sheet["D11"].font = Font(size=14, bold=True)

    sheet["D12"] = "Product"
    sheet["E12"] = "Revenue"

    sheet["D12"].font = Font(bold=True)
    sheet["E12"].font = Font(bold=True)

    product_row = 13

    for _, row_data in product_summary.iterrows():
        sheet[f"D{product_row}"] = row_data["Product"]
        sheet[f"E{product_row}"] = row_data["Revenue"]
        sheet[f"E{product_row}"].number_format = "₺#,##0"

        product_row += 1

# Detail Data sayfası
    if "Detail Data" in workbook.sheetnames:
        del workbook["Detail Data"]

    detail_sheet = workbook.create_sheet("Detail Data")

    for row_data in dataframe_to_rows(df, index=False, header=True):
        detail_sheet.append(row_data)

# Başlık biçimlendirme
    for cell in detail_sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(
            fill_type="solid",
            start_color="1F4E78",
            end_color="1F4E78",
        )
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

# Sütun genişlikleri
    for column_cells in detail_sheet.columns:
        length = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in column_cells
        )
        detail_sheet.column_dimensions[
            column_cells[0].column_letter
        ].width = length + 3

    workbook.save(output_file)