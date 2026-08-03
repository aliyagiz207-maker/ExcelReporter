import argparse

from config_loader import load_config
from file_manager import archive_report
from pathlib import Path
from pdf_report import generate_pdf_report
from excel_reader import read_excel_folder
from validator import validate_dataframe
from data_cleaner import clean_data
from kpi_calculator import calculate_kpis
from report_generator import generate_report
from logger import setup_logger


def main():
    parser = argparse.ArgumentParser(
    description="Excel Reporter")

    parser.add_argument(
    "--month",
    nargs="*",
    help="Read only selected Excel files (without .xlsx)")

    args = parser.parse_args()

    logger = setup_logger()

    logger.info("Program başlatıldı.")

    project_root = Path(__file__).resolve().parent.parent
    config = load_config()

    input_folder = project_root / "data" / "input"

    df = read_excel_folder(input_folder,selected_months=args.month)

    logger.info("Excel dosyaları başarıyla okundu.")

    validate_dataframe(df)

    df = clean_data(df)

    kpis, region_summary, product_summary = calculate_kpis(df)

    generate_report(kpis, region_summary, product_summary, df, config)
    generate_pdf_report(kpis, config)
    report_file = project_root / config["output_file"]
    archive_folder = project_root / "data" / "archive"

    archive_report(report_file, archive_folder)

    print("\n=== İlk 5 Satır ===")
    print(df.head())

    print("\n=== Veri Bilgisi ===")
    df.info()

    print("\n=== KPI Results ===")
    for key, value in kpis.items():
        if isinstance(value, float):
            print(f"{key}: {value:.2f}")
        else:
            print(f"{key}: {value}")

    logger.info("Program başarıyla tamamlandı.")


if __name__ == "__main__":
    main()