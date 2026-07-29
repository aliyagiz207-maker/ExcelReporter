from pathlib import Path

from excel_reader import read_excel_folder
from validator import validate_dataframe
from data_cleaner import clean_data
from kpi_calculator import calculate_kpis
from report_generator import generate_report
from logger import setup_logger


def main():
    logger = setup_logger()

    logger.info("Program başlatıldı.")

    project_root = Path(__file__).resolve().parent.parent

    input_folder = project_root / "data" / "input"

    df = read_excel_folder(input_folder)

    logger.info("Excel dosyaları başarıyla okundu.")

    validate_dataframe(df)

    df = clean_data(df)

    kpis, region_summary, product_summary = calculate_kpis(df)

    generate_report(kpis, region_summary, product_summary, df)

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