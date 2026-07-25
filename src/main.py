from report_generator import generate_report

from pathlib import Path

from excel_reader import read_excel
from validator import validate_dataframe
from data_cleaner import clean_data
from kpi_calculator import calculate_kpis
from logger import setup_logger


def main():
    logger = setup_logger()

    logger.info("Program başlatıldı.")

    project_root = Path(__file__).resolve().parent.parent

    excel_file = project_root / "data" / "input" / "sales_sample.xlsx"

    df = read_excel(excel_file)

    logger.info("Excel dosyası başarıyla okundu.")

    validate_dataframe(df)

    df = clean_data(df)

    kpis = calculate_kpis(df)
    generate_report(kpis)

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