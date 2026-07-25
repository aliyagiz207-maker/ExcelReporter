from pathlib import Path

from excel_reader import read_excel
from validator import validate_dataframe
from logger import setup_logger
from data_cleaner import clean_data

def main():
    logger = setup_logger()

    logger.info("Program başlatıldı.")

    project_root = Path(__file__).resolve().parent.parent

    excel_file = project_root / "data" / "input" / "sales_sample.xlsx"

    df = read_excel(excel_file)

    logger.info("Excel dosyası başarıyla okundu.")

    validate_dataframe(df)
    df = clean_data(df)

    print("\n=== İlk 5 Satır ===")
    print(df.head())

    print("\n=== Veri Bilgisi ===")
    df.info()

    logger.info("Program başarıyla tamamlandı.")


if __name__ == "__main__":
    main()