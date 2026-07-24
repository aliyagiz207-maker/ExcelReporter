from pathlib import Path

from excel_reader import read_excel
from validator import validate_dataframe

def main():
    project_root = Path(__file__).resolve().parent.parent

    excel_file = project_root / "data" / "input" / "sales_sample.xlsx"

    df = read_excel(excel_file)

    validate_dataframe(df)

    print("\n=== İlk 5 Satır ===")
    print(df.head())

    print("\n=== Veri Bilgisi ===")
    df.info()


if __name__ == "__main__":
    main()