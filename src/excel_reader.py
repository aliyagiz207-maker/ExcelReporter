from pathlib import Path

import pandas as pd


def read_excel(file_path: Path) -> pd.DataFrame:
    """
    Tek bir Excel dosyasını okuyup DataFrame olarak döndürür.
    """
    return pd.read_excel(file_path)


def read_excel_folder(
    folder_path: Path,
    selected_months: list[str] | None = None
) -> pd.DataFrame:
    """
    Klasördeki Excel dosyalarını okuyup tek DataFrame'e birleştirir.

    selected_months örnek:
    ["january", "march"]
    """

    excel_files = sorted(folder_path.glob("*.xlsx"))

    if selected_months:
        selected_months = [m.lower() for m in selected_months]

        excel_files = [
            file for file in excel_files
            if file.stem.lower() in selected_months
        ]

    if not excel_files:
        raise FileNotFoundError(
            "Seçilen kriterlere uygun Excel dosyası bulunamadı."
        )

    dataframes = []

    for excel_file in excel_files:
        print(f"Reading: {excel_file.name}")
        df = pd.read_excel(excel_file)
        dataframes.append(df)

    combined_df = pd.concat(dataframes, ignore_index=True)

    return combined_df