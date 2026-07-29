from pathlib import Path

import pandas as pd


def read_excel(file_path: Path) -> pd.DataFrame:
    """
    Tek bir Excel dosyasını okuyup DataFrame olarak döndürür.
    """
    return pd.read_excel(file_path)


def read_excel_folder(folder_path: Path) -> pd.DataFrame:
    """
    Klasördeki tüm Excel dosyalarını okuyup tek DataFrame'e birleştirir.
    """

    excel_files = sorted(folder_path.glob("*.xlsx"))

    if not excel_files:
        raise FileNotFoundError(
            f"'{folder_path}' klasöründe hiç Excel dosyası bulunamadı."
        )

    dataframes = []

    for excel_file in excel_files:
        print(f"Reading: {excel_file.name}")
        df = pd.read_excel(excel_file)
        dataframes.append(df)

    combined_df = pd.concat(dataframes, ignore_index=True)

    return combined_df