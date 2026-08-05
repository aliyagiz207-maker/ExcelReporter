from pathlib import Path

import pandas as pd


def read_file(file_path: Path) -> pd.DataFrame:
    """
    Excel veya CSV dosyasını okuyup DataFrame olarak döndürür.
    """

    if file_path.suffix.lower() == ".csv":
        return pd.read_csv(
        file_path,
        sep=";",
        dayfirst=True,
        parse_dates=["Date"],
    )

    return pd.read_excel(file_path)


def read_excel_folder(
    folder_path: Path,
    selected_months: list[str] | None = None,
) -> pd.DataFrame:
    """
    Klasördeki Excel dosyalarını okuyup tek DataFrame'e birleştirir.

    selected_months örnek:
    ["january", "march"]
    """

    files = sorted(
        list(folder_path.glob("*.xlsx")) 
        + list(folder_path.glob("*.csv"))
)

    if selected_months:
        selected_months = [m.lower() for m in selected_months]

        files = [
            file 
            for file in files
            if file.stem.lower() in selected_months
]

    if not files:
        raise FileNotFoundError(
    "Seçilen kriterlere uygun dosya bulunamadı."
)

    dataframes = []

    for file in files:
        print(f"Reading: {file.name}")
        df = read_file(file)
        dataframes.append(df)

    combined_df = pd.concat(dataframes, ignore_index=True)

    return combined_df