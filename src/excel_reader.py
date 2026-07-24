from pathlib import Path
import pandas as pd


def read_excel(file_path: Path) -> pd.DataFrame:
    """
    Excel dosyasını okuyup DataFrame olarak döndürür.
    """

    return pd.read_excel(file_path)