import pandas as pd
from loguru import logger


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Veriyi temizler ve temizlenmiş DataFrame'i döndürür.
    """

    logger.info("Veri temizleme başladı.")

    original_rows = len(df)

    # Tamamen boş satırları kaldır
    df = df.dropna(how="all")

    after_dropna = len(df)

    # Tekrar eden kayıtları kaldır
    df = df.drop_duplicates()

    final_rows = len(df)

    logger.info(f"Başlangıç satır sayısı: {original_rows}")
    logger.info(f"Boş satır silinen: {original_rows - after_dropna}")
    logger.info(f"Tekrarlı kayıt silinen: {after_dropna - final_rows}")
    logger.info(f"Kalan satır sayısı: {final_rows}")

    logger.info("Veri temizleme tamamlandı.")

    return df