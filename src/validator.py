import pandas as pd

REQUIRED_COLUMNS = [
    "Date",
    "Product",
    "Region",
    "Quantity",
    "UnitPrice",
    "UnitCost"
]


def validate_dataframe(df: pd.DataFrame) -> None:
    """
    Excel dosyasındaki veriyi kontrol eder.
    Sorun varsa hata verir.
    """

    # Dosya boş mu?
    if df.empty:
        raise ValueError("Excel dosyası boş.")

    # Eksik sütun var mı?
    missing = []

    for column in REQUIRED_COLUMNS:
        if column not in df.columns:
            missing.append(column)

    if missing:
        raise ValueError(
            f"Eksik sütun(lar): {', '.join(missing)}"
        )