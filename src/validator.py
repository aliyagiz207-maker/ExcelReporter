import pandas as pd

REQUIRED_COLUMNS = [
    "Date",
    "Product",
    "Region",
    "Quantity",
    "UnitPrice",
    "UnitCost"
]

NUMERIC_COLUMNS = [
    "Quantity",
    "UnitPrice",
    "UnitCost",
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

    # Sayısal olması gereken sütunlar gerçekten sayısal mı?
    non_numeric = []

    for column in NUMERIC_COLUMNS:
        if not pd.api.types.is_numeric_dtype(df[column]):
            non_numeric.append(column)

    if non_numeric:
        raise ValueError(
            f"Sayısal olması gereken sütun(lar) sayısal değil: "
            f"{', '.join(non_numeric)}"
        )

    # Date sütunu tarihe çevrilebiliyor mu?
    parsed_dates = pd.to_datetime(df["Date"], errors="coerce")

    if parsed_dates.isna().any():
        invalid_count = int(parsed_dates.isna().sum())
        raise ValueError(
            f"Geçersiz tarih içeren {invalid_count} satır bulundu."
        )

    # Negatif değer kontrolü
    negative_columns = []

    for column in NUMERIC_COLUMNS:
        if (df[column] < 0).any():
            negative_columns.append(column)

    if negative_columns:
        raise ValueError(
            f"Negatif değer içeren sütun(lar): "
            f"{', '.join(negative_columns)}"
        )