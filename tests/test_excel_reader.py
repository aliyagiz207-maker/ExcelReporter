import pandas as pd
import pytest

from src.excel_reader import read_file, read_excel_folder


def test_read_csv(tmp_path):
    csv_file = tmp_path / "test.csv"

    csv_file.write_text(
        "Date;Product;Region;Quantity;UnitPrice;UnitCost\n"
        "01.01.2026;Brake Pad;İzmir;10;100;60\n",
        encoding="utf-8",
    )

    result = read_file(csv_file)

    assert isinstance(result, pd.DataFrame)
    assert len(result) == 1
    assert list(result.columns) == [
        "Date",
        "Product",
        "Region",
        "Quantity",
        "UnitPrice",
        "UnitCost",
    ]


def test_read_excel(tmp_path):
    excel_file = tmp_path / "test.xlsx"

    df = pd.DataFrame({
        "Date": ["2026-01-01"],
        "Product": ["Brake Pad"],
        "Region": ["İzmir"],
        "Quantity": [10],
        "UnitPrice": [100],
        "UnitCost": [60],
    })

    df.to_excel(excel_file, index=False)

    result = read_file(excel_file)

    assert isinstance(result, pd.DataFrame)
    assert len(result) == 1
    assert list(result.columns) == list(df.columns)


def test_read_excel_folder_combines_files(tmp_path):
    first_file = tmp_path / "january.csv"
    second_file = tmp_path / "february.csv"

    content = (
        "Date;Product;Region;Quantity;UnitPrice;UnitCost\n"
        "01.01.2026;Brake Pad;İzmir;10;100;60\n"
    )

    first_file.write_text(content, encoding="utf-8")
    second_file.write_text(
        content.replace("01.01.2026", "01.02.2026"),
        encoding="utf-8",
    )

    result = read_excel_folder(tmp_path)

    assert len(result) == 2


def test_read_excel_folder_selected_months(tmp_path):
    january = tmp_path / "january.csv"
    february = tmp_path / "february.csv"

    content = (
        "Date;Product;Region;Quantity;UnitPrice;UnitCost\n"
        "01.01.2026;Brake Pad;İzmir;10;100;60\n"
    )

    january.write_text(content, encoding="utf-8")
    february.write_text(content, encoding="utf-8")

    result = read_excel_folder(
        tmp_path,
        selected_months=["january"],
    )

    assert len(result) == 1


def test_read_excel_folder_without_files(tmp_path):
    with pytest.raises(FileNotFoundError):
        read_excel_folder(tmp_path)