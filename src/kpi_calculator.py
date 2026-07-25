import pandas as pd


def calculate_kpis(df: pd.DataFrame) -> dict:
    """
    Temel KPI'ları hesaplar ve sözlük (dict) olarak döndürür.
    """

    total_quantity = df["Quantity"].sum()

    total_revenue = (df["Quantity"] * df["UnitPrice"]).sum()

    total_cost = (df["Quantity"] * df["UnitCost"]).sum()

    total_profit = total_revenue - total_cost

    if total_revenue == 0:
        profit_margin = 0
    else:
        profit_margin = (total_profit / total_revenue) * 100

    return {
        "Total Quantity": total_quantity,
        "Total Revenue": total_revenue,
        "Total Cost": total_cost,
        "Total Profit": total_profit,
        "Profit Margin": profit_margin
    }