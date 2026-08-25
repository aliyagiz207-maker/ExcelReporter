def calculate_kpis(df):
    # -------------------------------------------------
    # Satış ve maliyet hesaplamaları
    # -------------------------------------------------
    df = df.copy()

    df["Revenue"] = df["Quantity"] * df["UnitPrice"]
    df["Cost"] = df["Quantity"] * df["UnitCost"]
    df["Profit"] = df["Revenue"] - df["Cost"]

    # -------------------------------------------------
    # KPI hesaplamaları
    # -------------------------------------------------
    total_quantity = df["Quantity"].sum()

    total_revenue = df["Revenue"].sum()

    total_cost = df["Cost"].sum()

    total_profit = df["Profit"].sum()

    if total_revenue > 0:
        profit_margin = (
            total_profit / total_revenue
        ) * 100
    else:
        profit_margin = 0

    kpis = {
        "Total Quantity": total_quantity,
        "Total Revenue": total_revenue,
        "Total Cost": total_cost,
        "Total Profit": total_profit,
        "Profit Margin": profit_margin,
    }

    # -------------------------------------------------
    # Bölgelere göre satış özeti
    # -------------------------------------------------
    region_summary = (
        df.groupby("Region", as_index=False)[
            ["Revenue", "Cost", "Profit"]
        ]
        .sum()
        .sort_values(
            "Revenue",
            ascending=False,
        )
        .reset_index(drop=True)
    )

    region_summary["Margin"] = (
        region_summary["Profit"]
        .div(region_summary["Revenue"])
        .mul(100)
        .fillna(0)
    )

    # -------------------------------------------------
    # Ürünlere göre satış özeti
    # -------------------------------------------------
    product_summary = (
        df.groupby("Product", as_index=False)[
            ["Revenue", "Cost", "Profit"]
        ]
        .sum()
        .sort_values(
            "Revenue",
            ascending=False,
        )
        .reset_index(drop=True)
    )

    product_summary["Margin"] = (
        product_summary["Profit"]
        .div(product_summary["Revenue"])
        .mul(100)
        .fillna(0)
    )

    return kpis, region_summary, product_summary