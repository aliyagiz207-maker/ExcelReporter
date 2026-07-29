def calculate_kpis(df):
    # KPI hesaplamaları
    total_quantity = df["Quantity"].sum()

    total_revenue = (df["Quantity"] * df["UnitPrice"]).sum()

    total_cost = (df["Quantity"] * df["UnitCost"]).sum()

    total_profit = total_revenue - total_cost

    profit_margin = (total_profit / total_revenue) * 100

    kpis = {
        "Total Quantity": total_quantity,
        "Total Revenue": total_revenue,
        "Total Cost": total_cost,
        "Total Profit": total_profit,
        "Profit Margin": profit_margin,
    }

    # Bölgelere göre gelir özeti
    region_summary = (
        df.assign(Revenue=df["Quantity"] * df["UnitPrice"])
        .groupby("Region", as_index=False)["Revenue"]
        .sum()
        .sort_values("Revenue", ascending=False)
    )

    # Ürünlere göre gelir özeti
    product_summary = (
        df.assign(Revenue=df["Quantity"] * df["UnitPrice"])
        .groupby("Product", as_index=False)["Revenue"]
        .sum()
        .sort_values("Revenue", ascending=False)
    )

    return kpis, region_summary, product_summary