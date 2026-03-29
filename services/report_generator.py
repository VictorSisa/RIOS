lines.append(
    f"  {i}. {cat} | "
    f"Sales: ${data['sales']:,.0f} | YoY: {data['avg_growth']:+.1f}% | "
    f"Inventory: {data['inv_change']:+.1f}%"
)