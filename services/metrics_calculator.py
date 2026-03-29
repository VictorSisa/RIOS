# Metrics Calculator

This script includes comprehensive calculations for several key retail metrics:

## 1. GMROI (Gross Margin Return on Investment)

```python
def calculate_gmroi(sales, cost_of_goods_sold, inventory):
    gross_margin = sales - cost_of_goods_sold
    gmroi = gross_margin / inventory if inventory != 0 else 0
    return gmroi
```

## 2. Weeks of Supply

```python
def calculate_weeks_of_supply(inventory, weekly_sales):
    weeks_of_supply = inventory / weekly_sales if weekly_sales != 0 else 0
    return weeks_of_supply
```

## 3. Sell-through Rate

```python
def calculate_sell_through(sales, inventory):
    sell_through = sales / inventory if inventory != 0 else 0
    return sell_through
```

## 4. Year-over-Year (YoY) Trends

```python
def calculate_yoy(current_year_sales, prior_year_sales):
    yoy_change = (current_year_sales - prior_year_sales) / prior_year_sales * 100 if prior_year_sales != 0 else 0
    return yoy_change
```

# Example Usage
if __name__ == '__main__':
    sales = 10000  # Example sales figure
    cogs = 7000   # Cost of goods sold
    inventory = 5000  # Current inventory

    gmroi = calculate_gmroi(sales, cogs, inventory)
    print(f"GMROI: {gmroi}")  # GMROI Result

    weekly_sales = 200  # Example weekly sales
    weeks_of_supply = calculate_weeks_of_supply(inventory, weekly_sales)
    print(f"Weeks of Supply: {weeks_of_supply}")  # Weeks of Supply Result

    sell_through = calculate_sell_through(sales, inventory)
    print(f"Sell-through Rate: {sell_through}")  # Sell-through Result

    prior_year_sales = 8000  # Example prior year sales
    yoy = calculate_yoy(sales, prior_year_sales)
    print(f"Year-over-Year Change: {yoy}%")  # YoY Result
