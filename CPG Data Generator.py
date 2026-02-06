import random
import pandas as pd
from datetime import datetime, timedelta

# Helper functions
def random_upc():
    return ''.join(str(random.randint(0,9)) for _ in range(12))

def random_date_within_years(years=3):
    end = datetime.today()
    start = end - timedelta(days=365*years)
    return start + (end - start) * random.random()

def random_components():
    # 3 components summing to 100%
    a, b = random.random(), random.random()
    total = a + b + 1
    c1 = round((a/total)*100, 1)
    c2 = round((b/total)*100, 1)
    c3 = round(100 - c1 - c2, 1)
    return f"Comp1 {c1}%", f"Comp2 {c2}%", f"Comp3 {c3}%"

# Suppliers
suppliers = [
    ("AquaPure Brands", "SUP001"), ("Nature’s Pantry", "SUP002"), ("BrewMasters Co.", "SUP003"),
    ("HealthWise Labs", "SUP004"), ("Baker’s Best", "SUP005"), ("FruityBites Inc.", "SUP006"),
    ("CleanHome Goods", "SUP007"), ("SweetTreat Co.", "SUP008"), ("NutraSpread", "SUP009"),
    ("FreshFizz Beverages", "SUP010")
]

items = []
for i in range(1000):
    upc = random_upc()
    desc = f"CPG Item {i+1}"
    supplier_name, supplier_id = random.choice(suppliers)
    price = round(random.uniform(1.99, 19.99), 2)
    margin = round(random.uniform(0.05, 0.50), 2)
    cost = round(price * (1 - margin), 2)
    comp1, comp2, comp3 = random_components()
    store = random.randint(1,100)
    region = random.randint(1,10)
    sales_units = random.randint(10000,100000)
    total_sales = round(price * sales_units, 2)
    year = random.choice([2023, 2024, 2025])
    last_cost_change = random_date_within_years(3).strftime("%Y-%m-%d")

    items.append([
        upc, desc, supplier_name, supplier_id, price, margin, cost,
        comp1, comp2, comp3, store, region, sales_units,
        total_sales, year, last_cost_change
    ])

# Build dataframe
columns = [
    "UPC", "Item Description", "Supplier Name", "Supplier ID",
    "Sale Price", "Margin %", "Cost",
    "Component1", "Component2", "Component3",
    "Store #", "Region #", "Sales Units",
    "Total Sales", "Sales Year", "Last Cost Change"
]

df = pd.DataFrame(items, columns=columns)

# Export CSV
df.to_csv("cpg_consumables_1000.csv", index=False)

print("Generated 1000-item dataset: cpg_consumables_1000.csv")
print(df.head(10))
