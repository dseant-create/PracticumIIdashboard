"""
Generate supplier sales data from CPG data and export as JSON for dashboard
"""

import random
from datetime import date, timedelta
from pathlib import Path
import json

import numpy as np
import pandas as pd

# Configuration
N_ROWS = 1000
SEED = 42
random.seed(SEED)
np.random.seed(SEED)

SUPPLIERS = [
    ("Procter & Gamble", "USA"),
    ("Unilever", "USA"),
    ("Kimberly-Clark", "USA"),
    ("Nestlé", "USA"),
    ("PepsiCo", "USA"),
    ("The Coca-Cola Company", "USA"),
    ("Colgate-Palmolive", "USA"),
    ("Johnson & Johnson", "USA"),
    ("General Mills", "USA"),
    ("Kraft Heinz", "USA"),
    ("Mondelez International", "USA"),
    ("The Clorox Company", "USA"),
    ("Reckitt", "USA"),
    ("SC Johnson", "USA"),
    ("Church & Dwight", "USA"),
    ("Henkel", "USA"),
    ("Frito-Lay", "USA"),
    ("Conagra Brands", "USA"),
    ("Tyson Foods", "USA"),
    ("Ferrero", "USA"),
    ("Grupo Bimbo", "Mexico"),
    ("Gruma", "Mexico"),
    ("Maruchan", "China"),
    ("ITC Limited", "India"),
    ("Dabur", "India"),
    ("Godrej Consumer Products", "India"),
]

PRODUCT_TEMPLATES = [
    ("Grocery", "Cereal", ["Cheerios", "Corn Flakes", "Frosted Flakes", "Honey Oats"], ["12 oz", "18 oz", "24 oz"]),
    ("Grocery", "Snack Bars", ["Nature Valley", "Kind", "Quaker", "Clif"], ["6 ct", "12 ct", "18 ct"]),
    ("Grocery", "Potato Chips", ["Lay's", "Ruffles", "Pringles", "Kettle"], ["5 oz", "8 oz", "12 oz"]),
    ("Grocery", "Pasta", ["Barilla", "Ronzoni", "De Cecco", "Store Brand"], ["12 oz", "16 oz"]),
    ("Grocery", "Pasta Sauce", ["Rao's", "Prego", "Classico", "Store Brand"], ["24 oz", "32 oz"]),
    ("Grocery", "Coffee", ["Folgers", "Dunkin'", "Starbucks", "Store Brand"], ["12 oz", "20 oz", "32 oz"]),
    ("Grocery", "Tea", ["Lipton", "Twinings", "Tazo", "Store Brand"], ["20 ct", "40 ct"]),
    ("Grocery", "Soda", ["Coca-Cola", "Pepsi", "Sprite", "Dr Pepper"], ["12 pk 12 oz", "6 pk 16.9 oz"]),
    ("Grocery", "Bottled Water", ["Aquafina", "Dasani", "Smartwater", "Store Brand"], ["24 pk", "12 pk"]),
    ("Grocery", "Yogurt", ["Chobani", "Dannon", "Oikos", "Store Brand"], ["4 pk", "6 oz"]),
    ("Household", "Paper Towels", ["Bounty", "Scott", "Viva", "Store Brand"], ["6 rolls", "8 rolls", "12 rolls"]),
    ("Household", "Toilet Paper", ["Charmin", "Cottonelle", "Scott", "Store Brand"], ["6 rolls", "12 rolls", "18 rolls"]),
    ("Household", "Trash Bags", ["Hefty", "Glad", "Store Brand", "Simplehuman"], ["20 ct", "40 ct", "60 ct"]),
    ("Household", "Laundry Detergent", ["Tide", "Gain", "Persil", "Store Brand"], ["92 oz", "150 oz", "64 loads"]),
    ("Household", "Dish Soap", ["Dawn", "Palmolive", "Seventh Generation", "Store Brand"], ["24 oz", "32 oz"]),
    ("Household", "Disinfecting Wipes", ["Clorox", "Lysol", "Store Brand"], ["75 ct", "110 ct"]),
    ("Personal Care", "Toothpaste", ["Crest", "Colgate", "Sensodyne", "Store Brand"], ["4.8 oz", "6 oz"]),
    ("Personal Care", "Shampoo", ["Head & Shoulders", "Pantene", "Dove", "Store Brand"], ["12 oz", "20 oz"]),
    ("Personal Care", "Body Wash", ["Dove", "Old Spice", "Olay", "Store Brand"], ["18 oz", "24 oz"]),
    ("Personal Care", "Deodorant", ["Old Spice", "Secret", "Dove", "Degree"], ["2.6 oz", "3 oz"]),
    ("Baby", "Diapers", ["Pampers", "Huggies", "Luvs", "Store Brand"], ["Size 3 72 ct", "Size 4 84 ct", "Size 5 64 ct"]),
    ("Baby", "Baby Wipes", ["Pampers", "Huggies", "WaterWipes", "Store Brand"], ["3 pk 168 ct", "1 pk 72 ct"]),
    ("Pet", "Dog Food", ["Purina", "Pedigree", "Blue Buffalo", "Store Brand"], ["15 lb", "30 lb", "40 lb"]),
    ("Pet", "Cat Litter", ["Tidy Cats", "Arm & Hammer", "Fresh Step", "Store Brand"], ["14 lb", "20 lb", "35 lb"]),
]

def base_price_for(product_type: str) -> float:
    if product_type == "Diapers":
        return round(random.uniform(18.99, 44.99), 2)
    if product_type == "Dog Food":
        return round(random.uniform(14.99, 49.99), 2)
    if product_type == "Cat Litter":
        return round(random.uniform(9.99, 24.99), 2)
    if product_type == "Laundry Detergent":
        return round(random.uniform(8.99, 26.99), 2)
    if product_type in {"Paper Towels", "Toilet Paper", "Trash Bags", "Disinfecting Wipes"}:
        return round(random.uniform(5.49, 22.99), 2)
    if product_type in {"Soda", "Bottled Water"}:
        return round(random.uniform(3.99, 12.99), 2)
    return round(random.uniform(1.49, 14.99), 2)

def main():
    """Generate supplier sales data and export as JSON"""
    rows = []
    
    for i in range(N_ROWS):
        _, product_type, brands, sizes = random.choice(PRODUCT_TEMPLATES)
        brand = random.choice(brands)
        size = random.choice(sizes)
        
        flavor = ""
        if product_type in {"Yogurt", "Snack Bars", "Cereal", "Tea"}:
            flavor = random.choice(["", " - Vanilla", " - Strawberry", " - Honey", " - Mixed Berry", " - Original"])
        
        item_desc = f"{brand} {product_type}{flavor}, {size}"
        
        # Supplier
        supplier_name, supplier_country = random.choice(SUPPLIERS)
        
        # Base price
        p23 = base_price_for(product_type)
        
        # Units per year
        u23 = random.randint(10_000, 100_000)
        u24 = int(round(u23 * random.uniform(0.80, 1.20)))
        u25 = int(round(u24 * random.uniform(0.80, 1.20)))
        u24 = min(100_000, max(10_000, u24))
        u25 = min(100_000, max(10_000, u25))
        
        # Prices
        r1, r2 = random.uniform(0.02, 0.08), random.uniform(0.02, 0.08)
        p24 = round(p23 * (1 + r1), 2)
        p25 = round(p24 * (1 + r2), 2)
        
        # Total sales
        s23 = round(u23 * p23, 2)
        s24 = round(u24 * p24, 2)
        s25 = round(u25 * p25, 2)
        
        rows.append({
            "supplier": supplier_name,
            "country": supplier_country,
            "product": item_desc,
            "sales_2023": float(s23),
            "sales_2024": float(s24),
            "sales_2025": float(s25),
        })
    
    # Aggregate by supplier and year
    df = pd.DataFrame(rows)
    
    # Group by supplier and sum sales
    supplier_sales = {
        "2023": {},
        "2024": {},
        "2025": {}
    }
    
    for idx, row in df.iterrows():
        supplier = row["supplier"]
        country = row["country"]
        
        if supplier not in supplier_sales["2023"]:
            supplier_sales["2023"][supplier] = 0
            supplier_sales["2024"][supplier] = 0
            supplier_sales["2025"][supplier] = 0
        
        supplier_sales["2023"][supplier] += row["sales_2023"]
        supplier_sales["2024"][supplier] += row["sales_2024"]
        supplier_sales["2025"][supplier] += row["sales_2025"]
    
    # Format data for dashboard
    dashboard_data = []
    for supplier in sorted(supplier_sales["2023"].keys()):
        dashboard_data.append({
            "supplier": supplier,
            "y2023": round(supplier_sales["2023"][supplier], 2),
            "y2024": round(supplier_sales["2024"][supplier], 2),
            "y2025": round(supplier_sales["2025"][supplier], 2),
        })
    
    # Write JSON file
    output_file = Path("supplier_sales_data.json")
    with open(output_file, "w") as f:
        json.dump(dashboard_data, f, indent=2)
    
    print(f"Generated supplier sales data: {output_file}")
    print(f"Total suppliers: {len(dashboard_data)}")
    print(f"Total records processed: {len(df)}")

if __name__ == "__main__":
    main()
