import pandas as pd
import numpy as np

np.random.seed(42)

company_prefixes = [
    "Northwind", "BluePeak", "Vertex", "Nova", "Summit",
    "Atlas", "Orion", "Greenfield", "Silverline", "Redwood",
    "Brighton", "Crestview", "Oakridge", "Westbridge", "Clearwater"
]

company_suffixes = [
    "Retail", "Solutions", "Trading", "Supplies",
    "Group", "Technologies", "Distribution", "Services",
    "Industries", "Commerce"
]

customers = [
    f"{prefix} {suffix}"
    for prefix in company_prefixes
    for suffix in company_suffixes
]

products = [
    "Laptop Pro 14", "Wireless Mouse", "Mechanical Keyboard",
    "27-inch Monitor", "USB-C Dock", "Webcam HD",
    "External SSD", "Office Headset"
]

def create_month(start_id, rows):
    df = pd.DataFrame({
        "order_id": range(start_id, start_id + rows),
        "customer": np.random.choice(customers, rows),
        "product": np.random.choice(products, rows),
        "quantity": np.random.randint(1, 8, rows),
        "price": np.random.choice(
            [19.99, 34.50, 59.90, 89.99, 149.00, 229.00, 899.00],
            rows
        )
    })

    # Introducir valores ausentes
    missing_customer = np.random.choice(df.index, 12, replace=False)
    missing_quantity = np.random.choice(df.index, 8, replace=False)

    df.loc[missing_customer, "customer"] = None
    df.loc[missing_quantity, "quantity"] = np.nan

    # Introducir duplicados
    duplicates = df.sample(20, random_state=42)

    return pd.concat([df, duplicates], ignore_index=True)


january = create_month(10000, 500)
february = create_month(20000, 500)
march = create_month(30000, 500)

with pd.ExcelWriter("input/sales_multisheet.xlsx") as writer:
    january.to_excel(writer, sheet_name="January", index=False)
    february.to_excel(writer, sheet_name="February", index=False)
    march.to_excel(writer, sheet_name="March", index=False)

print("Professional demo dataset created.")