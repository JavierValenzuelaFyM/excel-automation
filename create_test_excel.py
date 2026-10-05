import pandas as pd

enero = pd.DataFrame({
    "order_id": [1001, 1002, 1002],
    "customer": ["Ana", "Carlos", "Carlos"],
    "product": ["Laptop", "Mouse", "Mouse"],
    "quantity": [1, 2, 2],
    "price": [850, 25, 25]
})

febrero = pd.DataFrame({
    "order_id": [1003, 1004, 1005],
    "customer": ["Lucia", "Javier", None],
    "product": ["Keyboard", "Monitor", "Mouse"],
    "quantity": [1, None, 3],
    "price": [60, 220, 25]
})

marzo = pd.DataFrame({
    "order_id": [1006, 1007],
    "customer": ["Marta", "Pedro"],
    "product": ["Laptop", "Monitor"],
    "quantity": [1, 2],
    "price": [850, 220]
})

with pd.ExcelWriter("input/sales_multisheet.xlsx") as writer:
    enero.to_excel(writer, sheet_name="January", index=False)
    febrero.to_excel(writer, sheet_name="February", index=False)
    marzo.to_excel(writer, sheet_name="March", index=False)

print("Excel con varias hojas creado.")