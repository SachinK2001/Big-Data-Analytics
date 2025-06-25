import pandas as pd
import random
from datetime import datetime, timedelta
import xlsxwriter

# --------------------------
# Step 1: Generate Dataset
# --------------------------

products = [
    ("Laptop", "Electronics", 60000),
    ("Smartphone", "Electronics", 30000),
    ("T-shirt", "Clothing", 500),
    ("Shoes", "Clothing", 1200),
    ("Blender", "Home Appliances", 2500),
    ("Notebook", "Stationery", 50),
    ("Pen", "Stationery", 20),
    ("Washing Machine", "Home Appliances", 18000),
    ("Headphones", "Electronics", 1500),
    ("Backpack", "Accessories", 800)
]

data = []
for i in range(1, 51):
    transaction_id = f"T{i:03}"
    customer_id = f"C{random.randint(1, 15):03}"
    product, category, price = random.choice(products)
    quantity = random.randint(1, 5)
    date = datetime(2025, 6, 1) + timedelta(days=random.randint(0, 27))
    data.append([transaction_id, customer_id, product, category, quantity, price, date.strftime("%Y-%m-%d")])

df = pd.DataFrame(data, columns=["TransactionID", "CustomerID", "Product", "Category", "Quantity", "Price", "Date"])
df["Revenue"] = df["Quantity"] * df["Price"]

# --------------------------
# Step 2: Map Phase
# --------------------------

map_product_quantity = df.groupby("Product")["Quantity"].sum().reset_index()
map_category_revenue = df.groupby("Category")["Revenue"].sum().reset_index()
map_customer_frequency = df.groupby("CustomerID")["TransactionID"].count().reset_index()
map_customer_frequency.columns = ["CustomerID", "Frequency"]

# --------------------------
# Step 3: Top 5 for Charts
# --------------------------

top5_products = map_product_quantity.sort_values(by="Quantity", ascending=False).head(5)
top5_customers = map_customer_frequency.sort_values(by="Frequency", ascending=False).head(5)

# --------------------------
# Step 4: Save to Excel
# --------------------------

workbook = xlsxwriter.Workbook("Big_Data_Mini_Project_Report_with_Charts.xlsx")

# Sheet 1: Full dataset
worksheet_data = workbook.add_worksheet("Dataset")
for col_num, col_name in enumerate(df.columns):
    worksheet_data.write(0, col_num, col_name)
    for row_num, value in enumerate(df[col_name], start=1):
        worksheet_data.write(row_num, col_num, value)

# Sheet 2: Top Products
worksheet_prod = workbook.add_worksheet("Top_Products")
worksheet_prod.write_column("A2", top5_products["Product"])
worksheet_prod.write_column("B2", top5_products["Quantity"])
worksheet_prod.write("A1", "Product")
worksheet_prod.write("B1", "Quantity")

chart1 = workbook.add_chart({'type': 'column'})
chart1.add_series({
    'name': 'Top 5 Products',
    'categories': ['Top_Products', 1, 0, 5, 0],
    'values':     ['Top_Products', 1, 1, 5, 1],
})
chart1.set_title({'name': 'Top 5 Selling Products'})
worksheet_prod.insert_chart('D2', chart1)

# Sheet 3: Revenue by Category
worksheet_cat = workbook.add_worksheet("Revenue_by_Category")
worksheet_cat.write_column("A2", map_category_revenue["Category"])
worksheet_cat.write_column("B2", map_category_revenue["Revenue"])
worksheet_cat.write("A1", "Category")
worksheet_cat.write("B1", "Revenue")

chart2 = workbook.add_chart({'type': 'pie'})
chart2.add_series({
    'name': 'Revenue by Category',
    'categories': ['Revenue_by_Category', 1, 0, len(map_category_revenue), 0],
    'values':     ['Revenue_by_Category', 1, 1, len(map_category_revenue), 1],
})
chart2.set_title({'name': 'Revenue by Category'})
worksheet_cat.insert_chart('D2', chart2)

# Sheet 4: Top Customers
worksheet_cust = workbook.add_worksheet("Top_Customers")
worksheet_cust.write_column("A2", top5_customers["CustomerID"])
worksheet_cust.write_column("B2", top5_customers["Frequency"])
worksheet_cust.write("A1", "CustomerID")
worksheet_cust.write("B1", "Frequency")

chart3 = workbook.add_chart({'type': 'column'})
chart3.add_series({
    'name': 'Top 5 Customers',
    'categories': ['Top_Customers', 1, 0, 5, 0],
    'values':     ['Top_Customers', 1, 1, 5, 1],
})
chart3.set_title({'name': 'Top 5 Customers by Frequency'})
worksheet_cust.insert_chart('D2', chart3)

# Save Excel
workbook.close()
print("Excel report generated successfully: Big_Data_Mini_Project_Report_with_Charts.xlsx")

