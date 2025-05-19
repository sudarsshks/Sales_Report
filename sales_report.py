# 📦 Import libraries
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import textwrap

# 🎯 Set style

sns.set_palette('pastel')

# 📂 Load the dataset
df = pd.read_excel("Superstore.xlsx")

# 📋 Basic info
print("Data Shape:", df.shape)
print("Missing Values:\n", df.isnull().sum())
df.drop_duplicates(inplace=True)

# 📅 Convert 'Order Date' to datetime
df['Order Date'] = pd.to_datetime(df['Order Date'])

# 📌 Feature engineering
df['Month'] = df['Order Date'].dt.to_period('M')

# 1️⃣ Total Sales & Profit
total_sales = df['Sales'].sum()
total_profit = df['Profit'].sum()
print(f"Total Sales: ₹{total_sales:,.2f}")
print(f"Total Profit: ₹{total_profit:,.2f}")

# 2️⃣ Sales by Region
region_sales = df.groupby('Region')['Sales'].sum().sort_values()

plt.figure(figsize=(8, 5))
region_sales.plot(kind='barh', color='skyblue')
plt.title('Sales by Region')
plt.xlabel('Total Sales (INR)')
plt.ylabel('Region')
plt.tight_layout()
plt.show()

# 3️⃣ Monthly Sales Trend
monthly_sales = df.groupby('Month')['Sales'].sum()

plt.figure(figsize=(10, 5))
monthly_sales.plot(marker='o', color='orange')
plt.title('Monthly Sales Trend')
plt.ylabel('Sales (INR)')
plt.xlabel('Month')
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# 4️⃣ Sales by Category
category_sales = df.groupby('Category')['Sales'].sum()

plt.figure(figsize=(6, 6))
plt.pie(category_sales, labels=category_sales.index, autopct='%1.1f%%', startangle=140)
plt.title('Sales by Category')
plt.axis('equal')
plt.tight_layout()
plt.show()

# 5️⃣ Top 10 Products by Profit
top_products = df.groupby('Product Name')['Profit'].sum().sort_values(ascending=False).head(10)
wrapped_labels = ['\n'.join(textwrap.wrap(name, 30)) for name in top_products.index]

plt.figure(figsize=(10, 6))
plt.barh(wrapped_labels, top_products.values, color='mediumseagreen')
plt.xlabel('Profit (INR)')
plt.title('Top 10 Products by Profit')
plt.gca().invert_yaxis()
plt.tight_layout()
plt.show()

df.to_csv("retail_sales_cleaned.csv", index=False)
print("Data exported successfully.")