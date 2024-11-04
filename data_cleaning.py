import pandas as pd

#load the dataset
data = pd.read_excel('Online Retail.xlsx', engine='openpyxl')


#check for duplicates
def categorize_unit_price(unit_price):
    if unit_price > 10:
        return 'High Value'
    elif 5 < unit_price <= 10:
        return 'Medium Value'
    else:
        return 'Low Value'


# Identifying duplicates based on InvoiceNo, StockCode, CustomerID, and UnitPrice
duplicates_check = (
    data[data.duplicated(subset=['InvoiceNo', 'StockCode', 'CustomerID', 'UnitPrice'], keep=False)]
    .assign(Description=lambda df: df['UnitPrice'].apply(categorize_unit_price))
)

#Remove duplicate from the DataFrame

unique_data = data.drop_duplicates(subset=['InvoiceNo', 'StockCode', 'CustomerID', 'UnitPrice']).copy()

#save unique data without duplicate
unique_data.to_excel('Online_retail_unique.xlsx', index=False)


#Detect Outliers using IQR method
def detect_outliers_iqr(df, column):
    Q1 = df[column].quantile(0.25)
    Q3 = df[column].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    return df[(df[column] < lower_bound) | (df[column] > upper_bound)]


# Identify outliers
outliers_quantity = detect_outliers_iqr(unique_data, 'Quantity')
outliers_unit_price = detect_outliers_iqr(unique_data, 'UnitPrice')

# Handle Outliers by Imputing with Median
# Replace outliers in Quantity and UnitPrice with median values
quantity_median = unique_data['Quantity'].median()
unit_price_median = unique_data['UnitPrice'].median()

unique_data.loc[unique_data['Quantity'].isin(outliers_quantity['Quantity'].values), 'Quantity'] = quantity_median
unique_data.loc[unique_data['UnitPrice'].isin(outliers_unit_price['UnitPrice'].values), 'UnitPrice'] = unit_price_median

unique_data.to_excel('Online_retail_unique.xlsx', index=False)
# Convert InvoiceDate to datetime and handle errors
unique_data['InvoiceDate'] = pd.to_datetime(unique_data['InvoiceDate'], errors='coerce')

# Format 'InvoiceDate' to show as DD/MM/YYYY H:M
unique_data['InvoiceDate'] = unique_data['InvoiceDate'].dt.strftime('%d/%m/%Y %H:%M')

# new column TransactionType to keep track of sale or return
unique_data.loc[:, 'TransactionType'] = unique_data['Quantity'].apply(lambda x: 'Return' if x < 0 else 'Sale')
unique_data.to_excel('Online_retail_unique_updated.xlsx', index=False)

# Reorder the columns to place TransactionType next to Quantity
cols = list(unique_data.columns)
if 'Quantity' in cols and 'TransactionType' in cols:
    qty_index = cols.index('Quantity')
    cols = cols[:qty_index + 1] + ['TransactionType'] + cols[qty_index + 1:]
    unique_data = unique_data[cols]

# Filling missing values for 'Description' with "No Description" and 'CustomerID' with 0
unique_data = unique_data.assign(
    Description=unique_data['Description'].fillna("No Description"),
    CustomerID=unique_data['CustomerID'].fillna(0)
)

# Drop the duplicate TransactionType column (keeping the first occurrence)-I created 2
unique_data = unique_data.loc[:, ~unique_data.columns.duplicated(keep='first')]

# Check and handle missing values
missing_data_report = unique_data.isnull().sum()

# Data Type Correction: Ensure all columns have the correct data types
unique_data['CustomerID'] = unique_data['CustomerID'].astype(int)
unique_data['Quantity'] = unique_data['Quantity'].astype(int)
unique_data['UnitPrice'] = unique_data['UnitPrice'].astype(float)
unique_data['InvoiceDate'] = unique_data['InvoiceDate'].astype(str)

#Final save of cleaned data
unique_data.to_excel('Online_retail_unique.xlsx', index=False)

#print("\nData Types:")
#print(unique_data.dtypes)
