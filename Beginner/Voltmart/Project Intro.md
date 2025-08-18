For this project, and perhaps the (TO BE CONTINUED)

import pandas as pd
import re
all_sheets = pd.ExcelFile(r'C:\Users\moham\OneDrive\Desktop\Python\DA\Voltmart\voltmart_multi_sheet.xlsx')
# CUSTOMER SHEET
df_cust = all_sheets.parse('Customers')

# Select columns to strip leading/trailing spaces
strip_cols = df_cust.select_dtypes(include=['object']).columns
# Strip leading/trailing spaces overall
df_cust[strip_cols] = df_cust[strip_cols].apply(lambda x: x.str.strip())

# Normalize strings
title_cols = ['Full Name','City','Country','Churned']

df_cust[title_cols] = df_cust[title_cols].apply(lambda x: x.str.title())

# Signup Date
df_cust['Signup Date'] = pd.to_datetime(df_cust['Signup Date'], format='mixed', errors='coerce')
df_cust['Signup Date'] = df_cust['Signup Date'].dt.strftime('%Y-%m-%d')

# Country
country_list = {'Mexico':'Mexico', 'Usa':'United States', 'Canada':'Canada', 'Mex':'Mexico', 'U.S.':'United States', 'United States':'United States', 'Can':'Canada'}
df_cust['Country']=df_cust['Country'].map(country_list)

# Churned
churned = {'Yes':'Yes','No':'No','Y':'Yes','N':'No'}
df_cust['Churned']=df_cust['Churned'].map(churned)
# PRODUCT SHEET

df_prod = all_sheets.parse('Products')

# Select columns to strip leading/trailing spaces
strip_cols_prod = df_prod.select_dtypes(include=['object']).columns
# Strip leading/trailing spaces overall
df_prod[strip_cols_prod] = df_prod[strip_cols_prod].apply(lambda x: x.str.strip())

# Category
category_list= {'LAPTOP':'Laptop', 'Laptops':'Laptop', 'smartphone':'Smartphone', 'Smartphone':'Smartphone', 'Tablet':'Tablet',
       'TABLET':'Tablet', 'Accessory':'Accessory', 'Accessorys':'Accessory', 'Tv':'TV', 'TV':'TV', 'accessory':'Accessory',
       'laptop':'Laptop', 'SMARTPHONE':'Smartphone'}
df_prod['Category'] = df_prod['Category'].map(category_list)

# List price
df_prod['List Price ($)'] = df_prod['List Price ($)'].str.replace(r'[$]|USD','',regex=True)
#df_prod
# ORDER SHEET

df_orders = all_sheets.parse('Orders')

# Select columns to strip leading/trailing spaces
strip_cols_ord = ['Order ID', 'Customer ID', 'Product ID', 'Order Date',
       'Order Amount ($)']
# Strip leading/trailing spaces overall
df_orders[strip_cols_ord] = df_orders[strip_cols_ord].apply(lambda x: x.str.strip())

# Order date
df_orders['Order Date'] = pd.to_datetime(df_orders['Order Date'], format='mixed', errors='coerce')
df_orders['Order Date'] = df_orders['Order Date'].dt.strftime('%Y-%m-%d')

# Quantity
quantity_list= {3:3,'one':2,'4':4,'three':3, 1:1,'3':3, 5:5,'four':4, 2:2, 4:4, 'five':5, 'two':2,'2':2, '1':1, '5':5}
df_orders['Quantity'] = df_orders['Quantity'].map(quantity_list)

# Order amount
df_orders['Order Amount ($)'] = df_orders['Order Amount ($)'].str.replace(r'[$,]|USD','',regex=True)

# Product ID
df_orders['Product ID'] = df_orders['Product ID'].str.replace(r'^(.{4})[-_ ]*', r'\1-',regex=True)
# -- fixing Product by matching to product in Order sheet --
# Function to normalize Product ID and Customer ID in 

def normalized_code(code, prefix):
    match = re.search(r'(\d+)', str(code))
    if match:
        number = int(match.group(1))
        return f'{prefix}-{number:04d}'
    return None

# Apply normalized codes
df_orders['Product ID'] = df_orders['Product ID'].apply(lambda x: normalized_code(x, 'PROD'))
df_orders['Customer ID'] = df_orders['Customer ID'].apply(lambda x: normalized_code(x,'CUST'))
df_orders['Order ID'] = df_orders['Order ID'].apply(lambda x: normalized_code(x,'ORD'))

df_orders
# PAYMENT SHEET

df_pay = all_sheets.parse('Payments')

# Select columns to strip leading/trailing spaces
strip_cols_pay = df_pay.select_dtypes(include=['object']).columns
# Strip leading/trailing spaces overall
df_pay[strip_cols_pay] = df_pay[strip_cols_pay].apply(lambda x:x.str.strip())

# Normalize Order/Payment column
def normalized_code(code, prefix):
    match = re.search(r'(\d+)', str(code))
    if match:
        number = int(match.group(1))
        return f'{prefix}-{number:04d}'
    return None

# Apply function
df_pay['Order ID'] = df_pay['Order ID'].apply(lambda x: normalized_code(x,'ORD'))
df_pay['Payment ID'] = df_pay['Payment ID'].apply(lambda x: normalized_code(x,'PAY'))

# Payment Date
df_pay['Payment Date'] = pd.to_datetime(df_pay['Payment Date'], format='mixed', errors='coerce')
df_pay['Payment Date'] = df_pay['Payment Date'].dt.strftime('%Y-%m-%d')

# Payment method
payment_list = {'PayPal':'PayPal', 'paypal':'PayPal', 'bank transfer':'Bank Transfer', 'CC':'Credit Card', 'creditcard':'Credit Card',
       'Bank Transfer':'Bank Transfer', 'Credit Card':'Credit Card'}
df_pay['Payment Method']= df_pay['Payment Method'].map(payment_list).str.title()

# Paid Amount
def convert_currency(val):
    if pd.isnull(val):
        return None

    val = str(val).strip()

    # Remove currency symbols or letters
    val = re.sub(r'[^\d,.\s-]', '', val)

    # Check for European format
    if re.search(r'[\d\s.]+,\d{2}$', val):
        val = val.replace(' ', '').replace('.', '').replace(',', '.')
    else:
        val = val.replace(',', '')

    try:
        return float(val)
    except ValueError:
        return None

df_pay['Paid Amount ($)'] = df_pay['Paid Amount ($)'].apply(convert_currency)
# SUPPORT TICKETS SHEET

df_supp = all_sheets.parse('Support_Tickets')

# Strip leading/trailing spaces
strip_col_supp = df_supp.select_dtypes(include=['object']).columns
# Apply stripping
df_supp[strip_col_supp]=df_supp[strip_col_supp].apply(lambda x:x.str.strip())

# Normalize Customer ID
def normalized_code(code, prefix):
    match = re.search(r'(\d+)', str(code))
    if match:
        number = int(match.group(1))
        return f'{prefix}-{number:04d}'
    return None

df_supp['Customer ID']=df_supp['Customer ID'].apply(lambda x:normalized_code(x,'CUST'))
df_supp['Ticket ID']=df_supp['Ticket ID'].apply(lambda x:normalized_code(x,'TICK'))

# Ticket date
df_supp['Ticket Date'] = pd.to_datetime(df_supp['Ticket Date'], format='mixed', errors='coerce')
df_supp['Ticket Date'] = df_supp['Ticket Date'].dt.strftime('%Y-%m-%d')

# Category
cat_list_supp = {'DELIV':'Delivery', 'Billing':'Billing', 'BILL':'Billing', 'warranty':'Warranty', 'account':'Account', 'delivery':'Delivery',
       'returns':'Return', 'tech':'Technical', 'Account':'Account', 'RET':'Return', 'Technical':'Technical', 'Warranty':'Warranty',
       'Delivery':'Delivery'}
df_supp['Category'] = df_supp['Category'].map(cat_list_supp)

# Description
df_supp['Description'] = df_supp['Description'].str.replace('.','')
df_supp[['Description','Note']] = df_supp['Description'].str.split(r';',expand=True)
df_supp['Note'] = df_supp['Note'].str.strip().str.capitalize()
df_supp = df_supp.fillna('N/A')

# Rating
rating_list = {'5':5, '***':3, '4':4, '****':4, '3':3, 'three':3, '2':2, 'one':1, '*****':5, 'two':2,'1':1, 'five':5, 'four':4, '**':2, '*':1}
df_supp['Rating'] = df_supp['Rating'].map(rating_list)
# CONVERT TO EXCEL
output_path = r'C:\Users\moham\OneDrive\Desktop\Python\DA\Voltmart\test.xlsx'

with pd.ExcelWriter(output_path) as writer:
    df_orders.to_excel(writer, sheet_name='Order', index=False)
    df_supp.to_excel(writer, sheet_name='Support Ticket', index=False)
    df_prod.to_excel(writer, sheet_name='Product', index=False)
    df_pay.to_excel(writer, sheet_name='Payment', index=False)
    df_cust.to_excel(writer, sheet_name='Customer', index=False)
