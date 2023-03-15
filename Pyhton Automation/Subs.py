import pandas as pd
import openpyxl as opx

filepath1 = 'C:/Users/lindb/Documents/CBL/Renewals/ALIV Renewals - Subscription Export.xlsx'
filepath2 = 'C:/Users/lindb/Documents/CBL/Renewals/ALIV Renewals - Opportunity Insert success.xlsx'


df = pd.read_excel(io=filepath1)
df2 = pd.read_excel(io=filepath2)

# Config the Subs export file
# Delete all rows where the column "Subscription Status" is not "Active" or "Setup" or "Suspended"
df = df[(df['Subscription Status'] == 'Active') | (df['Subscription Status'] == 'Setup') | (df['Subscription Status'] == 'Suspended')]
#Rename Account to Account ID


#CHANGE RECURRING CHARGE TO SALE PRICE 
df.rename(columns={'Recurring Charge': 'Sales Price'}, inplace=True)
df2.rename(columns={'ID': 'Opportunity ID'}, inplace=True)
# Create Key1 column on opp insert success which consists of Account and Type
df2['Key1'] = df2['Account ID'] + df2['Type']
# Create Key1 column on subs export which consists of Account and Contract Type
df['Key1'] = df['Account'] + df['Contract Type']

#Outer_join = pd.merge(df, df2, on='Key1', how='outer')

df3 = df.merge(df2, on="Key1")

df3['Key2'] = df3['Plan__r.Product ID'] + df3['Department']

# Create a Quantity Column to count how many have the same exact department and plan__r.product.id pairing.using this excel formula =COUNTIF(Key2:Key2,Key2)
df3['Quantity'] = df3.groupby('Key2')['Key2'].transform('count')

# remove duplicates from df3 based on column key 2
df3.drop_duplicates(subset='Key2', keep='first', inplace=True)

# Delete all columns except: Account, ID, Sales Price, Quantity, Department, Plan__r.Product ID, Plan__r.Product Family Detail, Plan__r.Product Family
df3 = df3[['Account', 'Opportunity ID', 'Sales Price', 'Quantity', 'Department', 'Plan__r.Product ID', 'Plan__r.Product Family Detail', 'Plan__r.Product Family']]

# Create a new column "SIM" = "Existing Postpaid" & Column "Customer Sale Type" = "EC Renewal"
df3.insert(0, 'SIM', 'Existing Postpaid')
df3.insert(1, 'Customer Sale Type', 'EC - Renewal')

# Rename Plan__r.Product ID to Product ID, Plan__r.Product Family Detail to Product Family Detail, Plan__r.Product Family to Product Family
df3.rename(columns={'Plan__r.Product ID': 'Product ID', 'Plan__r.Product Family Detail': 'Product Family Detail', 'Plan__r.Product Family': 'Product Family'}, inplace=True)

# reorder the columns to Account, ID, SIM, Customer Sale Type, Sales Price, Quantity, Department, Product ID, Product Family Detail, Product Family
df3 = df3[['Account', 'Opportunity ID', 'SIM', 'Customer Sale Type', 'Sales Price', 'Quantity', 'Department', 'Product ID', 'Product Family Detail', 'Product Family']]



df3.to_csv('C:/Users/lindb/Documents/CBL/Renewals/renewalOLIinsert.csv', index=False)

# createa a new column "match" and use column Key1 on subs export to vlookup Key1 on opp insert success
#df['match'] = df['Key1'].map(df2.set_index('Key1')['Key1'])

#df.to_csv('C:/Users/lindb/Documents/CBL/subsexport2.csv', index=False)