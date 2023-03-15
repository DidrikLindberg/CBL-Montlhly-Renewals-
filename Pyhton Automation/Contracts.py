import pandas as pd
import openpyxl as opx

filepath = 'C:/Users/lindb/Documents/CBL/May Renewals.xlsx'


df = pd.read_excel(io=filepath)

#Delete Column "Contract Start Date"
del df['Contract Start Date']

# Add Column "Contract Start Date" and populate with data from column "Contract End Date" + 1 day and add it right after column "Contract End Date"
df.insert(6, 'Contract Start Date', df['Contract End Date'] + pd.Timedelta(days=1))
# Change the format to mm/dd/yyyy
df['Contract Start Date'] = df['Contract Start Date'].dt.strftime('%m/%d/%Y')

# Delete Column "Contract End Date"
del df['Contract End Date']

# Change the name of column Contract Term (Months) to Contract Term
df.rename(columns={'Contract Term (months)': 'Contract Term'}, inplace=True)

# Add Column "Contract Name" and use this excel formula to populate it, using the column headers as indicators: = Account Name & " " & Contract Type & " Contract " & Term & " Months " & TEXT(Contract Start Date, "mm/dd/yyyy") Contract Start date is not in datetime format, so can not use dt/strftime
df.insert(3, 'Contract Name', df['Account Name'] + " " + df['Contract Type'] + " Contract " + df['Contract Term'].astype(str) + " Months " + df['Contract Start Date'])

# Add Column after Contract Name "Len" and apply the len formula to the Contract Name column
df.insert(4, 'Len', df['Contract Name'].str.len())




#df.insert(3, 'Contract Name', df['Account Name'] + " " + df['Contract Type'] + " Contract " + df['Contract Term'].astype(str) + " Months " + df['Contract Start Date'].dt.strftime('%m/%d/%Y'))

# Add colulmn "Owner Expiration Notice" and populate each row with 15
df.insert(9, 'Owner Expiration Notice', 15)

# Change the values in column "Status" to "Draft"
df['Status'] = 'Draft'  

# Add Column "Renewal" set values equal to "TRUE"
df.insert(11, 'Renewal', 'TRUE')

# Change name of Case Safe ID to Account ID
df.rename(columns={'Case Safe ID': 'Account ID'}, inplace=True)
# Delete Columns "Case Safe ID" and "Contract Number" and "Case Safe Contract ID" and "Corporate Contract Status"
del df['Contract Number']
del df['Case Safe Contract ID']
del df['Corporate Contract Status']



# save a copy of the file as a csv UTF 8
df.to_csv('C:/Users/lindb/Documents/CBL/RenewalsContracts.csv', index=False)




