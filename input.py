import os
import pandas as pd
import xlsxwriter
import re
import numpy as np
# import datacompy


# Create data import progress spreadsheet
writer = pd.ExcelWriter(os.path.join('output', 'data_import.xlsx'), engine='xlsxwriter')

# Import IEDS data
idf = pd.read_excel(os.path.join('data', 'New TLG Inventory 02-28-2020.xlsx'), sheet_name="Source-IEDS ref-Jeff")
# Remove server names that are null
idf = idf[idf['Server Name'].notna()]
# Remove server names that are numeric
idf = idf[idf['Server Name'].str.len().notna()]
# Rename column
idf.columns = idf.columns.str.replace('Server Name', 'Hostname')
idf.sort_values("Hostname", inplace=True)
# Export IEDS data
idf.to_excel(writer, sheet_name="IEDS")
print(f"IEDS: {idf.shape[0]} rows, {idf.shape[1]} columns")

# Import ESX Production Servers
esdf = pd.read_excel(os.path.join('data', 'TLG-MOB ESX Servers v1.xlsx'), sheet_name="TLG Production Servers")
esdf.sort_values("Hostname", inplace=True)
esdf.to_excel(writer, sheet_name="Prod Servers")
print(f"Prod Servers: {esdf.shape[0]} rows, {esdf.shape[1]} columns")

# Import ESX Production ESX
epdf = pd.read_excel(os.path.join('data', 'TLG-MOB ESX Servers v1.xlsx'), sheet_name="TLG Production ESX")
epdf.sort_values("Hostname", inplace=True)
epdf.to_excel(writer, sheet_name="Prod ESX")
print(f"Prod ESX: {epdf.shape[0]} rows, {epdf.shape[1]} columns")

# Combined ESX data
edf = pd.concat([epdf, esdf])
edf.sort_values("Hostname", inplace=True)
edf.to_excel(writer, sheet_name="ESX")
print(f"ESX Combined: {edf.shape[0]} rows, {edf.shape[1]} columns")

# Duplicates
dups = edf[edf.duplicated('Hostname')]
print(f"ESX Duplicate Rows: {dups.shape[0]}")
dups = edf[edf.duplicated('Hostname', keep=False)]
dups.to_excel(writer, sheet_name="ESX Dups")

## TODO - Do we drop duplicate rows?
# edf.drop_duplicates(subset="Hostname", keep=False, inplace=True)

adf = pd.merge(left=idf, right=edf, how='outer', left_on='Hostname', right_on='Hostname', suffixes=("", "_dup"))
adf = adf.sort_index(axis=1)
adf.sort_values("Hostname", inplace=True)
print(f"Duplicate Columns: {', '.join(adf.columns[adf.columns.str.endswith('_dup')].tolist())}")
# Update columns to work with pandas query
adf.columns = adf.columns.str.replace('#', 'Num')
print(f"IEDS ESX: {adf.shape[0]} rows, {adf.shape[1]} columns")
dups = adf[adf.duplicated('Hostname')]
print(f"IEDS ESX Merged Duplicate Rows: {dups.shape[0]}")
# Move hostname col to first column
adf = adf[['Hostname'] + [col for col in adf.columns if col != 'Hostname']]
# Repalce nans with empty string for non float columns
for col in adf.columns:
    if adf[col].dtype.kind == 'O':
        adf[col].replace(np.nan, '', regex=True, inplace=True)

# Capitalize first letter and lowercase the rest of each word (consistent data)
for col in adf.columns:
    if adf[col].dtype.kind == 'O':
        adf[col] = adf[col].str.strip().str.capitalize()

adf.to_excel(writer, sheet_name="IEDS ESX")
# df = df.replace(np.nan, '', regex=True)


### Filter working data set


# Filter
# Note duplicate columns:  Comments_dup, Environment_dup, OS_dup
filter_str = '`Hardware Abstraction` != "Vmguest" & \
  (`Environment` == "Production" | `Environment_dup` == "Production" | `Environment` == "" | `Environment_dup` == "") & \
  `IEDS status` == "Production" & \
  `Server Model` != "VMWARE VIRTUAL PLATFORM"'


fdf = adf[adf.eval(filter_str)]
print(f"Filtered Data: {fdf.shape[0]} rows, {fdf.shape[1]} columns")
fdf = fdf[['Hostname'] + [col for col in fdf.columns if col != 'Hostname']]
fdf.to_excel(writer, sheet_name="Filter Data")


### Rules


# Rules
# Each rule is applied to every row
# The last matching rule takes precedence
# Data words are fist letter capital rest undercase
# One to one mapping
rules = [
    ['`Memory (GB)` < 1000',  "DL3601s"],
    ['`# Cores` < 24', "DL3601s"],
    ['`Memory (GB)` >= 1000', "DL3602s"],
    ['`# Cores` >= 24', "DL3602s"],
    ['`# Cores` >= 48', "DL580"],
    ['`Purpose` == "Customer"', "DL580"],
    ['`Purpose` == "Customer - tra"', "DL580"],
    ['`Purpose` == "Archangel"', "Ignore"],
    ['`Server Model` == "Vmware virtual platform"', "Ignore"],
    ['`Purpose` == "Media server"', "Ignore"],
    ['`Purpose` == "Fbf"', "Ignore"],
]

rdf = pd.DataFrame(rules, columns=['Rule', 'Target'])
rdf.to_excel(writer, sheet_name="Rules")
rdf = rdf.replace(regex=r'# ', value="Num ")

# Apply Rules
tdf = pd.DataFrame()
print("Applying rules: ...")
# Iterate over rows
for i in range(0, len(fdf)):
    # Select row as data frame preserving field types
    server = pd.DataFrame(fdf.iloc[i:i+1, :])

    # Iterate over rules
    server['Rules'] = ''
    for idx, row in rdf.iterrows():
        rule = row['Rule']

        res = server.eval(rule)

        if res.iloc[0]:
            server['Target'] = rdf.iloc[idx, 1]
            server['Rules'] = server['Rules'] + str(idx) + ', '

    server = server.replace(regex=r', $', value="")
    tdf = tdf.append(server)

tdf.columns = tdf.columns.str.replace('Num', '#')
tdf = tdf[['Hostname'] + [col for col in tdf.columns if col != 'Hostname']]
tdf.to_excel(writer, sheet_name="Target")
writer.save()

