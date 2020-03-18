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
idf.columns

# Import ESX Production Servers
esdf = pd.read_excel(os.path.join('data', 'TLG-MOB ESX Servers v1.xlsx'), sheet_name="TLG Production Servers")
esdf.sort_values("Hostname", inplace=True)
esdf.to_excel(writer, sheet_name="ESX Prod Srv")
print(f"Prod Servers: {esdf.shape[0]} rows, {esdf.shape[1]} columns")

### VM
# Working off of the TLG prod servers tab
vfilter = '`Hardware Abstraction` == "VMGuest" & `MOTS Name` == "M-TLG"'
vmdf = esdf[esdf.eval(vfilter)].copy()
vmdf.sort_values("Hostname", inplace=True)
vmdf.to_excel(writer, sheet_name="ESX VMs")
print(f"VMGuests: {vmdf.shape[0]} rows, {vmdf.shape[1]} columns")

# Import ESX Production ESX
epdf = pd.read_excel(os.path.join('data', 'TLG-MOB ESX Servers v1.xlsx'), sheet_name="TLG Production ESX")
epdf.sort_values("Hostname", inplace=True)
epdf.to_excel(writer, sheet_name="Prod ESX Hosts")
print(f"Prod ESX: {epdf.shape[0]} rows, {epdf.shape[1]} columns")

# Combined ESX data
edf = pd.concat([epdf, esdf])
edf.sort_values("Hostname", inplace=True)
edf.to_excel(writer, sheet_name="Comb ESX")
print(f"ESX Combined: {edf.shape[0]} rows, {edf.shape[1]} columns")

# Duplicates
dups = edf[edf.duplicated('Hostname')]
print(f"ESX Duplicate Rows: {dups.shape[0]}")
dups = edf[edf.duplicated('Hostname', keep=False)]
# dups.to_excel(writer, sheet_name="ESX Dups")

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

# Label hardware abstraction
adf.loc[adf.eval('`Hardware Abstraction` == "" & `Server Model` == "Vmware virtual platform"'), 'Hardware Abstraction'] = 'Vmguest'
adf.loc[adf.eval('`Hardware Abstraction` == "" & ( `Server Model` != "Vmware virtual platform" | `Server Model` != "")'), 'Hardware Abstraction'] = 'Bare-metal'
adf.to_excel(writer, sheet_name="IEDS Comb ESX")


### VM
vdf = pd.merge(left=idf, right=vmdf, how='outer', left_on='Hostname', right_on='Hostname', suffixes=("", "_dup"))
vdf = vdf.sort_index(axis=1)
vdf.sort_values("Hostname", inplace=True)
vdf.columns = vdf.columns.str.replace('#', 'Num')
vdf = vdf[['Hostname'] + [col for col in vdf.columns if col != 'Hostname']]
for col in vdf.columns:
    if vdf[col].dtype.kind == 'O':
        vdf[col].replace(np.nan, '', regex=True, inplace=True)
for col in vdf.columns:
    if vdf[col].dtype.kind == 'O':
        vdf[col] = vdf[col].str.strip().str.capitalize()
# Label hardware abstraction
vdf.loc[vdf.eval('`Hardware Abstraction` == "" & `Server Model` == "Vmware virtual platform"'), 'Hardware Abstraction'] = 'Vmguest'
vdf.loc[vdf.eval('`Hardware Abstraction` == "" & ( `Server Model` != "Vmware virtual platform" | `Server Model` != "")'), 'Hardware Abstraction'] = 'Bare-metal'
vdf.to_excel(writer, sheet_name="IEDS VMs")


### Filter working data set


# Filter
# Note duplicate columns:  Comments_dup, Environment_dup, OS_dup
filter_str = '`Hardware Abstraction` == "Bare-metal" & \
  (`Environment` == "Production" | `Environment_dup` == "Production" | `Environment` == "" | `Environment_dup` == "") & \
  `IEDS status` == "Production" & \
  `Server Model` != "Vmware virtual platform"'

vfilter_str = '`Hardware Abstraction` == "Bare-metal" & \



fdf = adf[adf.eval(filter_str)]
print(f"Filtered Data: {fdf.shape[0]} rows, {fdf.shape[1]} columns")
# fdf = fdf[['Hostname'] + [col for col in fdf.columns if col != 'Hostname']]
fdf.to_excel(writer, sheet_name="Filter IEDS ESX")



# No HW Abstraction
# hdf = fdf[fdf.eval("`Hardware Abstraction` == ''")]
# print(f"No Hardware Abstraction: {hdf.shape[0]} rows, {hdf.shape[1]} columns")
# hdf.to_excel(writer, sheet_name="Filter No HWA")

# Non filtered data
nfdf = adf[~adf.eval(filter_str)]
print(f"Non Filtered Data: {nfdf.shape[0]} rows, {nfdf.shape[1]} columns")
nfdf.to_excel(writer, sheet_name="Non Filter IEDS ESX")


### VM
filter_str = '(`Environment` == "Production" | `Environment_dup` == "Production" | `Environment` == "" | `Environment_dup` == "") & \
  `IEDS status` == "Production"'
vfdf = vdf[vdf.eval(filter_str)]
print(f"Filtered Data: {vfdf.shape[0]} rows, {vfdf.shape[1]} columns")
vfdf.to_excel(writer, sheet_name="Filter IEDS VMs")

nvfdf = adf[~adf.eval(filter_str)]
print(f"Non Filtered Data: {nvfdf.shape[0]} rows, {nvfdf.shape[1]} columns")
nvfdf.to_excel(writer, sheet_name="Non Filter IEDS VMs")


# Target systems
targets = [
    ['DL3601s', 24, 128],
    ['DL3602s', 48, 256],
    ['DL380', 96, 1024],
    ['DL580', 128, 2048],
    ['', np.nan, np.nan],
]
tgdf = pd.DataFrame(targets, columns=['Target', 'Cores', 'Memory'])
tgdf.to_excel(writer, sheet_name="Targets")


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
    ['`Purpose` == "Customer - tra"', "DL580"],
]

rdf = pd.DataFrame(rules, columns=['Rule', 'Target'])
rdf.to_excel(writer, sheet_name="Rules")
rdf = rdf.replace(regex=r'# ', value="Num ")


def applyRules(df):

    tdf = pd.DataFrame()
    print("Applying rules: ...")
    # Iterate over rows
    for i in range(0, len(df)):
        # Select row as data frame preserving field types
        server = pd.DataFrame(df.iloc[i:i+1, :]).copy()

        # Iterate over rules
        server['Rules'] = ''
        server['Target'] = ''
        for idx, row in rdf.iterrows():
            rule = row['Rule']
            res = server.eval(rule)
            if res.iloc[0]:
                server['Target'] = rdf.iloc[idx, 1]
                server['Rules'] = server['Rules'] + str(idx) + ', '

        server = server.replace(regex=r', $', value="")

        # Calculations
        #Todo: Proper title of columns
        if server['Target'].iloc[0] != "":
            server['PctCores'] = server['Num Cores'] / tgdf[tgdf['Target'] == server['Target'].iloc[0]]['Cores'].iloc[0]
            server['PctMemory'] = server['Memory (GB)'] / tgdf[tgdf['Target'] == server['Target'].iloc[0]]['Memory'].iloc[0]
        else:
            server['PctCores'] = np.nan
            server['PctMemory'] = np.nan

        tdf = tdf.append(server)

    tdf.columns = tdf.columns.str.replace('Num', '#')
    tdf = tdf[['Hostname'] + [col for col in tdf.columns if col != 'Hostname']]
    return tdf


tfdf = applyRules(fdf)
tfdf.to_excel(writer, sheet_name="Target IEDS ESX")


###VM
tvfdf = applyRules(vfdf)
tvfdf.to_excel(writer, sheet_name="Target IEDS VMs")


writer.save()

