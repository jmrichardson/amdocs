import os
import pandas as pd
import xlsxwriter
import re
import numpy as np
# import datacompy

# Variables
ieds_xls = "New TLG Inventory 02-28-2020.xlsx"
ieds_tab = "Source-IEDS ref-Jeff"

esx_xls = "TLG-MOB ESX Servers v1.xlsx"
esx_server_tab = "TLG Production Servers"
esx_parent_tab = "TLG Production ESX"

parms_xls = "parameters.xlsx"
rules_tab = "Rules"
filter_tab = "Servers"

###

# Import IEDS data
idf = pd.read_excel(os.path.join('data', ieds_xls), sheet_name=ieds_tab)
# Remove server names that are null
idf = idf[idf['Server Name'].notna()]
# Remove server names that are numeric
idf = idf[idf['Server Name'].str.len().notna()]
# Rename column
idf.columns = idf.columns.str.replace('Server Name', 'Hostname')
idf.sort_values("Hostname", inplace=True)

print(f"IEDS: {len(idf)} rows, {len(idf.columns)} columns")

# Import ESX data
esdf = pd.read_excel(os.path.join('data', esx_xls), sheet_name=esx_server_tab)
esdf.sort_values("Hostname", inplace=True)
epdf = pd.read_excel(os.path.join('data', esx_xls), sheet_name=esx_parent_tab)
epdf.sort_values("Hostname", inplace=True)

# Merge IEDS and ESX data
df = pd.concat([epdf, esdf])
df.drop_duplicates(subset="Hostname", keep=False, inplace=True)
df = pd.merge(left=idf, right=df, how='outer', left_on='Hostname', right_on='Hostname', suffixes=("", "_ESX_Dup"))

# Sort, drop duplicate rows by hostname, reorder
df.sort_values("Hostname", inplace=True)
df = df.reindex(sorted(df.columns), axis=1)
df.drop_duplicates(subset="Hostname", keep=False, inplace=True)
df = df[['Hostname'] + [col for col in df.columns if col != 'Hostname']]
# df = df.replace(np.nan, '', regex=True)

# Update column names to work with pandas query (# as first character is no bueno)
df.columns = df.columns.str.replace('#', 'Num')

# Import filter
fdf = pd.read_excel(os.path.join('data', parms_xls), sheet_name=filter_tab)
fstr = fdf.iloc[0, 0]

# Filter data frame
df = df[df.eval(fstr)]
print(f"Filter: {len(df)} rows, {len(df.columns)} columns")

# Import rules
rdf = pd.read_excel(os.path.join('data', parms_xls), sheet_name=rules_tab)
rdf = rdf.replace(regex=r'# ', value = "Num ")


output = pd.DataFrame()

print("Applying rules: ...")
# Iterate over rows
for i in range(0, len(df)):
    # Select row as data frame preserving field types
    server = pd.DataFrame(df.iloc[i:i+1, :])

    # Iterate over rules
    server['Rules'] = ''
    for idx, row in rdf.iterrows():
        rule = row['Rule']

        rdx = idx + 2
        res = server.eval(rule)

        if res.iloc[0]:
            server['Target'] = rdf.iloc[idx, 1]
            server['Rules'] = server['Rules'] + str(rdx) + ', '

    server = server.replace(regex=r', $', value="")
    output = output.append(server)

output.columns = output.columns.str.replace('Num', '#')
output = output.dropna(axis=1, how='all')

# output.reset_index(level=0, inplace=True)
# output = output.rename(columns={'index':'Row #'})
# output['Row #'] = output['Row #'] + 2

map_tab = 'Mapping'



writer = pd.ExcelWriter(os.path.join('output', "data_analysis.xlsx"), engine='xlsxwriter')
output.to_excel(writer, sheet_name=map_tab, startrow=1, startcol=0, header=False, index=False)

workbook = writer.book
worksheet = writer.sheets[map_tab]

header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})

for col_num, value in enumerate(output.columns.values):
    worksheet.write(0, col_num, value, header_format)

def get_col_widths(dataframe):
    # First we find the maximum length of the index column
    # idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [max([len(str(s))+3 for s in dataframe[col].values] + [len(col)+3]) for col in dataframe.columns]



for i, width in enumerate(get_col_widths(output)):
    worksheet.set_column(i, i, width)




writer.save()

print('Done')




