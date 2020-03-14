import os
import pandas as pd
import xlsxwriter

# Variables
ieds_xls = "New TLG Inventory 02-28-2020.xlsx"
ieds_tab = "Source-IEDS ref-Jeff"
rules_xls = "rules.xlsx"
rules_tab = "Rules"
filter_tab = "Filter"

# Import IEDS into data frame
df = pd.read_excel(os.path.join('data', ieds_xls), sheet_name=ieds_tab)

# Remove server names that are null
df = df[df['Server Name'].notna()]

# Remove server names that are numeric
df = df[df['Server Name'].str.len().notna()]

print(f"IEDS: {len(df)} rows, {len(df.columns)} columns")

# Import filter
fdf = pd.read_excel(os.path.join('data', rules_xls), sheet_name=filter_tab)
fstr = fdf.iloc[0, 0]

# Filter data frame
df = df[df.eval(fstr)]
print(f"Filter: {len(df)} rows, {len(df.columns)} columns")

# Import rules
rdf = pd.read_excel(os.path.join('data', rules_xls), sheet_name=rules_tab)

# Update column names to work with pandas query (# as first character is no bueno)
df.columns = df.columns.str.replace('#', 'Num')
rdf = rdf.replace(regex=r'# ', value = "Num ")
# df.columns = df.columns.str.replace(' ', '_')

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
output.reset_index(level=0, inplace=True)
output = output.rename(columns={'index':'Row #'})
output['Row #'] = output['Row #'] + 2

map_tab = 'Mapping'
writer = pd.ExcelWriter(os.path.join('output', "data_analysis.xlsx"), engine='xlsxwriter')
output.to_excel(writer, sheet_name=map_tab)

workbook = writer.book
worksheet = writer.sheets[map_tab]

writer.save()

print('Done')




pd.Ex