import os
import pandas as pd
import pandasql

# Variables
ieds_xls = "New TLG Inventory 02-28-2020.xlsx"
ieds_tab = "Source-IEDS ref-Jeff"
rules_xls = "rules.xlsx"
rules_tab = "Rules"


# Import IEDS into data frame
df = pd.read_excel(os.path.join('data', ieds_xls), sheet_name=ieds_tab)

# Remove server names that are null
df = df[df['Server Name'].notna()]

# Remove server names that are numeric
df = df[df['Server Name'].str.len().notna()]

# Import rules
rdf = pd.read_excel(os.path.join('data', rules_xls), sheet_name=rules_tab)

### Filter examples
# df[df['# Cores'] < 48]
# pandasql.sqldf("SELECT * FROM df WHERE '# Cores' < 128.0")
# df.query('`City` == "HOOVER"')

# Update column names to work with pandas query (# as first character is no bueno)
df.columns = df.columns.str.replace('#', 'Num')
# df.columns = df.columns.str.replace(' ', '_')

# Filters
# f1 = '`Num Cores` <= 28'
# f1 = "df['# Cores' <= 128]"
# f1 = "'Num Cores' <= 24"
# f2 = '`Server Name` == "zlt11022"'

f1 = "`Num Cores` * .5 <= 8"


# Iterate over rows
for i in range(0, len(df)):
    # Select row as data frame preserving field types
    row = df.iloc[i:i+1, :]
    res = row.eval(f1)
    if res.iloc[0]:
        print("ya")
    else:
        print("na")



a = pd.DataFrame(row).T

# Iterate over df
for index, row in df.iterrows():
    print(row.dtypes)
    row = pd.DataFrame(row).T
    row.eval(f1)



new =


len(j.query(f1))

df['Num Cores'] * 0.5

df.query('`Num Cores`')

df['Num Cores']





