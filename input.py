import os
import pandas as pd
import xlsxwriter
import re
import numpy as np
# import datacompy


# Import data
df = pd.read_excel(os.path.join('data', 'Exported file.xlsx'))
# Remove server names that are null
print(f"Dataframe: {df.shape[0]} rows, {df.shape[1]} columns")

cpu_ratio = 1
df['Cons Cores'] = np.where(df["Server model"].str.contains("Gen9"),
                                 df["Present State Cores"],
                                 df["Present State Cores"] * cpu_ratio)


# Target systems
# 24 DIMMS in 360
# 48 DIMMS in 580
targets = [
    ['DL360-10-2-48-768', 48, 768],
    ['DL360-10-2-48-1536', 48, 1536],
    ['DL580-10-4-96-1536', 96, 1536],
    ['DL580-10-4-96-3072', 96, 3072],
    ['', np.nan, np.nan],
]
tgdf = pd.DataFrame(targets, columns=['Target', 'Cores', 'Memory'])

### Rules
# Each rule is applied to every row
# The last matching rule takes precedence
# Data words are fist letter capital rest undercase
# One to one mapping
rules = [
    ['s["Present State RAM GB"] <= 768 & s["Cons Cores"] <= 48', "DL360-10-2-48-768"],
    ['(s["Present State RAM GB"] > 768 & s["Present State RAM GB"] < 1536) & s["Cons Cores"] <= 48', "DL360-10-2-48-1536"],
    ['s["Cons Cores"] > 48 & s["Present State RAM GB"] <= 1536', "DL580-10-4-96-1536"],
    ['s["Present State RAM GB"] > 1536', "DL580-10-4-96-3072"],
    ['s["Server Purpose"].str.contains("Customer")', "DL580-10-4-96-3072"],
]
rdf = pd.DataFrame(rules, columns=['Rule', 'Target'])

# rule = 'df["Present State RAM GB"] <= 768 & df["Present State Cores"] <= 48'
# res = pd.eval(rule, engine='python')
# res = pd.eval(rule, engine='python')

tdf = pd.DataFrame()
print("Applying rules: ...")
# Iterate over rows
for i in range(0, len(df)):
    # Select row as data frame preserving field types
    s = pd.DataFrame(df.iloc[i:i+1, :]).copy()

    # Iterate over rules
    s['Rules'] = ''
    s['Target'] = ''
    for idx, row in rdf.iterrows():
        rule = row['Rule']
        res = pd.eval(rule, engine='python')
        if res.iloc[0]:
            s['Target'] = rdf.iloc[idx, 1]
            s['Rules'] = s['Rules'] + str(idx) + ', '

    s = s.replace(regex=r', $', value="")

    # Calculations
    #Todo: Proper title of columns
    if s['Target'].iloc[0] != "":
        s['PctCores'] = s['Cons Cores'] / tgdf[tgdf['Target'] == s['Target'].iloc[0]]['Cores'].iloc[0]
        s['PctMemory'] = s['Present State RAM GB'] / tgdf[tgdf['Target'] == s['Target'].iloc[0]]['Memory'].iloc[0]
    else:
        s['PctCores'] = np.nan
        s['PctMemory'] = np.nan

    tdf = tdf.append(s)






writer = pd.ExcelWriter(os.path.join('output', 'targets.xlsx'), engine='xlsxwriter')
workbook = writer.book
tdf.to_excel(writer, sheet_name="Servers")
worksheet = writer.sheets["Servers"]

hide_columns = ['PrimaryKey', 'CreationTimestamp', 'CreatedBy', 'ModificationTimestamp',
       'ModifiedBy', 'Server serial number', 'Vendor', 'Notes',
       'Markets', 'Markets by line',
       'MOTS Name', 'MOTS ID',  'Instance', 'OS',
       'OS Version', 'OS Major Version', 'Data Center', 'Hardware Abstraction',
       'Cluster Name', 'PartSurver URL',
       'Comments', 'Non_TLG MOTS', '~EOSL', 'Numbering',
       'App Dependencies', 'Add Date', 'Number of supported markets',
       'Number of targeted DCs', 'Targeted DCs', 'ESX host file foreign key',
       'Parent ESX host', 'ESX Host Count of VMs', 'ESX Host Hosted Cores',
       'ESX Host Hosted CPUs', 'ESX Host Hosted RAM',
       'ESX Host Hosted Storage', 'ESX Host Targeted Data Centers',
       'ESX Host Targeted Data Center Count',
       'Unique server name to count for present state',
       '∑Present State CPUs', 'Present State Cores_with_HT',
       '∑Present State cores',
       '∑Present State RAM GB', '∑Storage GB',
       'CPU by 2', 'Current State', 'Consolidation method',
       'g ESX host CPU overhead', 'g ESX host RAM overhead',
       'g ESX host Storage overhead', 'Future State', 'Future State CPU',
       'Future State Cores', 'Future State RAM', 'Future State Storage',
       'Future State 1_1', 'DR needed', 'Future State CPU DR',
       'Future State Cores DR', 'Future State RAM DR',
       'Future State Storage DR', 'Future State 1_1 DR', 'Target Data Center',
       'Cores per physical CPU', 'Count of servers', 'Migration method',
       'Current configuration from catalog',
       'Current configuration performance figure',
       'Replacement configuration from catalog',
       'Replacement configuration performance figure', 'Replacement fraction',
       'Consolidation group foreign key', 'Consolidation group',
       'HT supported', 'Working set',
       'List of markets by DBs_and_Apps', 'Number of markets by DB_and_apps',
       'Count of DBs_and_Apps', 'Fraction of cores Stin']

for i in range(1, len(tdf.columns)):
    col = tdf.columns[i-1]
    width = max([len(str(s)) for s in tdf[col].values] + [len(col)])
    if col in hide_columns:
        worksheet.set_column(i, i, width, None, {'hidden': True})
    else:
        worksheet.set_column(i, i, width)

tgdf.to_excel(writer, sheet_name="Targets")
rdf.to_excel(writer, sheet_name="Rules")

pt = pd.pivot_table(tdf,
                    index=["Target"],
                    values=["System name"],
                    aggfunc={"System name": len},
                    margins=True,
                    margins_name="Total",
                    dropna=False)
pt.to_excel(writer, sheet_name="TargetCount")

pt = pd.pivot_table(tdf,
                    index=["Server Purpose", "Present State Cores", "Present State RAM GB"],
                    columns=["Target"],
                    values=["System name"],
                    aggfunc={"System name": len},
                    margins=True,
                    margins_name="Total",
                    dropna=True)
pt.to_excel(writer, sheet_name="ServerPurposeCoresMem-Target")


pt = pd.pivot_table(tdf,
                    index=['App Tier', 'Server Purpose'],
                    columns=["Target"],
                    values=["System name"],
                    aggfunc={"System name": len},
                    margins=True,
                    margins_name="Total",
                    dropna=True)
pt.to_excel(writer, sheet_name="AppPurpose-Target")


pt = pd.pivot_table(tdf,
                    index=['App Tier'],
                    columns=["Target"],
                    values=["System name"],
                    aggfunc={"System name": len},
                    margins=True,
                    margins_name="Total",
                    dropna=False)
pt.to_excel(writer, sheet_name="App-Target")

pt = pd.pivot_table(tdf,
                    index=['Server model'],
                    columns=["Target"],
                    values=["System name"],
                    aggfunc={"System name": len},
                    margins=True,
                    margins_name="Total",
                    dropna=False)
pt.to_excel(writer, sheet_name="Model-Target")



writer.save()
writer.close()

