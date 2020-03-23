import pandas as pd
import os
import numpy as np
from rules import ieds_df, prod_master_df, esx_df, ieds_esx_df, sudeep_df, targets_df, rules_df, target_df, sudeep_ieds_esx_df
from math import ceil

writer = pd.ExcelWriter(os.path.join('output', 'targets.xlsx'), engine='xlsxwriter')
workbook = writer.book

target_df.to_excel(writer, sheet_name="Servers")
worksheet = writer.sheets["Servers"]

# pd.set_option('display.width', 400)
# pd.set_option('display.max_columns', 90)

hide_columns = [
       'Add Date', 'Additional Info', 'App Dependencies',
       'Belongs to ESX Cluster', 'CPU', 'CPU Model',
       'CPU model', 'City', 'Cluster Name', 'Comments',
       'Comments_dup', 'Consolidation group',
       'Consolidation group foreign key', 'Consolidation method',
       'Cores per physical CPU', 'Count of DBs_and_Apps', 'CreatedBy',
       'CreationTimestamp', 'Current State',
       'Current configuration from catalog',
       'Current configuration performance figure', 'DNS', 'DR needed',
       'ESX Cluster', 'ESX Cluster Name', 'ESX Farm',
       'Environment Details', 'Environment Type',
       'Environment_dup', 'Fraction of cores Stin', 'Future State',
       'Future State 1_1', 'Future State 1_1 DR', 'Future State CPU',
       'Future State CPU DR', 'Future State Cores',
       'Future State Cores DR', 'Future State RAM', 'Future State RAM DR',
       'Future State Storage', 'Future State Storage DR', 'HT supported',
       'Hardware Abstraction',
       'IEDS status', 'IP Address', 'Instance', 'Inventory check',
       'List of applications by DBs_and_Apps',
       'List of applications by DBs_and_Apps inline',
       'List of applications to show', 'List of markets by DBs_and_Apps',
       'List of markets by DBs_and_Apps inline',
       'List of markets to show', 'MOTS ID', 'MOTS Name', 'Market(s)',
       'Markets', 'Markets by line', 'Max CPU %', 'Max Disk IO',
       'Max Num of Cores Used', 'Max RAM % Used', 'Max RAM used (GB)',
       'Max Run Queue', 'Max Specints Used',
       'Migration method', 'ModificationTimestamp', 'ModifiedBy',
       'Non-TLG MOTS', 'Num CPU', 'Num Cores', 'Num Cores per Socket',
       'Num of Sockets', 'Number of markets by DB_and_apps', 'Numbering',
       'OS', 'OS Level', 'OS Major Version', 'OS Patch Level',
       'OS Version', 'OS_dup', 'Parent ESX', 'Present State CPUs',
       'Present State Cores_with_HT',
       'Present State Storage GB', 'PrimaryKey',
       'RAM (GB)', 'Replacement configuration from catalog',
       'Replacement configuration performance figure',
       'Replacement fraction', 'Server Description', 'Server serial number',
       'Specint', 'State', 'Storage (GB)', 'Server model',
       'Tablespace Size Allocated (GB)', 'Target Data Center',
       'Total NAS Allocated (GB)', 'Total NAS Used (GB)',
       'Total SAN Allocated (GB)', 'Total SAN Used (GB)', 'Unnamed: 40',
       'VCenter', 'VMware Guest key', 'VMware host key', 'VPMO', 'Vendor',
       'Working set', '~EOSL'
]


for i in range(1, len(target_df.columns)+1):
    col = target_df.columns[i-1]
    width = max([len(str(s)) for s in target_df[col].values] + [len(col)])
    if col in hide_columns:
        worksheet.set_column(i, i, width, None, {'hidden': True})
    else:
        worksheet.set_column(i, i, width)

targets_df.to_excel(writer, sheet_name="Targets",index=False)
rules_df.to_excel(writer, sheet_name="Rules")

pt = pd.pivot_table(target_df,
                    index=["Target"],
                    values=["System name"],
                    aggfunc={"System name": len},
                    margins=True,
                    margins_name="Total",
                    dropna=False)
pt.to_excel(writer, sheet_name="TargetCount")
# ieds_esx_df.to_excel(writer, sheet_name="IEDS ESX")

pt = pd.pivot_table(target_df,
                    index=['Data Center', 'App Tier',  "Consolidate", "Target"],
                    values=["System name", "Present State Cores"],
                    aggfunc={"System name": len, "Present State Cores": sum},
                    margins=True,
                    margins_name="Total",
                    dropna=True)

def cons(df):
    target_cores = targets_df[targets_df['Target'] == df['Target'].unique()[0]]['Cores'].iloc[0]
    sum_cores = df['Present State Cores'].sum()
    if df["Consolidate"].iloc[0] == "Many":
        return ceil(sum_cores / target_cores)
    else:
        return df['System name'].count()

cons_df = target_df.groupby(['Data Center', 'App Tier', 'Consolidate','Target']).apply(cons)
cons_df = cons_df.to_frame().reset_index()
cons_df.rename(columns={0:'Consolidated Count'}, inplace=True)
pt = pd.DataFrame(pt).reset_index()
pt.rename(columns={'System name':'Count'}, inplace=True)
pt_df = pt.merge(cons_df)
pt_df.to_excel(writer, sheet_name="Consolidation", index=False)

writer.save()
writer.close()





df = target_df
df['Target'].unique()[0]
cores = targets_df[targets_df['Target']==df['Target'].unique()[0]]['Cores'].iloc[0]

b = target_df.groupby(['Data Center', 'App Tier', 'Target']).agg({"Present State Cores": "sum"})

df['Present State Cores'].sum()






pt = pd.pivot_table(target_df,
                    index=["Server Purpose", "Present State Cores", "Present State RAM GB"],
                    columns=["Target"],
                    values=["System name"],
                    aggfunc={"System name": len},
                    margins=True,
                    margins_name="Total",
                    dropna=True)
pt.to_excel(writer, sheet_name="ServerPurposeCoresMem-Target")


pt = pd.pivot_table(target_df,
                    index=['App Tier', 'Server Purpose'],
                    columns=["Target"],
                    values=["System name"],
                    aggfunc={"System name": len},
                    margins=True,
                    margins_name="Total",
                    dropna=True)
pt.to_excel(writer, sheet_name="AppPurpose-Target")


pt = pd.pivot_table(target_df,
                    index=['App Tier'],
                    columns=["Target"],
                    values=["System name"],
                    aggfunc={"System name": len},
                    margins=True,
                    margins_name="Total",
                    dropna=False)
pt.to_excel(writer, sheet_name="App-Target")

pt = pd.pivot_table(target_df,
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

