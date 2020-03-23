import os
import pandas as pd
import numpy as np

print(f"Importing Data ...")

# Import IEDS data
ieds_df = pd.read_excel(os.path.join('data', 'New TLG Inventory 02-28-2020.xlsx'), sheet_name="Source-IEDS ref-Jeff")
ieds_df = ieds_df[ieds_df['Server Name'].notna()]
ieds_df = ieds_df[ieds_df['Server Name'].str.len().notna()]
ieds_df.columns = ieds_df.columns.str.replace('Server Name', 'Hostname')
ieds_df.sort_values("Hostname", inplace=True)

# Import prod master
prod_master_df = pd.read_excel(os.path.join('data', 'New TLG Inventory 02-28-2020.xlsx'), sheet_name="Prod Master")
prod_master_df = prod_master_df[prod_master_df['ServerName'].notna()]
prod_master_df = prod_master_df[prod_master_df['ServerName'].str.len().notna()]
prod_master_df.columns = prod_master_df.columns.str.replace('ServerName', 'Hostname')
prod_master_df.sort_values("Hostname", inplace=True)

# Import ESX Production Servers
esx_prd_srv_df = pd.read_excel(os.path.join('data', 'TLG-MOB ESX Servers v1.xlsx'), sheet_name="TLG Production Servers")
esx_prd_srv_df.sort_values("Hostname", inplace=True)

# Import ESX Production ESX
esx_prd_esx_df = pd.read_excel(os.path.join('data', 'TLG-MOB ESX Servers v1.xlsx'), sheet_name="TLG Production ESX")
esx_prd_esx_df.sort_values("Hostname", inplace=True)

# Combined ESX data
esx_df = pd.concat([esx_prd_esx_df, esx_prd_srv_df])
esx_df.sort_values("Hostname", inplace=True)
esx_df.drop_duplicates(subset="Hostname", keep=False, inplace=True)

# Combine Ieds and Esx
ieds_esx_df = pd.merge(left=ieds_df, right=esx_df, how='outer', left_on='Hostname', right_on='Hostname', suffixes=("", "_dup"))
ieds_esx_df = ieds_esx_df.sort_index(axis=1)
ieds_esx_df.sort_values("Hostname", inplace=True)
ieds_esx_df.columns = ieds_esx_df.columns.str.replace('#', 'Num')
dups = ieds_esx_df[ieds_esx_df.duplicated('Hostname')]
ieds_esx_df = ieds_esx_df[['Hostname'] + [col for col in ieds_esx_df.columns if col != 'Hostname']]

ieds_esx_df['Hostname'] = ieds_esx_df['Hostname'].str.lower()
# Label hardware abstraction
ieds_esx_df.loc[ieds_esx_df.eval('`Hardware Abstraction` == "" & `Server Model` == "Vmware virtual platform"'), 'Hardware Abstraction'] = 'Vmguest'
ieds_esx_df.loc[ieds_esx_df.eval('`Hardware Abstraction` == "" & ( `Server Model` != "Vmware virtual platform" | `Server Model` != "")'), 'Hardware Abstraction'] = 'Bare-metal'

# Import sudeep data
sudeep_df = pd.read_excel(os.path.join('data', 'Server_list 2020-03-19.xlsx'))
sudeep_df.sort_values("System name", inplace=True)
sudeep_df = sudeep_df[['System name'] + [col for col in sudeep_df.columns if col != 'System name']]

# Combine sudeep's data with ieds data, keep sudeep's data but fill in any missing values
sdf = sudeep_df.set_index('System name', drop=False)
iedf = ieds_esx_df.rename(columns={'Hostname':'System name'})
iedf = iedf.set_index('System name', drop=False)
iedf = iedf[iedf['System name'].isin(sdf['System name'])]
sudeep_ieds_esx_df = sdf.combine_first(iedf)
sudeep_ieds_esx_df = sudeep_ieds_esx_df.reset_index(drop=True)
sudeep_ieds_esx_df.sort_values("System name", inplace=True)
sudeep_ieds_esx_df = sudeep_ieds_esx_df[['System name'] + [col for col in sudeep_ieds_esx_df.columns if col != 'System name']]


# Add combined columns
sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df['Present State Cores'].isna(), 'Present State Cores'] = sudeep_ieds_esx_df["CPU"]
sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df['Present State Cores'].isna(), 'Present State Cores'] = sudeep_ieds_esx_df["Num Cores"]
sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df['Present State RAM GB'].isna(), 'Present State RAM GB'] = sudeep_ieds_esx_df["Memory (GB)"]
sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df['Present State RAM GB'].isna(), 'Present State RAM GB'] = sudeep_ieds_esx_df["RAM (GB)"]
sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df['Environment'].isna(), 'Environment'] = sudeep_ieds_esx_df["Environment_dup"]
sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df['Server Model'].isna(), 'Server Model'] = sudeep_ieds_esx_df["Hardware Model Description"]

# Modify data
sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df['CPU Speed'].isnull(), 'CPU Speed'] = sudeep_ieds_esx_df["CPU model"]
sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df['CPU Speed'] == "2.40GHz", 'CPU Speed'] = 2400
sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df['CPU Speed'] == "2.50GHz", 'CPU Speed'] = 2500
sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df['CPU Speed'] == "2.13GHz", 'CPU Speed'] = 2130
sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df['CPU Speed'] == "Intel(R) Xeon(R) CPU E7- 4870  @ 2.40GHz", 'CPU Speed'] = 2400
sudeep_ieds_esx_df = sudeep_ieds_esx_df.astype({"CPU Speed":int})

# sudeep_ieds_esx_df.to_excel("s.xlsx")




# Repalce nans with empty string for non float columns
for col in sudeep_ieds_esx_df.columns:
    if sudeep_ieds_esx_df[col].dtype.kind == 'O':
        sudeep_ieds_esx_df[col].replace(np.nan, '', regex=True, inplace=True)

# Capitalize first letter and lowercase the rest of each word (consistent data)
for col in sudeep_ieds_esx_df.columns:
    if sudeep_ieds_esx_df[col].dtype.kind == 'O':
        sudeep_ieds_esx_df[col] = np.where(sudeep_ieds_esx_df[col].str.lower().isnull(),
                                           sudeep_ieds_esx_df[col],
                                           sudeep_ieds_esx_df[col].str.lower().str.capitalize())


sudeep_ieds_esx_df['Consolidate'] ="Many"
sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df["Server Purpose"].str.contains("Customer"), "Consolidate"] = "One"
sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df["Server Purpose"].str.contains("drm"), "Consolidate"] = "One"


# Add CPU ratio compression column
# sudeep_ieds_esx_df['Core Consolidation Ratio'] = 1
# sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df['Server Model'].str.lower().str.contains("gen9", na=False), 'Core Consolidation Ratio'] = .9
# sudeep_ieds_esx_df.loc[sudeep_ieds_esx_df['Server Model'].str.lower().str.contains("g7", na=False), 'Core Consolidation Ratio'] = .8


# Specint Data
# spec = sudeep_ieds_esx_df[['Server Model', 'Specint','CPU Model', 'CPU Speed', 'Hardware Model Description']]
# spec = spec.groupby('Server Model')['Specint'].max().reset_index()
# spec.to_excel("output/specint.xlsx")

# perf = sudeep_ieds_esx_df[['Server Model', 'Hardware Model Description', 'CPU Model', 'CPU Speed','Specint']].drop_duplicates()
# perf = sudeep_ieds_esx_df[['Server Model', 'CPU Speed']].drop_duplicates()
# perf.to_excel("output/perf.xlsx")
# sudeep_ieds_esx_df['CPU Speed'].unique()

# sudeep_ieds_esx_df.to_excel("all.xlsx")


# pt = pd.pivot_table(sudeep_ieds_esx_df,
                    # index=["Server Model"],
                    # values=["System name"],
                    # aggfunc='count',
                    # margins=True,
                    # margins_name="Total",
                    # dropna=False)
# pt.to_excel("output/ModelCounts.xlsx")



