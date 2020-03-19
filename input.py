import os
import pandas as pd
import xlsxwriter
import re
import numpy as np
# import datacompy


# Create data import progress spreadsheet
writer = pd.ExcelWriter(os.path.join('output', 'targets.xlsx'), engine='xlsxwriter')

# Import data
df = pd.read_excel(os.path.join('data', 'Exported file.xlsx'))
# Remove server names that are null
print(f"Dataframe: {df.shape[0]} rows, {df.shape[1]} columns")


# Target systems
targets = [
    ['DL360/2s/48/768GB', 48, 768],
    ['DL360/2s/48/1.536TB', 48, 1536],
    ['DL580/4s/96/3TB', 96, 3072],
    ['', np.nan, np.nan],
]
tgdf = pd.DataFrame(targets, columns=['Target', 'Cores', 'Memory'])


### Rules
# Each rule is applied to every row
# The last matching rule takes precedence
# Data words are fist letter capital rest undercase
# One to one mapping
rules = [
    ['s["Present State RAM GB"] <= 768 & s["Present State Cores"] <= 48', "DL360/2s/48/768GB"],
    ['(s["Present State RAM GB"] > 768 & s["Present State RAM GB"] < 1536) & s["Present State Cores"] < 48', "DL360/2s/48/1.536TB"],
    ['s["Present State Cores"] > 48', "DL580/4s/96/3TB"],
    ['s["Present State RAM GB"] >= 3000', "DL580/4s/96/3TB"],
    ['s["Server Purpose"].str.contains("Customer")', "DL580/4s/96/3TB"],
]
rdf = pd.DataFrame(rules, columns=['Rule', 'Target'])

rule = 'df["Present State RAM GB"] <= 768 & df["Present State Cores"] <= 48'

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
        s['PctCores'] = s['Present State Cores'] / tgdf[tgdf['Target'] == s['Target'].iloc[0]]['Cores'].iloc[0]
        s['PctMemory'] = s['Present State RAM GB'] / tgdf[tgdf['Target'] == s['Target'].iloc[0]]['Memory'].iloc[0]
    else:
        s['PctCores'] = np.nan
        s['PctMemory'] = np.nan

    tdf = tdf.append(s)

tdf.to_excel(writer, sheet_name="Servers")
tgdf.to_excel(writer, sheet_name="Targets")
rdf.to_excel(writer, sheet_name="Rules")

writer.save()

