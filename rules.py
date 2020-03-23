import pandas as pd
import numpy as np
from input import ieds_df, prod_master_df, esx_df, ieds_esx_df, sudeep_df, sudeep_ieds_esx_df

print("Applying target rules ...")

# Target systems
targets = [
    ['DL360-10-low', 48, 768],
    ['DL360-10-high', 48, 1536],
    ['DL580-10-low', 96, 1536],
    ['DL580-10-high', 96, 3072],
    ['Oversize', 112, 3072],
    ['', np.nan, np.nan],
]
targets_df = pd.DataFrame(targets, columns=['Target', 'Cores', 'Memory'])

### Rules
# Each rule is applied to every row
# The last matching rule takes precedence
# Data words are fist letter capital rest undercase
# One to one mapping
# TODO: > 96 cores we need a label
rules = [
    ['s["Present State RAM GB"] <= 768 & s["Present State Cores"] <= 48', "DL360-10-low"],
    ['(s["Present State RAM GB"] > 768 & s["Present State RAM GB"] < 1536) & s["Present State Cores"] <= 48', "DL360-10-high"],
    ['s["Present State Cores"] > 48 & s["Present State RAM GB"] <= 1536', "DL580-10-low"],
    ['s["Present State RAM GB"] > 1536', "DL580-10-high"],
    ['s["Consolidate"] == "One"', "DL580-10-high"],
    ['s["Present State Cores"] > 96', "Oversize"],
]
rules_df = pd.DataFrame(rules, columns=['Rule', 'Target'])

# rule = 'df["Present State RAM GB"] <= 768 & df["Present State Cores"] <= 48'
# res = pd.eval(rule, engine='python')
# res = pd.eval(rule, engine='python')

target_df = pd.DataFrame()
# Iterate over rows
for i in range(0, len(sudeep_ieds_esx_df)):
    # Select row as data frame preserving field types
    s = pd.DataFrame(sudeep_ieds_esx_df.iloc[i:i+1, :]).copy()

    # Iterate over rules
    s['Rules'] = ''
    s['Target'] = ''
    for idx, row in rules_df.iterrows():
        rule = row['Rule']
        res = pd.eval(rule, engine='python')
        if res.iloc[0]:
            s['Target'] = rules_df.iloc[idx, 1]
            s['Rules'] = s['Rules'] + str(idx) + ', '

    s = s.replace(regex=r', $', value="")

    # Calculations
    # #Todo: Proper title of columns
    # if s['Target'].iloc[0] != "":
        # s['PctCores'] = s['Present State Cores'] / targets_df[targets_df['Target'] == s['Target'].iloc[0]]['Cores'].iloc[0]
        # s['PctMemory'] = s['Present State RAM GB'] / targets_df[targets_df['Target'] == s['Target'].iloc[0]]['Memory'].iloc[0]
    # else:
        # s['PctCores'] = np.nan
        # s['PctMemory'] = np.nan

    target_df = target_df.append(s)

