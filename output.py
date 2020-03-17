import os
import pandas as pd
import xlsxwriter
import re
import numpy as np
from input import tdf

# Do not modify imported df
df = tdf.copy()
df = df.replace(np.nan, '', regex=True)

# Create xlsxwriter
writer = pd.ExcelWriter(os.path.join('output', 'data_analysis.xlsx'), engine='xlsxwriter')

pt = pd.pivot_table(df, index=['Data Center', 'Target'],
                    columns="Environment",
                    values=["Hostname"],
                    aggfunc="count",
                    margins=True,
                    margins_name="Total",
                    dropna=False)
pt.to_excel(writer, sheet_name="DataCenter Target-Env")

pt = pd.pivot_table(df, index=['Data Center'],
                    columns="Target",
                    values=["Hostname"],
                    aggfunc="count",
                    margins=True,
                    margins_name="Total",
                    dropna=False)
pt.to_excel(writer, sheet_name="DataCenter-Target")

pt = pd.pivot_table(df, index=['Purpose'],
                    columns="Hardware Abstraction",
                    values=["Hostname"],
                    aggfunc="count",
                    margins=True,
                    margins_name="Total",
                    dropna=False)
pt.to_excel(writer, sheet_name="Purp-HWAbstract")


pt = pd.pivot_table(df, index=['Server Model'],
                    columns="Hardware Abstraction",
                    values=["Hostname"],
                    aggfunc="count",
                    margins=True,
                    margins_name="Total",
                    dropna=False)
pt.to_excel(writer, sheet_name="ServerModel-HWAbstract")


pt = pd.pivot_table(df[df.eval('`Hardware Abstraction` == ""')], index=['Hardware Model Description', 'Hostname'],
                    columns=["Server Model"],
                    values="Target",
                    aggfunc="count",
                    margins=True,
                    margins_name="Total",
                    dropna=False)
pt.to_excel(writer, sheet_name="Undetermined Hosts")




writer.save()







