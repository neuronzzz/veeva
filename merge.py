import glob

import pandas as pd

directory = "./Q4拜访计划ALL/*.xlsx"
output_doc = 'hehe.xlsx'
sheet = 'sheet1'

df = pd.DataFrame()

for file_name in glob.glob(directory):
    df = df.append(pd.read_excel(file_name), ignore_index=True)

writer = pd.ExcelWriter(output_doc)
df.to_excel(writer, sheet)
writer.save()

gp = df.groupby('Account Owner')
print(gp.count()['ID 18'])
