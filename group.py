import pandas as pd

# https://www.dataquest.io/blog/pandas-cheat-sheet/
# http://pandas.pydata.org/pandas-docs/stable/generated/pandas.DataFrame.groupby.html
# https://blog.csdn.net/claroja/article/details/71080293?utm_source=itdadao&utm_medium=referral
# https://blog.csdn.net/leonis_v/article/details/51832916

excel = 'Actelion.xlsx'
sheet = 'Sheet1'
column = 'Owner.Name'

df = pd.read_excel(excel, sheet_name=sheet)

for name_column, group_dataframe in df.groupby(column):
    writer = pd.ExcelWriter(name_column + '.xlsx')
    group_dataframe.to_excel(writer, 'Sheet1', index=False)
    writer.save()
