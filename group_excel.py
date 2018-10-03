import pandas as pd

# https://www.dataquest.io/blog/pandas-cheat-sheet/
# http://pandas.pydata.org/pandas-docs/stable/generated/pandas.DataFrame.groupby.html
# https://blog.csdn.net/claroja/article/details/71080293?utm_source=itdadao&utm_medium=referral
# https://blog.csdn.net/leonis_v/article/details/51832916

if __name__ == '__main__':
    input_excel = 'Actelion.xlsx'
    input_sheet = 'Sheet1'
    column_want_to_group = 'Owner.Name'

    df = pd.read_excel(input_excel, sheet_name=input_sheet)

    for group_name, group in df.groupby(column_want_to_group):
        writer = pd.ExcelWriter(group_name + '.xlsx')
        group.to_excel(writer, 'Sheet1', index=False)
        writer.save()
