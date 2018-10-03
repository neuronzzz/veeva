import pandas as pd

if __name__ == '__main__':
    doc1 = 'doc1.xlsx'
    doc2 = 'doc2.xlsx'
    doc_output = 'output_excel.xlsx'
    sheet1 = 'sheet_in_doc1'
    sheet2 = 'sheet_in_doc2'
    sheet_output = 'output_excel_sheet_name'
    col_a1 = 'col_a1'
    col_b1 = 'col_b1'
    col_a2 = 'col_a2'
    col_b2 = 'col_b2'
    df1 = pd.read_excel(doc1, sheetname=sheet1, header=0)
    df2 = pd.read_excel(doc2, sheetname=sheet2, header=0)

    df1[col_b1] = pd.merge(df1, df2, left_on=col_a1, right_on=col_a2)[col_b2]
    writer = pd.ExcelWriter(doc_output)
    df1.to_excel(writer, sheet_output, index=False)
    writer.save()
