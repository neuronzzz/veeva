import pandas as pd


def compare(s1, s2):
    print('s1: %s' % s1)
    print('s2: %s' % s2)
    try:
        return set(s1.split(';')) == set(s2.split(';'))
    except Exception:
        return False


if __name__ == '__main__':
    doc = 'HCP Compare List.xlsx'
    doc_x = 'Hello_Flora1.xlsx'
    sheet = 'HCP Compare'
    col1 = 'ATL Territory'
    col2 = 'Expected Territory'
    df = pd.read_excel(doc, sheetname=sheet, header=0)

    df['Equal'] = df.apply(lambda x: compare(x[col1], x[col2]), axis=1)
    writer = pd.ExcelWriter(doc_x)
    df.to_excel(writer, sheet)
    writer.save()
