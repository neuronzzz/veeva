import glob
from collections import defaultdict

import numpy as np
import pandas as pd
from pylev import levenshtein
from sklearn.cluster import DBSCAN


# 需求：合并文档，并合并重复的column。重复column根据编辑距离来计算，认为编辑距离小于1的就是重复的。合并的column要sum，并放在一个新column。
# 特殊问题：文件中有空的隐藏行需要去掉。

def lev_metric(x, y):
    i, j = int(x[0]), int(y[0])  # extract indices
    return levenshtein(column_names[i], column_names[j])


if __name__ == '__main__':
    directory = "./Q4拜访计划ALL/*.xlsx"
    output_doc = 'hehe.xlsx'
    sheet = 'sheet1'
    col_ID_18 = 'ID 18'

    df = pd.DataFrame()

    for file_name in glob.glob(directory):
        # df = df.append(pd.read_excel(file_name, sheetname=sheet), ignore_index=True)
        # for pd.ead_excel, if do not specify a sheet_name default will read the first sheet. this will be useful
        # especially if not clear about the name of first sheet but know only one sheet of first sheet is the
        # sheet want to use
        df_part = pd.read_excel(file_name)
        # 下面这个过滤的过程，可以看出DataFrame，仍然是按照行索引来过滤行，里面是用来过滤掉响应的索引，得到索引之后来获取行
        # 过滤null的几种方法：https://stackoverflow.com/questions/22551403/python-pandas-filtering-out-nan-from-a-data-selection-of-a-column-of-strings
        df_part = df_part[df_part[col_ID_18].notnull()]
        df = df.append(df_part, ignore_index=True)

    gp = df.groupby('Account Owner')
    print(gp.count()[col_ID_18])

    # writer = pd.ExcelWriter(output_doc)
    # df.to_excel(writer, sheet, index=False)
    # writer.save()

    # exit()

    # cluster column names base on the edit distance
    column_names = df.columns.values
    X = np.arange(len(column_names)).reshape(-1, 1)
    db = DBSCAN(eps=1, min_samples=1, metric=lev_metric).fit(X)
    # labels_ is the cluster
    print(db.labels_)

    # solution1: build dictionary, cluster is the key and merge columns to list as the value
    # https://www.cnblogs.com/herbert/archive/2013/01/09/2852843.html
    # https://www.tutorialspoint.com/How-to-create-Python-dictionary-with-duplicate-keys
    d = defaultdict(list)
    for k, v in zip(db.labels_, list(range(len(db.labels_)))):
        d[k].append(v)

    d = dict(d)

    # sum columns
    # ''.join(<list>) 要求list是string类型的list，起码不能是int类型的list，否则回报如下异常
    # TypeError: sequence item 0: expected str instance, int found
    for k, v in d.items():
        if len(v) > 1:
            df_group = df[v]
            df[' & '.join(df_group.columns)] = df_group.fillna(0).astype(float).sum(axis=1)
            df = df.drop(df.columns[v], axis=1)

    # try:
    #     df[' & '.join(df_group.columns)] = df_group.fillna(0).astype(float).sum(axis=1)
    # except Exception as e:
    #     print(e)

    # below solutions are not correct, do not stack then sum or apply other operations. Should directly sum on axis = 1
    # solution2: set db.labels_ as new column -> group/cluster
    # solution2: doest works as dataframe stack will be series, for each item in series will be a series for each record
    # solution2: for details, please check below values: df.stack

    writer = pd.ExcelWriter(output_doc)
    df.to_excel(writer, sheet, index=False)
    writer.save()
