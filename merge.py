import glob

import numpy as np
import pandas as pd
from pylev import levenshtein
from sklearn.cluster import DBSCAN
from collections import defaultdict


def lev_metric(x, y):
    i, j = int(x[0]), int(y[0])  # extract indices
    return levenshtein(cols[i], cols[j])


if __name__ == '__main__':
    directory = "./Q4拜访计划ALL/*.xlsx"
    output_doc = 'hehe.xlsx'
    sheet = 'sheet1'

    df = pd.DataFrame()

    for file_name in glob.glob(directory):
        df = df.append(pd.read_excel(file_name), ignore_index=True)

    gp = df.groupby('Account Owner')
    print(gp.count()['ID 18'])
    cols = df.columns.values

    X = np.arange(len(cols)).reshape(-1, 1)
    db = DBSCAN(eps=1, min_samples=1, metric=lev_metric).fit(X)
    print(db.labels_)

    # solution1: build dictionary and merge
    # https://www.cnblogs.com/herbert/archive/2013/01/09/2852843.html
    # https://www.tutorialspoint.com/How-to-create-Python-dictionary-with-duplicate-keys
    d = defaultdict(list)
    for k, v in zip(db.labels_, list(range(len(db.labels_)))):
        d[k].append(v)

    d = dict(d)

    # ''.join(<list>) 要求list是string类型的list，起码不能是int类型的list，否则回报如下异常
    # TypeError: sequence item 0: expected str instance, int found
    for k, v in d.items():
        if len(v) > 1:
            df_group = df[v]
            df[' & '.join(df_group.columns)] = df_group.fillna(0).sum(axis=1)

    # solution2: set db.labels_ as new column -> group/cluster
    # solution2: doest works as dataframe stack will be series, for each item in series will be a series for each record
    # solution2: for details, please check below values: df.stack

    writer = pd.ExcelWriter(output_doc)
    df.to_excel(writer, sheet, index=False)
    writer.save()
