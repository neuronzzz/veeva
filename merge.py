import glob

import numpy as np
import pandas as pd
from pylev import levenshtein
from sklearn.cluster import DBSCAN

directory = "./Q4拜访计划ALL/*.xlsx"
output_doc = 'hehe.xlsx'
sheet = 'sheet1'

df = pd.DataFrame()

for file_name in glob.glob(directory):
    df = df.append(pd.read_excel(file_name), ignore_index=True)

writer = pd.ExcelWriter(output_doc)
df.to_excel(writer, sheet, index=False)
writer.save()

gp = df.groupby('Account Owner')
print(gp.count()['ID 18'])

cols = df.columns.values


def lev_metric(x, y):
    return levenshtein(cols[x[0]], cols[y[0]])


X = np.arange(len(cols)).reshape(-1, 1)
db = DBSCAN(eps=1, min_samples=1, metric=lev_metric).fit(X)
print(db.labels_)

# todo: merge columns by sum
