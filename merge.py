import glob

import pandas as pd

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

# DBSCAN
from sklearn.cluster import DBSCAN
import Levenshtein
import numpy as np
from functools import partial
from sklearn.metrics.pairwise import pairwise_distances
from pylev import levenshtein

cols = df.columns.values
print(cols)


def lev_metric(x, y):
    i, j = int(x[0]), int(y[0])  # extract indices
    return levenshtein(cols[i], cols[j])


X = np.arange(len(cols)).reshape(-1, 1)
db = DBSCAN(eps=1, min_samples=1, metric=lev_metric).fit(X)
print(db.labels_)

# def mylev(s, u, v):
#     return levenshtein(s[int(u)], s[int(v)])


# print(pairwise_distances(np.arange(len(cols)).reshape(-1, 1), metric=partial(mylev, cols)))
