import glob
from collections import defaultdict

import numpy as np
import pandas as pd
from pylev import levenshtein
from sklearn.cluster import DBSCAN


# 需求：合并文档
# 问题：
# 1. 合并的文档，可能存在客户自己修改了column name，但是通常情况下，只会修改错误某个字符，比如:'2018 Q4'，修改成了'2018 Q3'，或者
# '1018 Q4'，这种问题，用编辑距离<1来衡量column是否是一样的
# 2. 对于因为客户修改，导致出来多个column，本质上是同一个column，需要做合并处理到一个column，并且，里面的值要做sum，并放在一个新column
# 3. 对于数值的cell，可能存在nan，也就是没有任何值
# 4. 对于本来是数值的cell，可能存在是字符型的cell
# 5. 特殊问题：文件中有空的隐藏行需要去掉

def lev_metric(x, y):
    i, j = int(x[0]), int(y[0])  # extract indices
    return levenshtein(column_names[i], column_names[j])


def dataframe_diff(df1, df2):
    return df1.append(df2).drop_duplicates(keep=False)


if __name__ == '__main__':
    directory = "./Q4拜访计划ALL/*.xlsx"
    output_doc = 'hehe1.xlsx'
    sheet = 'sheet1'
    col_ID_18 = 'ID 18'

    df = pd.DataFrame()

    union_columns = []

    count1 = 0
    count2 = 0
    for file_name in glob.glob(directory):
        # df = df.append(pd.read_excel(file_name, sheetname=sheet), ignore_index=True)
        # 对于pd.read_excel, 如果没有设置sheet_name，会读取第一个sheet，或者读取所有？如果对excel的处理就是处理第一个，那么不设置值
        # 比较好，如果设置值，值如果变化了，程序，或者文件必须要进行修改保持一致
        df_part = pd.read_excel(file_name)
        df1 = df_part
        # 下面这个过滤的过程，可以看出DataFrame，仍然是按照行索引来过滤行，里面是用来过滤掉响应的索引，得到索引之后来获取行
        # 过滤过程相当于[]中括号内部，传入一个条件，根据这个条件来过滤。本质上，这个条件返回的是一个布尔索引（series with True/False value），
        # 根据这个索引来决定哪些row保留
        # 过滤null的几种方法：https://stackoverflow.com/questions/22551403/python-pandas-filtering-out-nan-from-a-data-selection-of-a-column-of-strings
        # df_part = df_part.dropna(thresh=9)
        # 由于有的ID 18是空，所以需要用Name来过滤
        df_part = df_part[df_part['Name'].notnull()]
        df2 = df_part

        # todo: debug code
        count1 += df1.shape[0]
        count2 += df2.shape[0]

        if df1.shape[0] != df2.shape[0]:
            df_diff = dataframe_diff(df1, df2)
            print('--------------------diff------------------')
            print(file_name + '-------' + str(df1.shape[0]))
            print(file_name + '-------' + str(df2.shape[0]))
            print(df_diff)
            print('--------------------diff end------------------')
        # todo: debug code end

        # 下面是对column进行合并，并去重处理，利用了set，但是set会把list顺序打乱，所以要保存list的index，合并之后利用这个index重新排序
        # 下面df_part与df的顺序不能调换，因为df的构建有一个append操作，如果append的dataframe中有df没有个column，df在append之后
        # column顺序就会改变
        duplicated_columns = np.hstack((df_part.columns.values, df.columns.values))
        union_columns = list(set(duplicated_columns))
        union_columns.sort(key=duplicated_columns.tolist().index)
        df = df.append(df_part, ignore_index=True)

    # 下面这个主要用来保持dataframe的column与原始的一致，因为如何两个excel column不一致，append的时候，会重新按照A-Za-z把column排序
    df = df[union_columns]

    gp = df.groupby('Account Owner')
    column_names = df.columns.values
    print(gp.count()[col_ID_18])

    # 可以用exit()函数来退出脚本的执行
    # exit()

    # cluster column names base on the edit distance
    X = np.arange(len(column_names)).reshape(-1, 1)
    db = DBSCAN(eps=1, min_samples=1, metric=lev_metric).fit(X)
    # labels_ is the cluster, 结果是X对应的cluster，例如X行数是10，结果可能是[0，0，0，1，2，3，4，5，6，7]，表示前三个是cluster 0
    print(db.labels_)

    # solution1: build dictionary, cluster is the key and merge columns to list as the value
    # 下面这个方案是把DBSCAN计算得到的cluster转成dictionary，key是cluster，value是X的索引
    # https://www.cnblogs.com/herbert/archive/2013/01/09/2852843.html
    # https://www.tutorialspoint.com/How-to-create-Python-dictionary-with-duplicate-keys
    d = defaultdict(list)
    for k, v in zip(db.labels_, list(range(len(db.labels_)))):
        d[k].append(v)

    d = dict(d)

    # 合并在cluster中的column，并且做sum操作来合并
    # ''.join(<list>) 要求list是string类型的list，起码不能是int类型的list，否则回报如下异常
    # TypeError: sequence item 0: expected str instance, int found
    for k, v in d.items():
        if len(v) > 1:
            # dataframe.icol(i), i : int, slice, or sequence of integers
            # 对于新的pandas，版本超过0.19.2的，需要用icol来根据列索引获取列，而不能直接用df[v]这样的方式
            # df_group = df.icol(v)
            # 因为如下提示，讲icol替换成iloc
            # /Users/ryan/github/veeva/merge_excels.py:107: FutureWarning: icol(i) is deprecated. Please use .iloc[:,i]
            #   df_group = df.icol(v)
            df_group = df.iloc[:, v]
            df[' & '.join(df_group.columns)] = df_group.fillna(0).astype(float).sum(axis=1)
            df = df.drop(df.columns[v], axis=1)
            # try:
            #     df_group = df.icol()[v]
            #     df[' & '.join(df_group.columns)] = df_group.fillna(0).astype(float).sum(axis=1)
            #     df = df.drop(df.columns[v], axis=1)
            # except Exception as e:
            #     raise e

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
    print(count1)
    print(count2)
    print(df.shape[0])
