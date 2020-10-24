import numpy as np
import re
import openpyxl as px
import pandas as pd
import itertools
from openpyxl.utils.dataframe import dataframe_to_rows

class ExcelPaster:
    def __init__(self, book_name):
        #テンプレートから読み込みをする
        self.web=px.load_workbook(book_name) 

# ランダムウォーク生成
def random_walk(length=100, seed=1234):
    random.seed(a=seed)
    x = [0]
    for j in range(length - 1):
        step_x = random.randint(0, 1)
        if step_x == 1:
            x.append(x[j] + 1 + 0.05*np.random.normal())
        else:
            x.append(x[j] - 1 + 0.05*np.random.normal())
    return x

dtidx = pd.date_range(start='2017-10-1', end='2017-12-31', freq='1d')

length = len(dtidx)
X = pd.DataFrame([random_walk(length, seed) for seed in [1, 10, 100]], 
                 columns=dtidx, index=['A', 'B', 'C']).T