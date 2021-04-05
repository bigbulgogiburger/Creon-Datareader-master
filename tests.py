import pandas as pd
import numpy as np


data = pd.read_csv('E:/big12/python-project/note/categories/제약기업선정.csv', encoding='utf-8')
info_list = data['code'].tolist()
print(info_list,type(info_list))

