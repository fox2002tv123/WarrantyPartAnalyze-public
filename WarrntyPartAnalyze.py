'''
Author: bob
Date: 2022-04-28 22:20:25
LastEditors: bob
LastEditTime: 2022-04-28 22:20:26
FilePath: \WarrantyPartAnalyze-master\WarrntyPartAnalyze.py
Description: 此文件用于分析配件数据

Copyright (c) 2022 by bob, All Rights Reserved. 
'''
# %%
# 获取data文件夹下的所有文件xlsx文件
import glob
# 获取当前完整路径
import os
import datetime
import pandas as pd
import numpy as np
path=os.getcwd() # 获取当前工作目录
# 创建函数，获取当前工作目录，文件名，文件名后缀
# step -3 获取当前工作目录
def get_current_path():
    
    path=os.getcwd() # 获取当前工作目录
    path
    file_names=glob.glob(f'{path}\\data\\*.xlsx')
    # file_names=['data\\'+i for i in file_names]
    return file_names


# %%
# 加载数据
# import pandas as pd
# import numpy as np
# concat函数用于拼接数据
# step-2 将数据拼接成一个dataframe
# 创建函数，用于拼接数据
def concat_data(file_names):
    src_data=[]
    for i in file_names:
        data=pd.read_excel(i)
        data['文件名']=i.replace(f'{path}\\data\\','') # 添加文件名列
        src_data.append(data) # 将每个文件的数据拼接到一起
        
    df_src=pd.concat(src_data) # 将拼接好的数据放到一个dataframe中

    df_src.head()
    return df_src

# %%
# df_src['Parts Number'].unique()

# %%
# 创建函数，用于当df_src['Parts Number'].unique()>1时，通过拆分数据，将数据拆分成多个数据
# step -1 将数据拆分成多个数据
def split_data(df_src):
    df_src_split=[]
    for i in df_src['Parts Number'].unique():
        df_src_split.append(df_src[df_src['Parts Number']==i])
    return df_src_split

# %%
# 运行step 0
# 处理数据
# df = df_src.copy()
# 提取文件名中的时间
# 创建函数，用于提取文件名中的时间


def extract_time(df):

    # import datetime # 已经在上一步中导入
    # 提取时间
    df['时间'] = df['文件名'].apply(lambda x: datetime.datetime.strptime(
        x.split(' ')[0], 'TABLE6_%Y-%m-%d'))
    # 按经销商-时间排序-升序
    df = df.sort_values(by=['Dealer No.', '时间'])
    df.index = df['时间']
    df[['Dealer No.', 'QTY', 'Reorder QTY', '时间']].head()
    return df


# %%
# df.info()

# %%


# 创建函数，lambda x:x['QTY']+x['Reorder QTY']-x['QTY'].shift(1)-x['Reorder QTY'].shift(1)
# 运行step2
def get_diff(x):
    # 
    c1=x['QTY'].shift(1)+x['Reorder QTY'].shift(1)-x['QTY']-x['Reorder QTY']
    c1=c1.map(lambda x:x if x>0 else 0) # 如果x<0，则返回0，原因是0的时候，不能减去
    # print(c1)
    
    return c1 
    
# 创建函数，单一零件号的情况

# 运行step1
def get_single_part(df):
    
    # df下一行(QTY+Reorder QTY)减去df上一行(QTY+Reorder QTY)
    res_all=df.groupby(['Dealer No.','Dealer Name','Parts Number']).apply(get_diff)
    # 添加汇总列
    res_all['汇总']=res_all.sum(axis=1) #! 单一数据可使用

    # 转换列名称为日期2022-4-28 00:00:00 为2022-4-28
    res_all.columns=res_all.columns.map(lambda x:str(x)[:10]) #! 单一数据可使用

    res_all
    # 生成报表
    # todo 修正了输出 PartNum_16137404081-2022-04-28.xlsx
    # import datetime # 已经在上一步中导入
    part_num=df['Parts Number'].unique()[0].replace(' ','')
    res_all.to_excel(f'{path}/output/PartNum_{part_num}-{datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")}.xlsx')



# %%
if __name__=='__main__':
    # 
    file_names=get_current_path() # 获取当前工作目录
    df_src=concat_data(file_names) # 将数据拼接成一个dataframe
    df_src_split=split_data(df_src) # 将数据按零件号拆分成多个数据
    # 循环调用get_single_part函数
    for i in df_src_split:
        df=extract_time(i) # 提取文件名中时间，变成时间列
        get_single_part(df) # 各个打印报表
    
    



