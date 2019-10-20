import pandas as pd
import clear_comment as cc
import re


df = pd.read_excel('京东商品评论(1).xlsx')
kw = pd.read_excel('京东商品评论(1).xlsx',sheet_name='Sheet2')
kw = kw[['场景','风格','怀孕时间']]

labeled = cc.label_keywords(kw,df,'评价内容')
all = [x for x in labeled.columns if re.search('场景|风格|怀孕时间',x)]
labeled_combine = cc.column_combine(labeled,all,',','标注')

labeled_combine[['评价内容','商品属性','页面标题','标注']].to_excel('1019新.xlsx')

new_df = pd.read_excel('1019新.xlsx')

scene = '|'.join([str(i) for i in kw['场景']])
style = '|'.join([str(i) for i in kw['风格']])
preg = '|'.join([str(i) for i in kw['怀孕时间']])

new_df['尺寸'] = new_df['商品属性'].apply(lambda x:re.findall('SS|S|M|L|LL',x,flags=re.IGNORECASE)).astype(str).apply(lambda x:re.sub('\[|\]|','',x))
new_df['场景'] = new_df['标注'].apply(lambda x:re.findall(scene,x)).astype(str).apply(lambda x:re.sub('\[|\]|\' ','',x))
new_df['风格'] = new_df['标注'].apply(lambda x:re.findall(style,x)).astype(str).apply(lambda x:re.sub('\[|\]|\' ','',x))
new_df['怀孕时间'] = new_df['标注'].apply(lambda x:re.findall(preg,x)).astype(str).apply(lambda x:re.sub('\[|\]|\' ','',x))
new_df.to_excel('1020新.xlsx')

df = pd.read_excel('1020新.xlsx')
df['时装'] = df['页面标题'].apply(lambda x:re.findall('春装|夏装|秋装|冬装',x)).astype(str).apply(lambda x:re.sub('\[|\]|\' ','',x))
df.to_excel('时装.xlsx')

