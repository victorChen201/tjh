# -*- coding: utf-8 -*-
from docxtpl import DocxTemplate, RichText, InlineImage
import pandas as pd
from matplotlib import pyplot as plt
import math
from datetime import datetime
import matplotlib.dates as mdates
from docx.shared import Mm, Inches, Pt
import numpy as np
# plt.rcParams['font.sans-serif'] = ['SimHei'] # 指定默认字体
# plt.rcParams['font.family'] = ['SimHei AR PL UKai CN ']
plt.rcParams['axes.unicode_minus'] = False  # 解决保存图像是负号'-'显示为方块的问题
# plt.rcParams["font.family"]="STSong"  #解决浓度单位显示乱码情况
plt.rcParams['font.sans-serif'] = ['SimHei']
df = pd.read_excel(r"淮南市秋季源解析数据.xlsx", header=0)
del_columns = ['城市','膜编号','TC', 'Li', 'Be', 'Sc', 'Ti', 'V', 'Cr', 'Mn', 'Co', 'Ni', 'Cu', 'Zn', 'Ga', 'Ge', 'As', 'Se', 'Rb', 'Sr', 'Ag', '111Cd', '114Cd', 'Cs', 'Ba', 'Tl', '206Pb', '207Pb', '208Pb', 'Bi']

'''
:名词解析
:水溶性离子: NO3、SO4、NH4,Cl,Na,K,Mg,Ca
:PM2_5：WSIN/[(OC*1.5+EC+WSIN)*1.13],WSIN=[SO4,NO3,C1,NH4,Na,K,Mg,Ca]
:硝酸盐: N03:
:硫酸盐: SO4
:铵盐: NH4
:有机碳: OC
:无机碳：EC
'''
PM2_5_FAZHI = 60.0
city = df['城市'][0]
year = df['采样时间'][0].year
month = df['采样时间'][0].month
season = '冬季'
if month >= 3 and month <=5:
    season = '春季'
elif month >= 6 and month <=8:
    season = '夏季'
elif month >= 9 and month <= 11:
    season = '秋季'
else:
    season ='冬季'

df.drop(del_columns, axis=1, inplace=True)
df.dropna(inplace=True, axis=0)
def pm2_5(df):
    '''
    :功能：计算PM2.5
    :PM2_5：WSIN/[(OC*1.5+EC+WSIN)*1.13],WSIN=[SO4,NO3,C1,NH4,Na,K,Mg,Ca]
    :param df: df中每行值
    :return: 返回每行的PM2.5
    '''
    wsin_v = df[wsin].sum() #df['NO3'] + df['SO4'] + df['NH4'] + df['Cl'] + df['Na'] + df['K'] + df['Mg'] + df['Ca']
    y = (df['OC']*1.5 + df['EC'] + wsin_v)*1.13
    pm2_5 = 100*wsin_v/y
    return pm2_5.round(2)

# df.dropna(axis=1, how='any', thresh=1, subset=None, inplace=True)
columns = ['OC', 'EC', 'SO4', 'NO3', 'Cl', 'NH4', 'Na', 'K', 'Mg', 'Ca']
columns_sum = df[columns].apply(lambda x: x.sum().round(2))
# print(columns_sum)
wsin = ['SO4','NO3','Cl','NH4','Na','K','Mg','Ca']
wsin_value = columns_sum[wsin].sum(axis=0).round(2)
PM2_5 =100*wsin_value/((columns_sum['OC']*1.5 + columns_sum['EC'] + wsin_value)*1.13)
decreaseFP1 = []
#化学成分平均值
chemica_mean = df[columns].apply(lambda x: x.mean().round(2))
mass_colunms = ['OC','SO4','NO3']
chemical_constituents = chemica_mean[mass_colunms].sort_values(ascending=False).index.values
# print(chemica_mean)

columns_pm2_5 = df[columns].apply(lambda x: pm2_5(x), axis=1, result_type='expand')
columns_pm2_5.columns=['pm2_5']
# print(columns_pm2_5)
df = pd.concat([df,columns_pm2_5],axis=1,join='inner')
# print(df)
df1 = df[df['pm2_5']>PM2_5_FAZHI]
df2 = df[df['pm2_5']<PM2_5_FAZHI]
s1 = df1[columns].apply(lambda x: x.mean().round(2))
s2 = df2[columns].apply(lambda x: x.mean().round(2))
s = (s1-s2)/s1
increaseFP = s.sort_values(ascending=False)
# print(increaseFP)
columns.append('站点')
# print(df[columns])

sites_detail = ""
sites_mean = df[columns].groupby(['站点',],axis=0).mean().round(2)
# print(sites_mean)
site_num = sites_mean.shape[0]
for key in sites_mean.index.values:
    site_module = "%s站%s占比最高的三大组分分别为%s、%s和%s、分别占比%3.1f、%3.1f和%3.1f；"
    s = sites_mean.loc[key].sort_values(ascending=False)
    sum = s.sum()
    s = 100*s/sum
    s = site_module%(key, season, s.index[0],  s.index[1], s.index[2], s.values[0],s.values[1],s.values[2])
    sites_detail = sites_detail + s
sites_detail = sites_detail[0:-1]
# print(sites_detail)

table_colunms = ['OC','EC','SO4','NO3','Cl','NH4']
legend = ['有机碳','无机碳','硫酸盐','硝酸盐','氯盐','铵盐']
table_others = ['Na','K','Mg','Ca']
df_sites = dict(list(df.groupby(['站点',],axis=0)))

# fig,axes=plt.subplots(3, 1, figsize=(12,12))
plt.figure(figsize=(12,6))
ax = plt.gca()
# ax.axis('off') #去掉边框
# print(img.index.values)
table_vals = []
plt.cla()

plt.xlabel('')
plt.subplots_adjust(bottom=0.14)
tb = table_colunms.copy()
tb.append('采样时间')
img_name = []
for key in df_sites.keys():
    df_date = df_sites[key][tb].copy(deep=True)
    # print(df_date.index)
    others = df_sites[key][table_others].apply(lambda x: x.sum(), axis=1, result_type='expand')
    # others.columns = ['钠钾钙镁盐',]
    # print(others)
    df_date = pd.concat([df_date, others], axis=1, join='inner')
    df_date.rename(columns={'0':'钠钾钙镁盐',}, inplace=True)
    df_date.columns.values[-1]='钠钾钙镁盐'
    print(df_date.columns.values)
    df_date['采样时间'] = df_date['采样时间'].apply(lambda x: x.__format__('%Y-%m-%d'))
    df_date.set_index('采样时间', inplace=True)
    plt.title(key,loc='left')
    plt.ylabel('质量浓度(µg/m\u00B3)', fontdict={'family': 'STSong', 'size': 12})
    df_date.plot(kind='bar',ax=ax,stacked=True)
    # plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    # plt.gca().xaxis.set_major_locator(mdates.DayLocator())
    # plt.gcf().autofmt_xdate() # 自动旋转日期标记
    name = key+'化学组成时间序列变化.png'
    img_name.append(name)
    plt.savefig(name)
    # plt.show()
    plt.cla()

tb = table_colunms
plt.axis('equal')
img1_name = []
plt.cla()
for key in sites_mean.index.values:
    s = sites_mean.loc[key]
    sum = s.sum()
    s = 100 * s / sum
    df_date = s[tb].copy(deep=True)
    others = s[table_others].sum()
    df_date['钠钾钙镁盐'] = others
    print(df_date)
    plt.title(key, loc='center')
    plt.pie(df_date.values,explode=[0]*df_date.size,labels=df_date.index,autopct='%.2f%%',pctdistance=0.6,labeldistance=1.2,shadow=True,startangle=0,radius=1.5,frame=False)
    # df_date.plot(kind='pie',ax=ax)
    # plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    # plt.gca().xaxis.set_major_locator(mdates.DayLocator())
    # plt.gcf().autofmt_xdate() # 自动旋转日期标记
    plt.legend()
    name = key+'化学组成构成.png'
    img1_name.append(name)
    plt.savefig(name)
    # plt.show()
    plt.cla()


tpl = DocxTemplate('淮南市大气细颗粒物来源解析工作_模板.docx')
context = {
    ####标题及第一部分####
    'city': city,
    'season': season,
    'PM2_5': PM2_5_FAZHI,
    'c': "OC、EC",
    's': "SO4",
    'chemical_constituents1': chemical_constituents[0],
    'chemical_constituents2': chemical_constituents[1],
    'chemical_constituents3': chemical_constituents[2],
     #increase_fastest_pollutant
    'increaseFP1': increaseFP.index[0],
    'increaseFP2': increaseFP.index[1],
    'increaseFP3': increaseFP.index[2],
    #increase_fastest_pollutant percent
    'increaseFPP1': round(increaseFP.values[0],2),
    'increaseFPP2': round(increaseFP.values[1],2),
    'increaseFPP3': round(increaseFP.values[2],2),
    #减少最多的成分
    'decreaseFP1': increaseFP.index[-1],
    'decreaseFP2': increaseFP.index[-2],
    'decreaseFPP1': round(math.fabs(increaseFP.values[-1]),2),
    'decreaseFPP2': round(math.fabs(increaseFP.values[-2]),2),
    #####第二部分####
    'site_num': site_num,
    'sites_detail': sites_detail,
    #####第三部分####
    'image1': InlineImage(tpl, img_name[0], width=Mm(150)),
    'image2': InlineImage(tpl, img_name[1], width=Mm(150)),
    'image3': InlineImage(tpl, img_name[2], width=Mm(150)),
    'image4': InlineImage(tpl, img1_name[0], width=Mm(150)),
    'image5': InlineImage(tpl, img1_name[1], width=Mm(150)),
    'image6': InlineImage(tpl, img1_name[2], width=Mm(150))
}
# print(context)
tpl.render(context)
tpl.save('淮南市大气细颗粒物来源解析工作.docx')
