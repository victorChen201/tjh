# -*- coding: utf-8 -*-
from docxtpl import DocxTemplate, RichText, InlineImage
import pandas as pd
from matplotlib import pyplot as plt
from docx.shared import Mm, Inches, Pt
import numpy as np
# plt.rcParams['font.sans-serif'] = ['SimHei'] # 指定默认字体
# plt.rcParams['font.family'] = ['SimHei AR PL UKai CN ']
plt.rcParams['axes.unicode_minus'] = False # 解决保存图像是负号'-'显示为方块的问题
# plt.rcParams["font.family"]="STSong"  #解决浓度单位显示乱码情况
plt.rcParams['font.sans-serif'] = ['SimHei']
df = pd.read_excel(r"淮南市秋季源解析数据.xlsx", header=0)

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
# df.dropna(inplace=True)
del_columns = ['Li', 'Be', 'Sc', 'Ti', 'V', 'Cr', 'Mn', 'Co', 'Ni', 'Cu', 'Zn', 'Ga', 'Ge', 'As', 'Se', 'Rb', 'Sr', 'Ag', '111Cd', '114Cd', 'Cs', 'Ba', 'Tl', '206Pb', '207Pb', '208Pb', 'Bi']
df.drop(del_columns,axis=1,inplace=True)
# df.dropna(axis=1, how='any',
#                      thresh=5, subset=None, inplace=False)
columns = ['OC', 'EC', 'TC', 'SO4', 'NO3', 'Cl', 'NH4', 'Na', 'K', 'Mg', 'Ca']
columns_sum = df[columns].apply(lambda x: x.sum().round(2))
wsin = ['SO4','NO3','Cl','NH4','Na','K','Mg','Ca']
wsin_value = columns_sum[wsin].sum(axis=0).round(2)
PM2_5 =100*wsin_value/((columns_sum['OC']*1.5 + columns_sum['EC'] + wsin_value)*1.13)

columns_mean = df[columns].apply(lambda  x:x.mean().round(2))
OC_mean = columns_mean['OC']
EC_mean = columns_mean['EC']
Water_soluble_mean = columns_mean['SO4'] + columns_mean['NO3'] + columns_mean['NH4']
columns.append('站点')
# print(df[columns])
sites_detail = ""
sites_mean = df[columns].groupby(['站点',],axis=0).mean().round(2)
print(sites_mean[wsin])
site_mean = {}
for key in sites_mean.index.values:
    sites_module = "其中%s%sOC、EC和水溶性离子浓度分别为%5.2fµg/m\u00B3、%5.2fµg/m\u00B3和%5.2fµg/m\u00B3，水溶性离子三大主要组分为NO3、NH4和SO4，浓度分别为%5.2fµg/m\u00B3、%5.2fµg/m\u00B3和%5.2fµg/m\u00B3"
    site_mean['site_name'] = key
    site_mean['OC'] = sites_mean.loc[key]['OC']
    site_mean['EC'] = sites_mean.loc[key]['EC']
    site_mean['SO4'] = sites_mean.loc[key]['SO4']
    site_mean['NO3'] = sites_mean.loc[key]['NO3']
    site_mean['NH4'] = sites_mean.loc[key]['NH4']
    site_mean['Water_soluble_mean'] = site_mean['SO4'] + site_mean['NO3'] + site_mean['NH4']
    s = sites_module%(site_mean['site_name'], season, site_mean['OC'], site_mean['EC'],site_mean['Water_soluble_mean'],site_mean['NO3'],site_mean['NH4'],site_mean['SO4'])
    sites_detail = sites_detail + ";" + s
# print(sites_detail)
mass_colunms = ['OC','SO4','NO3','NH4']
mass_concentration = columns_mean[mass_colunms].sort_values(ascending=False).index.values
# print(mass_concentration)


table_colunms = ['OC','EC','SO4','NO3','Cl','NH4']
table_others = ['Na','K','Mg','Ca']

# fig,axes=plt.subplots(1, 1, figsize=(12,12))
sites_others = sites_mean[table_others]

other = sites_others.apply(lambda x: x.sum(), axis=1)
img = sites_mean[table_colunms].copy(deep=True)
img['其它']=other
plt.figure(figsize=(8,2))
ax = plt.gca()
ax.axis('off') #去掉边框
# print(img.index.values)
col_labels = img.columns.values
row_labels = img.index.values
table_vals = []
for key in row_labels:
    # print(img.loc[key].values)
    table_vals.append(list(img.loc[key].values.round(2)))
print(col_labels)
print(row_labels)
print(table_vals)
row_colors = ['green']* len(row_labels)
col_colours = ['green']* len(col_labels)
my_table = plt.table(cellText=table_vals, colWidths=[0.12] * len(col_labels),
                     rowLabels=row_labels, colLabels=col_labels,
                     rowColours=row_colors, colColours=col_colours,
                     loc='best')
# cells = my_table._cells
# print(my_table)
# for cell in my_table._cells:
#     print(cell)
#     if cell == (0,-1):
#         # my_table._cells[cell].set('test')
#         print(dir(my_table._cells[cell]))
my_table.auto_set_font_size(False)
my_table.set_fontsize(12)
my_table.scale(1, 2)

plt.savefig('不同站点PM2.5化学组成浓度情况.png')
plt.cla()
print(img.T)
img.T.plot(kind='bar')
plt.ylabel('质量浓度(µg/m\u00B3)',fontdict={'family' : 'STSong', 'size'   : 12})
# plt.show()
plt.savefig('不同站点PM2.5化学组成.png')
"""
文档第一部分
一、  当前排名形势
city:城市
year：年份
season：季节
PM2_5：WSIN/[(OC*1.5+EC+WSIN)*1.13],WSIN=[SO4,NO3,C1,NH4,Na,K,Mg,Ca]
"""
# city = '淮南市'
# year = '年份'
# season = '季节'
# PM2_5 = 60


"""
文档第二部分
二、  季节
OC_mean：OC平均值
EC_mean：EC平均值
Water_soluble_mean：水溶性离子平均值,NO3、SO4、NH4,CI,Na,K,Mg,Ca
site_1：站点1
site1_OC_mean：OC平均值
site1_EC_mean：EC平均值
site1_Water_soluble_mean：水溶性离子平均值
N03_1: 硝酸盐
SO4_1：硫酸盐
NH4_1：铵盐
site_2: 站点2
site2_OC_mean：OC平均值
site2_EC_mean：EC平均值
site2_Water_soluble_mean：水溶性离子平均值
N03_2: 硝酸盐
SO4_2：硫酸盐
NH4_2：铵盐
site_3: 站点3
site3_OC_mean：OC平均值
site3_EC_mean：EC平均值
site3_Water_soluble_mean：水溶性离子平均值
N03_3: 硝酸盐
SO4_3：硫酸盐
NH4_3：铵盐
"""


"""
文档第三部分
三、 {{有机碳}}、{{硝酸盐}}、{{氨盐}}和{{硫酸盐}}（依次为质量浓度由大到小）
"""
# mass_concentration1 = 'NO3'
# mass_concentration2 = 'SO4'
# mass_concentration3 = 'NH4'
# mass_concentration4 = 'OC'

"""
文档第四部分
四、  图表
"""
img1 = '不同站点PM2.5化学组成浓度情况.png'
img2 = '不同站点PM2.5化学组成.png'
"""
文档第四部分
五、 该市未来24小时管控重点
"""

tpl = DocxTemplate('季节颗粒物化学组分.docx')

context = {
    ####标题及第一部分####
    'city': city,
    'year': year,
    'season': season,
    'PM2_5': PM2_5.round(2),

    #####第二部分####
    'OC_mean': OC_mean,
    'EC_mean': EC_mean,
    'Water_soluble_mean': Water_soluble_mean,
    'sites_detail': sites_detail,

    #####第三部分####
    'mass_concentration1': mass_concentration[0],
    'mass_concentration2': mass_concentration[1],
    'mass_concentration3': mass_concentration[2],
    'mass_concentration4': mass_concentration[3],
    'image1': InlineImage(tpl, img1, width=Mm(150)),
    'image2': InlineImage(tpl, img2, width=Mm(150))

}
# print(context)
tpl.render(context)
tpl.save('生成季节颗粒物化学组分.docx')
