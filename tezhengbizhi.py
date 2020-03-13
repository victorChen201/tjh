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


# df.dropna(axis=1, how='any', thresh=1, subset=None, inplace=True)
columns = ['OC', 'EC', 'SO4', 'NO3', 'Cl', 'NH4', 'Na', 'K', 'Mg', 'Ca']
columns.append('站点')
# print(df[columns])

sites_detail = ""
sites_mean = df[columns].groupby(['站点',],axis=0).mean().round(2)

sites = []
siteSN = []
siteOE = []
imageSN_name = "image_sn.png"
imageOE_name = "image_oe.png"
SN_ymax = 0
OE_ymax = 0
for key in sites_mean.index.values:
    sn = round(sites_mean.loc[key]['SO4']/sites_mean.loc[key]['NO3'],2)
    oe = round(sites_mean.loc[key]['OC']/sites_mean.loc[key]['EC'],2)
    if sn > SN_ymax:
        SN_ymax = sn
    if oe > OE_ymax:
        OE_ymax = oe
    siteSN.append(sn)
    siteOE.append(oe)
    sites.append(key)
    print(key,sn,oe)

plt.figure()

# ax.axis('off') #去掉边框
# print(img.index.values)
plt.cla()
bar_width = 0.3
plt.ylim(0,int(SN_ymax+1))
# plt.subplots_adjust(bottom=0.14)
plt.bar(sites, siteSN, width=bar_width, alpha=0.4, color='b',label=sites)
plt.ylabel('比值')
plt.title("SO4/NO3")
plt.savefig(imageSN_name)
plt.cla()
plt.ylim(0,int(OE_ymax+1))
plt.bar(sites, siteOE, width=bar_width, alpha=0.4, color='b',label=sites)
plt.ylabel('比值')
plt.title("OC/EC")
plt.savefig(imageOE_name)
siteOE.sort(reverse=False)
siteSN.sort(reverse=False)
tpl = DocxTemplate('特征比值_模板.docx')
context = {
    ####标题及第一部分####
    'city': city,
    'season': season,
    'site1': sites[0],
    'site2': sites[1],
    'site3': sites[2],
     #increase_fastest_pollutant
    'sn1': siteSN[0],
    'sn2': siteSN[1],
    'sn3': siteSN[2],
    #increase_fastest_pollutant percent
    'sn_min': int(siteSN[0]+1),

    #减少最多的成分
    'oe1': siteOE[0],
    'oe2': siteOE[1],
    'oe3': siteOE[2],
    'oe_max': int(siteOE[-1]),
    #####第二部分####
    'oe_min': int(siteOE[0]),
    'imageSN': InlineImage(tpl, imageSN_name, width=Mm(150)),
    'imageOE': InlineImage(tpl, imageOE_name, width=Mm(150))
}
# print(context)
tpl.render(context)
tpl.save('特征比值生成.docx')
