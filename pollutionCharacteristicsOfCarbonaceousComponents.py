# -*- coding: utf-8 -*-
from docxtpl import DocxTemplate, InlineImage
import pandas as pd
from matplotlib import pyplot as plt
import math
from docx.shared import Mm
import numpy as np
import pylab as pl
# plt.rcParams['font.sans-serif'] = ['SimHei'] # 指定默认字体
# plt.rcParams['font.family'] = ['SimHei AR PL UKai CN ']
plt.rcParams['axes.unicode_minus'] = False  # 解决保存图像是负号'-'显示为方块的问题
# plt.rcParams["font.family"]="STSong"  #解决浓度单位显示乱码情况
plt.rcParams['font.sans-serif'] = ['SimHei']
df = pd.read_excel(r"淮南市组分数据数据.xlsx", header=0)
inorganicElements="Na	K	Mg	Ca	Li	Be	Sc	Ti	V	Cr	Mn	Co	Ni	Cu	Zn	Ga	Ge	As	Se	Rb	Sr	Ag	Cd	Cs	Ba	Tl	Pb	Bi"
otherElements = "OC	EC	TC	SO4	NO3	Cl	NH4"

inorganicElements = inorganicElements.split('	')
otherElements = otherElements.split('	')
print(inorganicElements)

pollutant_standard = 60.0
city = df['城市'][0]
year = df['采样时间'][0].year
month = df['采样时间'][0].month
season = '冬季'



def season_date(x):
    '''
    :返回按季度分组的值
    :param df:
    :return:
    '''
    # print(x)
    month = x['采样时间'].month
    if month >= 3 and month <=5:
         s="春季"
    elif month >= 6 and month <=8:
        s = '夏季'
    elif month >= 9 and month <= 11:
        s = '秋季'
    else:
        s ='冬季'
    x['季节'] = s
    return x
    # season.append(x)

# df.drop(del_columns, axis=1, inplace=True)
# df.dropna(inplace=True, axis=0)
df.fillna(value=0,inplace=True)
columns =['OC','EC','TC'] #inorganicElements+otherElements
#总平均值
columns_mean = df[columns].apply(lambda x: x.mean().round(2))
mean_OC = columns_mean['OC']
mean_EC = columns_mean['EC']
mean_OC_EC = mean_OC/mean_EC
relate = '大于' if mean_OC_EC>2.3 else "小于"
scale_oc_ec = df['OC']/df['EC']
scale_OC_EC_min = 2.0 if min(scale_oc_ec) <2 else round(min(scale_oc_ec),2)
def soc(df):
    '''
    :功能：计算PM2.5
    :PM2_5：WSIN/[(OC*1.5+EC+WSIN)*1.13],WSIN=[SO4,NO3,C1,NH4,Na,K,Mg,Ca]
    :param df: df中每行值
    :return: 返回每行的PM2.5
    '''
    soc = df['OC'] - df['EC']*scale_OC_EC_min #df['NO3'] + df['SO4'] + df['NH4'] + df['Cl'] + df['Na'] + df['K'] + df['Mg'] + df['Ca']
    return soc.round(2)
columns_soc = df[columns].apply(lambda x: soc(x), axis=1, result_type='expand')
SOC_mean = columns_soc.mean()
OC_percent = 100*SOC_mean/df['OC'].mean()
#相关性https://jingyan.baidu.com/article/60ccbceba822f464cab1972a.html
def linefit(x , y):
    N = float(len(x))
    sx,sy,sxx,syy,sxy=0,0,0,0,0
    for i in range(0,int(N)):
        sx  += x[i]
        sy  += y[i]
        sxx += x[i]*x[i]
        syy += y[i]*y[i]
        sxy += x[i]*y[i]
    a = (sy*sx/N -sxy)/( sx*sx/N -sxx)
    b = (sy - a*sx)/N
    r = abs(sy*sx/N-sxy)/math.sqrt((sxx-sx*sx/N)*(syy-sy*sy/N))
    return a,b,r
# std_OC = df['OC'].std()
# std_EC = df['EC'].std()
# z_oc = (df['OC']-mean_OC)/std_OC
# z_ec = (df['EC']-mean_EC)/std_EC
# r = np.sum(z_oc*z_ec)/df['OC'].__len__()
a,b,r = linefit(df['EC'].values , df['OC'].values)
r = abs(r)
if r<0.3:
    relationship="低度相关"
elif r<0.8:
    relationship="中度相关"
elif r<1:
    relationship="高度相关"
else:
    relationship=""

# r = np.corrcoef(df['OC'],df['EC'])
#季节总平均值
columns_season = columns + ['季节','站点']
seasons = df.apply(lambda x: season_date(x),axis=1)
seasons = seasons[columns_season]
seasons_mean = seasons.groupby(['季节'],axis=0).mean()
OC_sort=seasons_mean.sort_values(by='OC',ascending=False)
x=""
for s in OC_sort.index:
    x = x+s+'>'
OC_season = x[:-1]
EC_sort=seasons_mean.sort_values(by='EC',ascending=False)
x=""
for s in EC_sort.index:
    x = x+s+'>'
EC_season = x[:-1]
print(seasons_mean)

#站点总平均值
# columns_site = columns + ['站点']
sites_mean = seasons.groupby(['站点'],axis=0).mean()
print(sites_mean.sort_values(by='OC',inplace=True,ascending=False))
x='其中'
y='平均值分别达到'
for site in sites_mean.index:
    x = x+site+'、'
    y = y+str(round(sites_mean.loc[site]['OC'],2))+'µg/m\u00B3'+'、'
sites_OC_mean = x[:-1] + y[:-1]
x='其中'
y='平均值分别达到'
for site in sites_mean.index:
    x = x+site+'、'
    y = y+str(round(sites_mean.loc[site]['EC'],2))+'µg/m\u00B3'+'、'
sites_EC_mean = x[:-1] + y[:-1]
#各个站点每个季节的均值
seasons_mean.reset_index(inplace=True)
seasons_mean['站点'] ='全市'
# columns_site = columns + ['站点','季节']
df1 = seasons.append(seasons_mean,sort=True)
site_season_mean = df1.groupby(['站点','季节'],axis=0).mean()
print(site_season_mean.unstack(level=1))



site_season_mean=site_season_mean.unstack(level=1)
fig,axes=plt.subplots(1, 1, figsize=(6,4))
# axes.axis('off') #去掉边框
# ax = plt.figure(figsize=(8,6))
plt.ylabel('元素浓度(µg/m\u00B3)', fontdict={'family': 'STSong', 'size': 12})
site_season_mean['EC'].plot(ax=axes,kind='bar')
plt.title('EC')
plt.subplots_adjust(bottom=0.23)
pl.xticks(rotation=360)
font1 = {'size': 8}
plt.legend(prop=font1,ncol=4)

plt.savefig("image_EC.png",bbox_inches='tight')
plt.cla()

plt.ylabel('元素浓度(µg/m\u00B3)', fontdict={'family': 'STSong', 'size': 12})
site_season_mean['OC'].plot(ax=axes,kind='bar')
plt.title('OC')
plt.subplots_adjust(bottom=0.23)
pl.xticks(rotation=360)
font1 = {'size': 8}
plt.legend(prop=font1,ncol=4)
plt.savefig("image_OC.png",bbox_inches='tight')
plt.cla()



plt.ylabel('元素浓度(µg/m\u00B3)', fontdict={'family': 'STSong', 'size': 12})
site_season_mean['TC'].plot(ax=axes,kind='bar')
plt.title('TC')
plt.subplots_adjust(bottom=0.23)
pl.xticks(rotation=360)
font1 = {'size': 8}
plt.legend(prop=font1,ncol=4)
plt.savefig("image_TC.png",bbox_inches='tight')
plt.cla()


df.plot('EC','OC',ax=axes,kind='scatter',color='g')
plt.xlabel('EC(µg/m\u00B3)', fontdict={'family': 'STSong', 'size': 12})
plt.ylabel('OC(µg/m\u00B3)', fontdict={'family': 'STSong', 'size': 12})
c = '-' if b<0 else '+'
x = df['EC']
y = df['EC']*a+b
plt.plot(x,y,color='black', linewidth=2.0, linestyle='-')
plt.text(1, 20, "y = %.2fx%c%.2f"%(a,c,abs(b)), size = 12, alpha = 0.5)
plt.text(1, 18, "r = %.2f"%r, fontdict={'family': 'STSong', 'size': 12}, alpha = 0.5)
plt.subplots_adjust(bottom=0.23)
plt.savefig("image_EC_OC.png",bbox_inches='tight') #pad_inches=0,bbox_inches='tight'去掉底部空行
plt.cla()

tpl = DocxTemplate('碳质组分污染特征_模板.docx')
context = {
    ####标题及第一部分####
    'city': city,
    'year': year,
    'mean_OC': columns_mean['OC'],
    'sites_OC_mean': sites_OC_mean,
    'mean_EC': columns_mean['EC'],
    'sites_EC_mean': sites_EC_mean,
    'site_max': sites_mean.index[0],
    'season_max': OC_sort.index[0],
    'OC_season': OC_season,
    'EC_season': EC_season,
    'image_OC': InlineImage(tpl, "image_OC.png", width=Mm(150)),
    'image_EC': InlineImage(tpl, "image_EC.png", width=Mm(150)),
    'image_TC': InlineImage(tpl, "image_TC.png", width=Mm(150)),
    'relationship':relationship,
    'r2':round(r,2),
    'mean_OC_EC': round(mean_OC_EC,2),
    'relate': relate,
    'image_EC_OC': InlineImage(tpl, "image_EC_OC.png", width=Mm(150)),
    'scale_OC_EC_min': scale_OC_EC_min,
    'mean_SOC': round(SOC_mean,2),
    'OC_percent': round(OC_percent,2),
}
# # print(context)
tpl.render(context)
tpl.save('碳质组分污染特征_模板报告.docx')
