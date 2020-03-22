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
df = pd.read_excel(r"淮南市组分数据数据.xlsx", header=0)
inorganicElements="Na	K	Mg	Ca	Li	Be	Sc	Ti	V	Cr	Mn	Co	Ni	Cu	Zn	Ga	Ge	As	Se	Rb	Sr	Ag	Cd	Cs	Ba	Tl	Pb	Bi"
otherElements = "OC	EC	TC	SO4	NO3	Cl	NH4"

inorganicElements = inorganicElements.split('	')
otherElements = otherElements.split('	')
print(inorganicElements)

pollutant_standard = 60.0
city = df['城市'][0]
# year = df['采样时间'][0].year
# month = df['采样时间'][0].month
# season = '冬季'

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
columns = inorganicElements+otherElements
columns_mean = df[columns].apply(lambda x: x.mean().round(2))
#总浓度
total_concentration = columns_mean[inorganicElements].sum()
print(total_concentration)
inorganicElementsTotalPercent = 100*total_concentration/columns_mean.sum()
print(inorganicElementsTotalPercent)
#化学成分平均值
#无机元素浓度前八的名字及占总浓度值
head8 = columns_mean[inorganicElements].sort_values(ascending=False).head(8)
head8_total = 100*head8.sum()/total_concentration
head8_concentration = head8/total_concentration
head8_element_names = head8.index
head8_element_values = head8.values
head8_element_percent = head8_concentration.values
#季节变化
columns = inorganicElements + ['季节']
seasons = df.apply(lambda x: season_date(x),axis=1)
seasons_mean = seasons[columns].groupby(['季节'],axis=0).mean()
print(seasons_mean)
seasons_total = seasons_mean.apply(lambda x: x.sum(),axis=1)
# print(df[columns])
seasons_total.sort_values(ascending=False,inplace=True)
k=''
v=''
for i in range(seasons_total.size):
    if (i !=0) and (i==seasons_total.size-1):
        k = k+'和'+seasons_total.index[i]
        v = '%s和%.2fng/m\u00B3'%(v,seasons_total.values[i])
    elif i==0:
        k = seasons_total.index[i]
        v = '%.2fng/m\u00B3'%seasons_total.values[i]
    else:
        k = k + '、'+ seasons_total.index[i]
        v = '%s、%.2fng/m\u00B3'%(v,seasons_total.values[i])
#站点变化
columns = inorganicElements + ['站点']
sites_mean = df[columns].groupby(['站点'],axis=0).mean()
print(sites_mean)
sites_total = sites_mean.apply(lambda x: x.sum(),axis=1)
# print(df[columns])
sites_total.sort_values(ascending=False,inplace=True)
sites_print = ''
sites_printStart = ''
sites_printOther = ''
sites_printEnd = ''
for i in range(sites_total.size):
    if (i != 0) and (i == sites_total.size - 1):
        sites_printEnd = '%s最低，为%.2fng/m\u00B3'%( sites_total.index[i], sites_total.values[i])
    elif i == 0:
        sites_printStart = "%s无机元素总浓度略高于其他监测点，为%.2fng/m\u00B3"%(sites_total.index[i],sites_total.values[i])
    elif i>1:
        sites_printOther = '%s、%s(%.2fng/m\u00B3)'%(sites_printOther, sites_total.index[i], sites_total.values[i])
    else:
        sites_printOther = '，%s(%.2fng/m\u00B3)'%( sites_total.index[i], sites_total.values[i])
sites_print = sites_printStart + sites_printOther +'次之,' + sites_printEnd
#轻重污染日的对比
def func1(x):
    '''
    :功能：计算PM2.5,wsin = ['SO4','NO3','Cl','NH4','Na','K','Mg','Ca']
    :PM2_5：WSIN/[(OC*1.5+EC+WSIN)*1.13],WSIN=[SO4,NO3,C1,NH4,Na,K,Mg,Ca]
    :param x:
    :return:
    '''
    wsin = ['SO4', 'NO3', 'Cl', 'NH4', 'Na', 'K', 'Mg', 'Ca']
    wsin_v = x[wsin].sum()  # df['NO3'] + df['SO4'] + df['NH4'] + df['Cl'] + df['Na'] + df['K'] + df['Mg'] + df['Ca']
    y = (x['OC']*1.5 + x['EC'] + wsin_v)*1.13
    pm2_5 = 100*wsin_v/y
    # df['PM2_5'] = pm2_5.round(2)
    if pm2_5>pollutant_standard:
        x['污染严重度'] = '重'
    else:
        x['污染严重度'] = '轻'
    return x
df_add_pm2_5 = df.apply(lambda x:func1(x),axis=1)
pm2_5_mean = df_add_pm2_5.groupby(['污染严重度'],axis=0).mean()
pm2_5_total = pm2_5_mean[inorganicElements].apply(lambda x:x.sum(),axis=1)
ratop = pm2_5_total['重']/pm2_5_total['轻']
p = pm2_5_mean[inorganicElements].div(pm2_5_total,axis=0)
elements_ratop = pm2_5_mean.div(pm2_5_mean.loc['轻'],axis=1).loc['重']
elements_ratop.sort_values(ascending=False,inplace=True)
elements_stables = abs(elements_ratop - 1)
elements_stables.sort_values(ascending=False,inplace=True)
# print(sites_detail)
# fig,axes=plt.subplots(3, 1, figsize=(12,12))
plt.figure(figsize=(8,6))
plt.ylabel('元素浓度(µg/m\u00B3)', fontdict={'family': 'STSong', 'size': 12})
columns_mean.plot(kind='bar')
plt.savefig("image1.png")
plt.cla()

plt.ylabel('元素浓度(µg/m\u00B3)', fontdict={'family': 'STSong', 'size': 12})
seasons_total.plot(kind='bar')
plt.savefig("image3.png")
plt.cla()

print(seasons_mean)
for line in seasons_mean.index:
    plt.plot(seasons_mean.columns,seasons_mean.loc[line].values,marker='*',label=line)
plt.ylabel('元素浓度(µg/m\u00B3)', fontdict={'family': 'STSong', 'size': 12})
plt.legend()
plt.savefig("image4.png")
plt.cla()

sites_total.plot(kind='bar')
plt.ylabel('元素浓度(µg/m\u00B3)', fontdict={'family': 'STSong', 'size': 12})
plt.savefig("image5.png")
plt.cla()

for line in sites_mean.index:
    plt.plot(sites_mean.columns,sites_mean.loc[line].values,marker='*',label=line)
plt.ylabel('元素浓度(µg/m\u00B3)', fontdict={'family': 'STSong', 'size': 12})
plt.legend()
plt.savefig("image6.png")
plt.cla()

elements_ratop.plot(kind='bar')
plt.ylabel('重污染日与轻污染日元素浓度比值', fontdict={'family': 'STSong', 'size': 12})
plt.savefig("image7.png")
plt.cla()

explode = [0]*9
y=np.append(head8_element_names,'其他')
x=np.append(head8_element_percent,1-head8_element_percent.sum())
textprops= {'fontsize':10,'color':'r'}
plt.pie(x=x, labels=y,radius=0.8,center=(0,0),explode=explode, textprops=textprops, autopct='%3.1f%%')
plt.legend(head8_element_names,loc='upper right')
plt.savefig("image2.png")
plt.close()

for key_name in pm2_5_mean.index:
    plt.figure(figsize=(10, 6))

    pt=pm2_5_mean.loc[key_name][inorganicElements]#p.loc["重"]
    if key_name == "重":
        startangle = 10
    else:
        startangle = 30
    # columns1 ="Na	K	Mg	Ca	Li	Be	Sc	Ti	V	Cr	Mn	Co	Ni	Cu	Zn	Ba	Tl	Pb"
    # columns1 = columns1.split('	')
    # others = "Ga	Ge	As	Se	Rb	Sr	Ag	Cd	Cs	Bi"
    # others = others.split('	')
    index = pt[inorganicElements].sort_values(ascending=False).index
    others = index[-14:]
    columns1 = index[0:-14]
    indexs =[columns1,others]
    others_sum = pt[others].sum()

    for i in range(0,2):
        x = pt[indexs[i]]
        # x.sort_values(ascending=False,inplace=True)
        if i ==0:
            values= np.append(x.values,others_sum)
            # explode[i] = 0.06
            labels = np.append(x.index,'其它')
        else:
            values = x.values
            labels = x.index
        labelsAddPercent = []
        for label,value in zip(labels,values):
            labelsAddPercent.append("%s %.2f%%"%(label,100*value/pt.sum()))
        labels = labelsAddPercent
        print(labelsAddPercent)
        explode = [0]*(values.size)
        #startangle=10,
        ax = plt.pie(x=values,radius=1/(i+1),center=(i*2,0),explode=explode, startangle=startangle, labeldistance=0.6)
        # axes[i].set_title("第" + str(i + 1) + "张图")
        # plt.subplots_adjust(left=None, bottom=None, right=None, top=None, wspace=0.5, hspace=1)  # 调整子图间距
        # plt.xticks(())
        # plt.yticks(())
        num=0
        x_num = 0
        y_num = 0
        for item in ax[0]:
            center = item.center
            theta1 = item.theta1
            theta2 = item.theta2
            theta = theta1 + (theta2 - theta1)/2
            xytext_y = 0
            xytext_x = 0
            xy_x = 0
            xy_y = 0
            r = item.r
            if theta > 360:
                theta = theta-360
                x = r * math.cos(math.pi * theta / 180)
                y = r * math.sin(theta * math.pi / 180)
            elif theta>270:
                theta = 360-theta
                x = r * math.cos(math.pi*theta/180)
                y  = (-1)*r * math.sin(theta*math.pi/180)
            elif theta >180:
                theta = theta-180
                x = (-1)*r * math.cos(theta*math.pi/180)
                y = (-1) * r * math.sin(theta*math.pi/180)
            elif theta>90:
                theta=180-theta
                x = (-1)*r * math.cos(theta*math.pi/180)
                y =   r * math.sin(theta*math.pi/180)
            else:
                x = r * math.cos(theta*math.pi/180)
                y = r * math.sin(theta*math.pi/180)
            if (theta2 - theta1)>20:
                ax[1][num]._text = "\n".join(labels[num].split(' '))
                num = num + 1
                continue

            if x <0:
                # print(labels[num],x,y)
                num = num + 1
                continue
            if y<0:
                xytext_y = -1.0+(0.1)*y_num
                xytext_x = center[0]+r+0.1*y_num
                xy_x = center[0]+x
                xy_y = center[1]+y
                y_num=y_num+1
                if y_num >6 and x_num <6:
                    xytext_y = 0.8 + 0.2 * x_num
                    xytext_x = center[0] + r +0.6 - 0.02 * x_num
                    xy_x = center[0] + x
                    xy_y = center[1] + y
                    x_num = x_num + 1
            else:
                xytext_y = 0.6 + 0.1 * x_num
                xytext_x = center[0] + r + 0.6 - 0.02 * x_num
                xy_x = center[0] + x
                xy_y = center[1] + y
                x_num = x_num+1
            print(labels[num],xy_x,xy_y,xytext_x,xytext_y)
            plt.annotate("", xytext=(xytext_x,xytext_y), xy=(xy_x, xy_y), color="r", weight="bold",
                         arrowprops=dict(arrowstyle="-", connectionstyle="arc3", color="black"),horizontalalignment='left', verticalalignment='bottom')
            plt.annotate(labels[num], xytext=(xytext_x+0.1, xytext_y), xy=(xytext_x,xytext_y), color="r", weight="bold",
                         arrowprops=dict(arrowstyle="-", connectionstyle="arc3", color="black"), horizontalalignment='left',
                         verticalalignment='bottom')

            if '其它' in labels[num]  :
                plt.annotate("", xy=(1.9,0.48), xytext=(xy_x, xy_y),color="r",weight="bold",arrowprops=dict(arrowstyle="-",connectionstyle="arc3",color="black"))
                plt.annotate("", xy=(1.6,-0.30), xytext=(xy_x, xy_y),color="r",weight="bold",arrowprops=dict(arrowstyle="-",connectionstyle="arc3",color="black"))
            num = num + 1
    plt.subplots_adjust(left=0.3)
    # plt.xlim(0,2)
    # plt.ylim(0,2)
    # plt.legend()# 调整子图间距
    plt.title(key_name+"污染日各无机元素浓度百分比")
    plt.savefig(key_name+'.png')
    # plt.show()

tpl = DocxTemplate('无机元素模板.docx')
context = {
    ####标题及第一部分####
    'city': city,
    'total_concentration': total_concentration,
    'inorganic_elements_percent': round(inorganicElementsTotalPercent,2),
    'element1': head8_element_names[0],
    'element2': head8_element_names[1],
    'element3': head8_element_names[2],
    'element4': head8_element_names[3],
    'element5': head8_element_names[4],
    'element6': head8_element_names[5],
    'element7': head8_element_names[6],
    'element8': head8_element_names[7],
    'head8_total': round(head8_element_values[0],2),
    'head_concentration1': round(head8_element_values[0],2),
    'head_concentration2': round(head8_element_values[1],2),
    'head_concentration3': round(head8_element_values[2],2),
    'head_concentration4': round(head8_element_values[3],2),
    'head_concentration5': round(head8_element_values[4],2),
    'head_concentration6': round(head8_element_values[5],2),
    'head_concentration7': round(head8_element_values[6],2),
    'head_concentration8': round(head8_element_values[7],2),

    "head8_percent1":round(head8_element_percent[0],2),
    "head8_percent2":round(head8_element_percent[1],2),
    "head8_percent3":round(head8_element_percent[2],2),
    "head8_percent4":round(head8_element_percent[3],2),
    "head8_percent5":round(head8_element_percent[4],2),
    "head8_percent6":round(head8_element_percent[5],2),
    "head8_percent7":round(head8_element_percent[6],2),
    "head8_percent8":round(head8_element_percent[7],2),
    #无机元素总浓度变化
    "season_head1":seasons_total.index[0],
    "season_head2":seasons_total.index[1],
    "season_head3":seasons_total.index[-1],
    "seasons":k,
    "concentrations":v,
    #1.3.1总浓度的空间变化
    "sites_concentration": sites_print,
    #冬日重污染日与轻污染日无机元素浓度比较
    "pollutant_standard":pollutant_standard,
    "element_high_average":round(pm2_5_total['重'],2),
    "element_low_average":round(pm2_5_total['轻'],2),
    "element_ratop": round(ratop,2),
    "element_increase_highest1":elements_ratop.index[0],
    "element_increase_highest2":elements_ratop.index[1],
    "element_increase_highest3":elements_ratop.index[2],
    "element_increase_highest":elements_ratop.index[0],
    "element_decrease_low1":elements_ratop.index[-1],
    "element_decrease_low2":elements_ratop.index[-2],
    "element_decrease_low3":elements_ratop.index[-3],
    "element_stable1":elements_stables.index[-1],
    "element_stable2":elements_stables.index[-2],
    "element_stable3":elements_stables.index[-3],
    'image1': InlineImage(tpl, "image1.png", width=Mm(150)),
    'image2': InlineImage(tpl, "image2.png", width=Mm(150)),
    'image3': InlineImage(tpl, "image3.png", width=Mm(150)),
    'image4': InlineImage(tpl, "image4.png", width=Mm(150)),
    'image5': InlineImage(tpl, "image5.png", width=Mm(150)),
    'image6': InlineImage(tpl, "image6.png", width=Mm(150)),
    'image7': InlineImage(tpl, "image7.png", width=Mm(150)),
    'image8': InlineImage(tpl, "重.png", width=Mm(150)),
    'image9': InlineImage(tpl, "轻.png", width=Mm(150)),
}
# # print(context)
tpl.render(context)
tpl.save('无机元素分析报告.docx')
