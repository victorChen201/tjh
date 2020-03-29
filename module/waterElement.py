# -*- coding: utf-8 -*-
import pandas as pd
from docxtpl import DocxTemplate, RichText, InlineImage
from matplotlib import pyplot as plt
import math
import numpy as np
from matplotlib.table import Table
from docx.shared import Mm, Inches, Pt
plt.rcParams['axes.unicode_minus'] = False  # 解决保存图像是负号'-'显示为方块的问题
# plt.rcParams["font.family"]="STSong"  #解决浓度单位显示乱码情况
plt.rcParams['font.sans-serif'] = ['SimHei']
df = pd.read_excel(r"淮南市组分数据数据.xlsx", header=0)

wsin = ['SO4','NO3','Cl','NH4','Na','K','Mg','Ca','季节','站点']
pm25 = ['SO4','NO3','Cl','NH4','Na','K','Mg','Ca','OC','EC','季节','站点']
def func1(df):
    '''
    :功能：计算水溶性离子占PM25的浓度
    :PM2_5：WSIN/[(OC*1.5+EC+WSIN)*1.13],WSIN=[SO4,NO3,C1,NH4,Na,K,Mg,Ca]
    :param df: df中每行值
    :return: 返回每行的PM2.5
    '''
    wsin = ['SO4', 'NO3', 'Cl', 'NH4', 'Na', 'K', 'Mg', 'Ca']
    pm25 = ['SO4', 'NO3', 'Cl', 'NH4', 'Na', 'K', 'Mg', 'Ca', 'OC', 'EC']
    r = 100*df[wsin].sum()/df.sum()#df['NO3'] + df['SO4'] + df['NH4'] + df['Cl'] + df['Na'] + df['K'] + df['Mg'] + df['Ca']
    return r
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
def percent(x):
    # print(x)
    p = 100*x/x.sum()
    return p
def checkerboard_table(data, fmt='{:.2f}', bkg_colors=['yellow', 'white']):
    fig, ax = plt.subplots(figsize=(8,6))
    ax.set_axis_off()
    # bbox = [0, 0, 2, 2]
    tb = Table(ax)
    # tb.FONTSIZE=24
    # bbox = [0, 0, 1, 1]
    nrows, ncols = data.shape
    width, height = 1.0 / ncols, 1.0 / nrows
    # Add cells
    for (i, j), val in np.ndenumerate(data):
        # Index either the first or second item of bkg_colors based on
        # a checker board pattern
        idx = [j % 2, (j + 1) % 2][i % 2]
        color = bkg_colors[idx]
        tb.add_cell(i, j, width, height, text=val if isinstance(val,str) else fmt.format(val),
                    loc='center', facecolor='none')

    # Row Labels...
    for i, label in enumerate(data.index):
        tb.add_cell(i, -1, width, height, text=label, loc='right',
                     facecolor='g')
        # Column Labels...
    for j, label in enumerate(data.columns):
        label = label.split('_')[0]
        if j == len(data.columns)-2:
            label = "监测站（全部）"
        if j%2 == 0:
            tb.add_cell(-1, j, width, height, text=label, loc='center',
                        facecolor='r')
            ax.add_table(tb)
    return fig
def fuc2(ax,labels):
    '''
    :功能，
    :param ax:
    :return:
    '''
    num = 0
    x_num = 0
    y_num = 0
    for item in ax[0]:
        center = item.center
        theta1 = item.theta1
        theta2 = item.theta2
        theta = theta1 + (theta2 - theta1) / 2
        xytext_y = 0
        xytext_x = 0
        xy_x = 0
        xy_y = 0
        r = item.r
        x = r * math.cos(math.pi * theta / 180)
        y = r * math.sin(theta * math.pi / 180)
        # if theta > 360:
        #     theta = theta - 360
        #     x = r * math.cos(math.pi * theta / 180)
        #     y = r * math.sin(theta * math.pi / 180)
        # elif theta > 270:
        #     theta = 360 - theta
        #     x = r * math.cos(math.pi * theta / 180)
        #     y = (-1) * r * math.sin(theta * math.pi / 180)
        # elif theta > 180:
        #     theta = theta - 180
        #     x = (-1) * r * math.cos(theta * math.pi / 180)
        #     y = (-1) * r * math.sin(theta * math.pi / 180)
        # elif theta > 90:
        #     theta = 180 - theta
        #     x = (-1) * r * math.cos(theta * math.pi / 180)
        #     y = r * math.sin(theta * math.pi / 180)
        # else:
        #     x = r * math.cos(theta * math.pi / 180)
        #     y = r * math.sin(theta * math.pi / 180)
        if (theta2 - theta1) > 20:
            ax[1][num]._text = "\n".join(labels[num].split(' '))
            num = num + 1
            continue

        # if x < 0:
        #     # print(labels[num],x,y)
        #     num = num + 1
        #     continue
        if y < 0:
            xytext_y = -1.0 + (0.1) * y_num
            xytext_x = center[0] + r + 0.1 * y_num
            xy_x = center[0] + x
            xy_y = center[1] + y
            y_num = y_num + 1
            if y_num > 6 and x_num < 6:
                xytext_y = 0.8 + 0.2 * x_num
                xytext_x = center[0] + r + 0.6 - 0.02 * x_num
                xy_x = center[0] + x
                xy_y = center[1] + y
                x_num = x_num + 1
        else:
            xytext_y = 0.6 + 0.1 * x_num
            xytext_x = center[0] + r + 0.6 - 0.02 * x_num
            xy_x = center[0] + x
            xy_y = center[1] + y
            x_num = x_num + 1
        print(labels[num], xy_x, xy_y, xytext_x, xytext_y)
        plt.annotate("", xytext=(xytext_x, xytext_y), xy=(xy_x, xy_y), color="r", weight="bold",
                     arrowprops=dict(arrowstyle="-", connectionstyle="arc3", color="black"), horizontalalignment='left',
                     verticalalignment='bottom')
        plt.annotate(labels[num], xytext=(xytext_x + 0.1, xytext_y), xy=(xytext_x, xytext_y), color="r", weight="bold",
                     arrowprops=dict(arrowstyle="-", connectionstyle="arc3", color="black"), horizontalalignment='left',
                     verticalalignment='bottom')
        num = num + 1
tpl = DocxTemplate('水溶性离子原件.docx')
context = {
    ####标题及第一部分####
    'season1_title': '',
    'season1_content1': '',
    'season1_content2': '',
    'season1_content3': '',
    'season1_content4': '',
    'season1_image1_title': "",
    'season1_image2_title': "",
    'season1_image3_title': "",
    'season1_image4_title': "",
    'season1_image1': '',
    'season1_image2': '',
    'season1_image3': '',
    'season1_image4': '',

}
df = df.apply(lambda x: season_date(x),axis=1)
df.fillna(value=0, inplace=True)
columns = wsin
all_sites_table = df[columns].groupby(['季节'],axis=0).mean()
season_group = df[pm25].groupby(['季节'],axis=0,)
all_seasons = []
tz = {}
num1 = 1
for key,season in season_group:
    context['season%i_title'%num1] = key+"PM2.5中水溶性无机离子浓度特征"
    context["season%i_image1_title"%num1] = "表6.3-2 不同站点水溶性无机离子浓度值及在TWSI中所占百分比单位：μg/m\u00B3"
    print(key,season)
    season_pm25 = season.copy(deep=True)
    season = season[wsin]
    #各站点水溶性离子占PM25的百分比
    sites_pm25 = season_pm25.groupby(['站点'], axis=0).mean()
    sites_pm25_percent = sites_pm25.apply(lambda x:func1(x),axis=1,result_type='expand')
    tmp = ''
    for i in range(sites_pm25_percent.values.size):
        if i == 0:
            tmp = str(round(sites_pm25_percent.values[i],2))+"%"
        elif i == sites_pm25_percent.values.size-1:
            tmp = tmp+'和'+str(round(sites_pm25_percent.values[i],2))+"%"
        else:
            tmp = tmp + '、' + str(round(sites_pm25_percent.values[i], 2))+"%"

    context['season%i_content1' % num1] = "%s水溶性离子占PM2.5分别为%s（表6.3-2）。"%('、'.join(sites_pm25_percent.index),tmp)

    season.drop('季节', inplace=True, axis=1)
    sites_table = season.groupby(['站点'],axis=0).mean().T
    # print(sites_table)
    percent_table = sites_table.apply(lambda x: percent(x), axis=0)
    sites_combin=''
    num2=2
    for k in sites_table.columns:
        context['season%i_content%i' % (num1,num2)] = "%s最主要的水溶性无机离子NO3\u207B、SO4\u00B2\u207B和NH4+浓度分别为：%.2fμg/m\u00B3、%.2fμg/m\u00B3、%.2fμg/m\u00B3，占TWSI的%.2f%%，%.2f%%，%.2f%%，其余5种水溶性无机离子之和占TWSI的%.2f%%（图6.3-6）。"\
                                                      %(k,sites_table[k].loc['NO3'],sites_table[k].loc['SO4'],sites_table[k].loc['NH4'],percent_table[k].loc['NO3'],percent_table[k].loc['SO4'],percent_table[k].loc['NH4'],percent_table[k].loc[['Cl','Na','K','Mg','Ca']].sum())
        context['season%i_image%i_title' % (num1,num2)] = "图6.3-6 %sPM2.5中水溶性无机离子所占的百分比" % k

        combin = pd.merge(sites_table[k],percent_table[k],left_index=True,right_index=True,how='inner')
        if isinstance(sites_combin,str):
            sites_combin = combin
        else:
            sites_combin = pd.merge(sites_combin, combin, left_index=True, right_index=True, how='inner')
        plt.figure(figsize=(10, 6))
        labels = percent_table[k].index
        # print(percent_table)
        explode = [0] * (percent_table[k].size)
        values = percent_table[k].values
        # startangle=10,
        ax = plt.pie(x=values,labels=labels, radius=1, center=(0, 0), explode=explode, startangle=0,autopct='%3.1f%%',
                     labeldistance=1.2)
        # fuc2(ax,labels)
        plt.title(k)
        plt.savefig(key+'_'+k+'.png')
        context['season%i_image%i' % (num1,num2)] = InlineImage(tpl, key+'_'+k+'.png', width=Mm(150))
        num2 = num2 + 1
        # plt.show()

    p = 100*all_sites_table.loc[key]/all_sites_table.loc[key].sum()
    total_site = pd.merge(all_sites_table.loc[key].T,p.T,left_index=True,right_index=True,how='inner')
    # combin.columns = combin.columns.sort_values(ascending=True)
    combin = pd.merge(sites_combin,total_site,left_index=True,right_index=True,how='inner')
    # combin.columns.sort_values(ascending=True)
    v = ("平均值,百分比%," * (sites_table.shape[1] + 1)).split(',')
    df1 = pd.DataFrame([v[0:-1]], columns=combin.columns,index=["水溶性离子"])
    df1 = df1.append(combin,ignore_index=False)
    checkerboard_table(df1)
    plt.savefig(key+'.png')
    context['season%i_image1' % num1] = InlineImage(tpl, key+'.png', width=Mm(150))
    num1 = num1+1

# # print(context)
tpl.render(context)
tpl.save('水溶性离子.docx')