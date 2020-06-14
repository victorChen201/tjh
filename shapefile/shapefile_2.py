# -*- coding: utf-8 -*-
import shapefile
import shapely.geometry as geometry
import pandas as pd
from mpl_toolkits.basemap import Basemap
from PIL import Image
from matplotlib.patches import Circle
from matplotlib.pyplot import MultipleLocator
from matplotlib.offsetbox import (TextArea, DrawingArea, OffsetImage,
                                  AnnotationBbox)
import pylab as plt
import os
import numpy as np
import matplotlib.patches as mpatches
FONTSIZE=24
font1 = {'family': 'Times New Roman',
        'color':  'purple',
        'weight': 'normal',
        'size': 24,
        }

def add_north(ax, labelsize=10, loc_x=0.98, loc_y=1, width=0.02, height=0.1, pad=0.12):
    """
    画一个比例尺带'N'文字注释
    主要参数如下
    :param ax: 要画的坐标区域 Axes实例 plt.gca()获取即可
    :param labelsize: 显示'N'文字的大小
    :param loc_x: 以文字下部为中心的占整个ax横向比例
    :param loc_y: 以文字下部为中心的占整个ax纵向比例
    :param width: 指南针占ax比例宽度
    :param height: 指南针占ax比例高度
    :param pad: 文字符号占ax比例间隙
    :return: None
    """
    minx, maxx = ax.get_xlim()
    miny, maxy = ax.get_ylim()
    ylen = maxy - miny
    xlen = maxx - minx
    left = [minx + xlen*(loc_x - width*.5), miny + ylen*(loc_y - pad)]
    right = [minx + xlen*(loc_x + width*.5), miny + ylen*(loc_y - pad)]
    top = [minx + xlen*loc_x, miny + ylen*(loc_y - pad + height)]
    center = [minx + xlen*loc_x, left[1] + (top[1] - left[1])*.4]
    triangle = mpatches.Polygon([left, top, right, center], color='k')
    # print(minx + maxx*loc_x,miny + maxy*loc_y)
    ax.text(s='N',
            x=minx + xlen*loc_x,
            y=miny + ylen*loc_y,
            fontsize=labelsize,
            horizontalalignment='center',
            verticalalignment='bottom')
    ax.add_patch(triangle)
def add_logo(ax,img,text="中国环境科学研究院监制",w=200,h=200,start=480):
    # Get the axes object from the basemap and add the AnnotationBbox artist
    img = img.resize((h, w), Image.ANTIALIAS)
    # print(img.size)
    plane = np.array(img)
    x = lon1
    y = lat1
    offsetbox = OffsetImage(plane, zoom=.1)
    ab = AnnotationBbox(offsetbox,xy=(x,y),
                        xybox=(start, -12),
                        xycoords='data',
                        boxcoords="offset points",
                        frameon=False)
    ax.add_artist(ab)
    offsetbox = TextArea(text, minimumdescent=False)
    ab = AnnotationBbox(offsetbox, xy=(x,y),
                        xybox=(start+70, -16),
                        xycoords='data',
                        boxcoords="offset points",
                        frameon=False,fontsize=FONTSIZE)
    ax.add_artist(ab)
def add_legend(ax,label,color='red',start=270):
    x = lon1
    y = lat1
    if label=="SUOMIVIIRS":
        start = start+110
        tmp = 42
    elif label=="MODIS":
        tmp = 30
    else:
        start = start+50
        tmp = 36
    da = DrawingArea(10, 10, 0, 0)
    p = Circle((5, 5), 5,color=color)
    da.add_artist(p)
    ab = AnnotationBbox(da, xy=(x,y),
                        xybox=(start, -14),
                        xycoords='data',
                        boxcoords="offset points",
                        frameon=False,
                        # boxcoords=("axes fraction", "data"),
                        box_alignment=(0., 0.5)
                        )
    ax.add_artist(ab)

    offsetbox = TextArea(label, minimumdescent=False)
    ab = AnnotationBbox(offsetbox, xy=(x, y),
                        xybox=(start+tmp, -16),
                        xycoords='data',
                        boxcoords="offset points",
                        frameon=False, fontsize=FONTSIZE)
    ax.add_artist(ab)
def add_title(ax,title,start=150):
    # x = lon1 + (lon2 - lon1) / 4
    x = lon1
    y = lat1
    # d1 = 0.02

    offsetbox = TextArea(title, minimumdescent=False)
    ab = AnnotationBbox(offsetbox, xy=(x, y),
                        xybox=(start, -16),
                        xycoords='data',
                        boxcoords="offset points",
                        frameon=False,fontsize=FONTSIZE)
    ax.add_artist(ab)
    # plt.text(x,y,"中国环境科学研究院监制",fontsize=10,bbox=dict(facecolor='red', alpha=0.5))
def find_point(shapefile,df):
    '''
    :param shapefile: shapefile文件
    :param df: 经纬度点
    :return: 属于shapefile所在区域的位置
    '''
    in_shape_points = []
    latitudes = []
    longitudes = []
    indexs = []
    table_data = []
    rds = shapefile.shapeRecords()
    print(rds[0])
    for index, row in df.iterrows():
        lat = row['latitude']
        lon = row['longitude']
        if lon < lon1 or lon > lon2 or lat < lat1 or lat > lat2:
            pass
        else:
            print(lon, lat)
            latitudes.append(lat)
            longitudes.append(lon)
            indexs.append(index)
            table_data.append(row.values)
            pt = (lon, lat)
            for rd in rds:
                if geometry.Point(pt).within(geometry.shape(rd.shape)):
                    in_shape_points.append(pt)
                    print(pt)
                    break
                else:
                    # not_in_shape_points.append(pt)
                    pass
    selected_lon = [elem[0] for elem in in_shape_points]
    selected_lat = [elem[1] for elem in in_shape_points]
    return (selected_lon,selected_lat),table_data
def get_centerpoint(lis):
    '''
    :param lis: 坐标列表
    :return: 重心位置
    '''
    area = 0.0
    x, y = 0.0, 0.0
    a = len(lis)
    for i in range(a):
        lat = lis[i][0]  # weidu
        lng = lis[i][1]  # jingdu
        if i == 0:
            lat1 = lis[-1][0]
            lng1 = lis[-1][1]
        else:
            lat1 = lis[i - 1][0]
            lng1 = lis[i - 1][1]
        fg = (lat * lng1 - lng * lat1) / 2.0
        area += fg
        x += fg * (lat + lat1) / 3.0
        y += fg * (lng + lng1) / 3.0
    x = x / area
    y = y / area
    return x, y
def save_table(table_data,save_name,columns):
    if len(table_data) > 0:
        col_colours = ['red']*len(columns)
        the_table1 = plt.table(cellText=table_data, colWidths=[0.15]*len(columns), colLabels=columns, colColours=col_colours, loc='center',cellLoc="left")
        the_table1.auto_set_font_size(False)
        the_table1.set_fontsize(FONTSIZE)
        the_table1.scale(1, 2)
        for cell in  the_table1._cells:
            # print(cell)
            if cell[1] ==-1:
                the_table1._cells[cell].set_edgecolor(None)
        plt.axis('off')
        # plt.show()
        plt.savefig(image_path+os.sep+save_name)
        plt.cla()

if __name__ == "__main__":
    plt.rcParams['axes.unicode_minus'] = False  # 解决保存图像是负号'-'显示为方块的问题
    plt.rcParams['font.sans-serif'] = 'STSong,Times New Roman'  # 中文设置成宋体，除此之外的字体设置成New Roman
    # plt.rcParams["font.family"]="STSong"  #解决浓度单位显示乱码情况
    # plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['savefig.dpi'] = 200  # 图片像素
    plt.rcParams['figure.dpi'] = 200  # 分辨率
    # fontsize=10
    fig = plt.figure(figsize=(20,20))
    ax = fig.add_subplot(111)
    cpath = os.path.realpath(__file__)
    cpath = os.path.dirname(cpath)
    path = cpath + os.sep + "各城市区县底图"
    csv_path = cpath + os.sep + "file"
    image_path = cpath + os.sep + "image"
    # colunms = ["latitude", "longitude", "acq_time", "confidence"]
    colunms = ["satellite", "acq_date","latitude", "longitude", "acq_time", "confidence"]
    cols = ["卫星名称", "日期", "经度", "纬度", "小时值", "置信度"]
    colunms_name = [""]
    map_file_names = os.listdir(path)
    for map_file_name in map_file_names:
        if map_file_name.endswith(".shp"):# and map_file_name.startswith("三沙市"):
            shapefile_path=path+os.sep+map_file_name.split(".")[0]
            all_list=[]
            city_name = ''
            r_city = shapefile.Reader(shapefile_path)
            for rec in enumerate(r_city.records()):
                shape = r_city.shape(rec[0])
                all_list.append(shape.bbox)
                city_name = rec[1][5]
            lon1=min([x[0] for x in all_list])
            lon2=max([x[2] for x in all_list])
            lat1=min([x[1] for x in all_list])
            lat2=max([x[3] for x in all_list])
            print(lon1,lon2,lat1,lat2)
            m = Basemap(llcrnrlon=lon1,llcrnrlat=lat1,urcrnrlon=lon2,urcrnrlat=lat2,projection = 'cyl')
            m.readshapefile(shapefile_path,'whatevername',color='black')
            # Add the plane marker at the last point.
            img = Image.open('logo_20200611212020.png')
            add_logo(ax, img)
            add_north(ax)

            file = shapefile.Reader(shapefile_path)
            date = ''
            table_data = []
            csv_files = os.listdir(csv_path)
            for file_name in csv_files:
                if file_name.endswith("Russia_Asia_24h.csv"):
                    color = "red"
                    if file_name.startswith("J1_VIIRS_C2"):
                        color = "red"
                        label = "J1VIIRS"
                    elif file_name.startswith("MODIS_C6_Russia"):
                        color = "yellow"
                        label = "MODIS"
                    else:
                        color = "orange"
                        label = "SUOMIVIIRS"
                    df = pd.read_csv(csv_path + os.sep + file_name)[colunms]
                    df["satellite"]=label
                    date = df['acq_date'][0].replace("-", "年", 1).replace("-", "月", 1)
                    selected, tb = find_point(file, df)
                    table_data = table_data + tb
                    m.scatter(selected[0], selected[1], latlon=True, s=30, marker="o", color=color,label=label)
                    add_legend(ax, label, color)
            for rec in enumerate(r_city.records()):
                shape = r_city.shape(rec[0])
                x,y = get_centerpoint(shape.points)
                fontdict={"fontsize":FONTSIZE}
                ax.text(x,y,rec[1][1],fontdict=fontdict)
            box = ax.get_position()
            add_title(ax,city_name+"24小时火点分布图"+date)
            # plt.title(city_name+"24小时火点分布图"+date,x=box.x0,y=-0.1,fontweight="normal")
            # lg = plt.legend(ncol=3,loc="upper center",frameon=False,handletextpad=0,columnspacing=0,bbox_to_anchor=(0.5,0), fontsize=12)

            plt.subplots_adjust(left=0.0, right=1,top=0.96)
            plt.axis("off")
            plt.savefig(image_path + os.sep + city_name + "24小时火点分布图" + date,bbox_inches = 'tight' ,pad_inches = 0.3)
            # plt.savefig(image_path+os.sep+city_name+"24小时火点分布图"+date)

            print(image_path+os.sep+city_name+"24小时火点分布图"+date)
            # leg = plt.gca().get_legend()
            # ltext = leg.get_texts()
            # arts = ax.artists
            # print(fig.dpi)
            # print(fig.get_size_inches())
            #
            # plt.show()
            plt.cla()
            save_table(table_data,city_name,cols)

#
