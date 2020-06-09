# -*- coding: utf-8 -*-
import shapefile
import shapely.geometry as geometry
import pandas as pd
from mpl_toolkits.basemap import Basemap

import matplotlib
import numpy as np
import pylab as plt
import os

def draw_shapefile(df,shapefile_path=r"C:\Users\KLYG\python\citymap\安阳市\Anyang",save_name="Anyang"):
    file = shapefile.Reader(shapefile_path)
    for city_rcd in file.shapeRecords():  # 遍历每一条shaperecord
        date1 = df["acq_date"]
        csv_name = city_rcd.record[12] + r"MODIS-24小时" +date1[0].replace("/","_") +"-" + date1[date1.size-1].replace("/","_")


    lon1 = file.bbox[0]
    lat1 = file.bbox[1]
    lon2 = file.bbox[2]
    lat2 = file.bbox[3]
    print(lon1, lat1, lon2, lat2)
    m = Basemap(llcrnrlon=lon1, llcrnrlat=lat1, urcrnrlon=lon2, urcrnrlat=lat2, projection='cyl')
    in_shape_points = []
    not_in_shape_points = []
    latitudes = []
    longitudes = []
    indexs = []
    table_data = []
    rds = file.shapeRecords()
    for index,row in df.iterrows():
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
            pt = (lon,lat)
            if geometry.Point(pt).within(geometry.shape(rds[0].shape)):
                in_shape_points.append(pt)
                print(pt)
            else:
                # not_in_shape_points.append(pt)
                pass
    selected_lon = [elem[0] for elem in in_shape_points]
    selected_lat = [elem[1] for elem in in_shape_points]
    m.readshapefile(shapefile_path, 'whatevername', color='gray')
    m.scatter(selected_lon, selected_lat, latlon=True, s=30, marker="o", color='red')
    # m.scatter(114, 22.5, latlon=True, s=60, marker="o") # 这是一个测试点
    # m.drawmeridians(np.arange(10, 125, 0.5), labels=[1, 0, 0, 1])
    # m.drawparallels(np.arange(15, 30, 0.3),labels=[1,0,0,0])  #画纬度平行线
    #
    plt.axis('off')
    # plt.show()
    plt.savefig(image_path+os.sep+save_name)
    if indexs.__len__() > 0:
        plt.clf()
        row_colors = ['red']*df.shape[1]
        col_colours = ['red']*df.shape[1]
        the_table1 = plt.table(cellText=table_data, colWidths=[0.15]*df.shape[1], colLabels=df.columns, colColours=col_colours, loc='center',cellLoc="left")
        the_table1.auto_set_font_size(False)
        the_table1.set_fontsize(10)
        the_table1.scale(1, 2)
        for cell in  the_table1._cells:
            # print(cell)
            if cell[1] ==-1:
                the_table1._cells[cell].set_edgecolor(None)
        plt.axis('off')
        plt.show()
        plt.savefig(image_path+os.sep+csv_name)
        plt.clf()
if __name__ == '__main__':
    # 取得深圳的shape
    cpath = os.path.realpath(__file__)
    cpath = os.path.dirname(cpath)
    path = cpath + os.sep + "citymap"
    csv_path = cpath + os.sep + "file"
    image_path = cpath + os.sep + "image"
    filenames = os.listdir(path)
    colunms = ["latitude","longitude","acq_date","acq_time","satellite","confidence"]
    df = pd.read_csv(csv_path + os.sep +"MODIS_C6_Russia_Asia_24h.csv")[colunms]
    for filename in filenames:
        shapefile_path = path + os.sep + filename
        shapefile_name = os.listdir(shapefile_path)[0].split(".")[0]
        shapefile_path = shapefile_path + os.sep + shapefile_name
        draw_shapefile(df,shapefile_path,shapefile_name)
