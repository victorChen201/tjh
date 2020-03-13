# -*- coding: utf-8 -*-
from docxtpl import DocxTemplate, RichText, InlineImage
import jinja2
from docx.shared import Mm, Inches, Pt
# import pymysql
import pandas as pd
import numpy as np

df =pd.read_excel(r"淮南市秋季源解析数据.xlsx",header=0)

"""
标题部分：
period:第几期
report_year报告生成年份
report_month报告生成月份
report_day:报告生成日
city_name:城市名
season_value：月份所属季节
"""
    
period=317
report_year='2019'
report_month='12'
report_day='26'
city_name='昌吉州'



"""
文档第一部分
一、  当前排名形势
before_day:报告前一天
composite_index：综合指数
composite_target：综合指数增加、综合指数减少
composite_value：综合指数增加了多少或减少了多少

o3_target：臭氧浓度升高或降低
o3_value：臭氧浓度升高或降低了多少

fine_days：优良天数
fine_target：优良天数度增加或减少
fine_value：优良天数度增加或减少了多少

"""
img_1 = '昌吉州_20191226_月平均.png'
before_day='25'
composite_index='5.56'
composite_target='减少'
composite_value='0.01'

o3_target='降低'
o3_value='1'

rank_value='30'
compare_city='乌鲁木齐市'
rank_target='差于'
compare_rank='32'
compare_composite_index='0.19'

fine_days='193'
fine_target='增加'
fine_value='25'

compare_target='少'
compare_day='13'

    


"""
文档第二部分
二、  年度及秋冬季目标完成情况
total_pm25_value：PM2.5年累计浓度
later_year：后一年  当前年份的后一年
avg_pm25_value：秋冬季PM2.5平均浓度
compare_pm25_period：
"""  
pm25_target='55'
total_pm25_value='53'
avg_pm25_value='55'
differ_pm25value='2'
later_year='2020'    
season={'1':'秋冬季','2':'','3':'','4':'','5':'','6':'','7':'','8':'秋冬季','9':'秋冬季','10':'秋冬季','11':'秋冬季','12':'秋冬季'}
season_value=season[report_month]
avg_pm25_value='55'
avg_pm25_target='下降'
pm25_target_value='8.3%'
heavy_pollute_day='1'
heavy_pollute_target='减少'
heavy_pollute_day_last='1'





"""
文档第三部分
三、  昨日空气质量回顾
"""
img_2='昌吉州-时间雷达图.jpg'
aqi_value='65'
quality_value='良'
primary_poll='NO2'
no2_day_value='52'
pm10_day_value='70'
pm25_day_value='24'
poll_type='偏二次'

lower_time='早上'
lower_clock='7'
lower_value='48'
upper_time='晚上'
upper_clock='20'
upper_value='77'
    


"""
文档第四部分
四、  空气质量预报
"""


"""
文档第四部分
五、 该市未来24小时管控重点
"""


tpl = DocxTemplate('自动化模板-1226.docx')


context = {
####标题及第一部分####
    'image_1': InlineImage(tpl, img_1),
    'city_name':city_name,
    'report_year': report_year,
    'report_month':report_month,
    'report_day':report_day,
    'before_day':before_day,
    'composite_index':composite_index,
    'composite_target':composite_target,
    'composite_value':composite_value,

    'o3_target':o3_target,
    'o3_value':o3_value,

    'rank_value':rank_value,
    'compare_city':compare_city,
    'rank_target':rank_target,
    'compare_rank':compare_rank,
    'compare_composite_index':compare_composite_index,

    'fine_days':fine_days,
    'fine_target':fine_target,
    'fine_value':fine_value,

    'compare_target':compare_target,
    'compare_day':compare_day,
#####第二部分####
    'pm25_target':pm25_target,
    'total_pm25_value':total_pm25_value,
    'avg_pm25_value':avg_pm25_value,
    'differ_pm25value':differ_pm25value,
    'later_year':later_year,  
    'season_value':season[report_month],
    'avg_pm25_value':avg_pm25_value,
    'avg_pm25_target':avg_pm25_target,
    'pm25_target_value':pm25_target_value,
    'heavy_pollute_day':heavy_pollute_day,
    'heavy_pollute_target':heavy_pollute_target,
    'heavy_pollute_day_last':heavy_pollute_day_last,
    #####第三部分####
    'image_2': InlineImage(tpl, img_2, width=Mm(150)),
    'aqi_value':aqi_value,
    'quality_value':quality_value,
    'primary_poll':primary_poll,
    'no2_day_value':no2_day_value,
    'pm10_day_value':pm10_day_value,
    'pm25_day_value':pm25_day_value,
    'poll_type':poll_type,

    'lower_time':lower_time,
    'lower_clock':lower_clock,
    'lower_value':lower_value,
    'upper_time':upper_time,
    'upper_clock':upper_clock,
    'upper_value':upper_value
}
print(context)
tpl.render(context)
tpl.save('生成文档-1226.docx')
