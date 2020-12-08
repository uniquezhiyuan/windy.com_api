import requests
import json
import time
import xlrd
import xlwt
from math import sqrt, atan, pi
from numpy import mean

def points_read():
    work_book = xlrd.open_workbook('./天气预报/点位表.xlsx')
    sheet = work_book.sheet_by_name('点位表')
    name = sheet.col_values(0)[3:]
    coordinate = sheet.col_values(1)[3:]
    altitude = sheet.col_values(2)[3:]
    lat = [round(int(i[1:3]) + int(i[4:6])/60 + int(i[7:9])/3600, 4) for i in coordinate]
    lon = [round(int(i[12:14]) + int(i[15:17])/60 + int(i[18:20])/3600, 4) for i in coordinate]
    name_dic = {}
    for i in range(len(name)):
        name_dic[name[i]] = [round(lat[i], 3), round(lon[i], 3), int(altitude[i])]
    return name_dic


def forecast_point(name, lat, lon, alt, days):
    post = {
        "lat": lat,
        "lon": lon,
        "model": "gfs",
        "parameters": ["temp", "wind", "precip", "rh", "pressure", "ptype", "lclouds", "mclouds", "hclouds", "windGust"],
        "levels": ["surface"],
        "key": "dre0kCydf9Ce80uHxAYaQs2Vd2TCrBzV"
        }
    r = requests.post('https://api.windy.com/api/point-forecast/v2', json = post)
    resault = json.loads(r.text)
    tomorrow = int(time.time() + 1 * 86400)
    index = []
    for i in resault['ts']:
        if time.strftime("%Y-%m-%d", time.localtime(int(str(i)[:10]))) == time.strftime("%Y-%m-%d", time.localtime(tomorrow)):
            index.append(resault['ts'].index(i))
    
    hour = []
    temp = []
    wind_u = []
    wind_v = []
    gust = []
    precip = []
    rh = []
    pressure = []
    ptype = []
    lcloud = []
    mcloud = []
    hcloud = []
    temp_resault = 0
    wind_number_resault = 0
    wind_dir = 0
    for i in index:
        hour.append(time.strftime("%Y年%m月%d日%H时", time.localtime(int(str(resault['ts'][i])[:10]))))
        temp.append(round(resault['temp-surface'][i] - 273.15, 1))
        if resault['wind_u-surface'][i] == 0.0 or resault['wind_u-surface'][i] == 0:
            wind_u.append(0.00001)
        if resault['wind_u-surface'][i] == -0.0 or resault['wind_u-surface'][i] == -0:
            wind_u.append(-0.00001)
        else:
            wind_u.append(round(resault['wind_u-surface'][i], 1))
        if resault['wind_v-surface'][i] == 0.0 or resault['wind_v-surface'][i] == 0:
            wind_v.append(0.00001)
            if resault['wind_v-surface'][i] == -0.0 or resault['wind_v-surface'][i] == -0:
                wind_v.append(-0.00001)
        else:
            wind_v.append(round(resault['wind_v-surface'][i], 1))
        gust.append(round(resault['gust-surface'][i], 1))
        precip.append(round(resault['past3hprecip-surface'][i], 1))
        rh.append(round(resault['rh-surface'][i], 1))
        pressure.append(round(resault['pressure-surface'][i], 1))
        ptype.append(resault['ptype-surface'][i])
        lcloud.append(resault['lclouds-surface'][i])
        mcloud.append(resault['mclouds-surface'][i])
        hcloud.append(resault['hclouds-surface'][i])
    
    forecast = [hour, temp, wind_u, wind_v, gust, precip, rh, pressure, ptype, lcloud]
    temp_resault = str(round(min(forecast[1]), 1)) + '~' + str(round(max(forecast[1]), 1))
    wind_number = [sqrt(wind_u[i]**2 + wind_v[i]**2) for i in range(len(hour))]
    wind_number_resault = str(round(min(wind_number), 1)) + '~' + str(round(max(wind_number), 1))  # 风速区间
    wind_dir = str(90 - round(atan(wind_v[5]/wind_u[5])/pi*180, 0))  # 下午17时风向
    rh_resault = str(round(mean(forecast[6]),0))
    precip_resault = str(sum(forecast[5]))
    pressure_resault = str(int(mean(forecast[7])/1000*((1-alt/44300)**5.256))*10)
    gust_resault = str(max(forecast[4]))
    weather = '晴'
    if ptype == 5:
        weather == '零星小雪'
        if 0.1 <= sum(forecast[5]) < 2.5:
            weather = '小雪'
        if 2.5 <= sum(forecast[5]) < 5.0:
            weather = '中雪'
        if 5.0 <= sum(forecast[5]) < 10.0:
            weather ='大雪'
        if 10.0 <= sum(forecast[5]):
            weather = '暴雪'
    
    if ptype == 1:
        if 0 <= sum(forecast[5]) < 10:
            weather = '小雨'
        if 10 <= sum(forecast[5]) < 25:
            weather = '中雨'
        if 25 <= sum(forecast[5]) < 50:
            weather = '大雨'
        if 50 <= sum(forecast[5]):
            weather = '暴雨'
    
    if ptype == 7:
        weather = '雨夹雪'
    
    if ptype == 0:
        if 0.2 <= mean(lcloud + mcloud + hcloud) < 0.6:
            weather = '多云'
        if mean(lcloud + mcloud + hcloud) >= 0.6:
            weather = '阴'
    else:
        weather = '晴'
    
    forecast_resault = [weather, temp_resault, rh_resault, pressure_resault, wind_number_resault, wind_dir, gust_resault, precip_resault]
    return {name: forecast_resault}

POINTS = points_read()

NAME, LAT, LON, ALT = '布尼寺', POINTS['布尼寺'][0], POINTS['布尼寺'][1], POINTS['布尼寺'][2]
DAYS = 1
a = forecast_point(NAME, LAT, LON, ALT, DAYS)

a = forecast_point('狮泉河', 32.5, 80, 4300, 1)

while 1:
    try:
        for i in POINTS:
            print(forecast_point(i, POINTS[i][0], POINTS[i][1], POINTS[i][2], 1))
    except:
        pass


