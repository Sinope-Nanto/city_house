from django.contrib import admin

# Register your models here.

from city.models import City
from django.core.exceptions import ObjectDoesNotExist
from index.models import CalculateResult
from city.enums import CityArea
from index.admin.addinfo import AddNewMonth

def UpdataDatabase(year:int, month:int):
    for i in range(2006,year):
        for j in range(1,13):
            AddNewMonth(i,j)
    for i in range(1,month):
        AddNewMonth(year,i)
    return True

def UpdataCityIndex():
    data_file = open('media/init_data/cityindexdata.csv')
    volumn_file = open('media/init_data/cityvolumndata.csv')
    code = 1
    while True:
        data = data_file.readline()
        vol = volumn_file.readline()
        if data == "":
            break
        data_list = data.split(',')
        vol_list = vol.split(',')
        city = code
        code += 1
        for i in range(2,len(data_list)):
            year = (i-2)//12 + 2006
            month = (i-2)%12 + 1
            datarow = CalculateResult.objects.get(city_or_area=True,city=city,year=year,month=month)
            datarow.index_value = float(data_list[i])
            datarow.trade_volume = int(vol_list[i])
            if i < 3:
                pass
            else:
                try:
                    datarow.chain_index = float(data_list[i])/float(data_list[i-1]) - 1
                except ZeroDivisionError:
                    datarow.chain_index = 0
            if i < 14:
                pass
            else:
                try:
                    datarow.year_on_year_index = float(data_list[i])/float(data_list[i-12]) - 1
                except ZeroDivisionError:
                    datarow.year_on_year_index = 0
            datarow.index_value_base09 = float(data_list[i])/float(data_list[38])*100
            datarow.save()
    return True

def UpdataTotalData():
    data_file = open('media/init_data/totaldata.csv')
    data = []
    data_90 = []
    for i in range(0,18):
        data.append(data_file.readline().split(','))
    for i in range(18,36):
        data_90.append(data_file.readline().split(','))
    for i in range(1,len(data[0])):
        year = (i-1)//12 + 2006
        month = (i-1)%12 + 1 
        try:
            dataline = CalculateResult.objects.get(year=year,month=month,city_or_area=False,area=CityArea.QUNGUO)
            dataline_90 = CalculateResult.objects.get(year=year,month=month,city_or_area=False,area=CityArea.QUANGUO_90)
        except ObjectDoesNotExist:
            return False
        
        var_order = [
            dataline.trade_volume, dataline.trade_volume_under_90, dataline.trade_volume_90_144, dataline.trade_volume_above_144,
            dataline.index_value, dataline.index_value_under90, dataline.index_value_90144, dataline.index_value_above144,
            dataline.volume_year_on_year, dataline.volume_chain, dataline.year_on_year_index, dataline.chain_index,
            dataline.year_on_year_index_under90, dataline.chain_index_under90,
            dataline.year_on_year_index_90144, dataline.chain_index_above144,
            dataline.year_on_year_index_above144, dataline.chain_index_above144            
        ]
        var_order_90 = [
            dataline_90.trade_volume, dataline_90.trade_volume_under_90, dataline_90.trade_volume_90_144, dataline_90.trade_volume_above_144,
            dataline_90.index_value_base09, dataline_90.index_value_under90_base09, dataline_90.index_value_90144_base09, dataline_90.index_value_above144_base09,
            dataline_90.volume_year_on_year, dataline_90.volume_chain, dataline_90.year_on_year_index, dataline_90.chain_index,
            dataline_90.year_on_year_index_under90, dataline_90.chain_index_under90,
            dataline_90.year_on_year_index_90144, dataline_90.chain_index_above144,
            dataline_90.year_on_year_index_above144, dataline_90.chain_index_above144                
        ]
        for j in range(0,4):
            var_order[j] = int(data[j][i])
            var_order_90[j] = int(data_90[j][i])
        for j in range(4,18):
            var_order[j] = float(data[j][i])
            var_order_90[j] = float(data_90[j][i])
        dataline.save()
        dataline_90.save()
    return True

def UpdataCityList():
    city_file = open('media/init_data/citylist.csv',encoding = 'utf-8')
    city_file.readline()
    while True:
        
        city = city_file.readline()
        if city == "":
            break
        city = city.split(',')
        try:
            row = City.objects.get(code=int(city[0]))
            row.delete()
        except ObjectDoesNotExist:
            pass
        row = City(name=city[1],code=int(city[0]),area=int(city[2]),line=int(city[3]),
        special_area=int(city[4]),block=int(city[5]),ifin40=int(city[6]),ifin90=int(city[7]))
        row.save()
    return True 

def UpdataAreaIndex():
    data_file = open('media/init_data/areadata.csv',encoding='utf-8')
    data = []
    for i in range(0,84):
        data.append(data_file.readline().split(','))
    for i in range(1,len(data[0])):
        year = (i-1)//12 + 2006
        month = (i-1)%12 + 1
        data_row = []
        for area in range(CityArea.DONGBU,CityArea.num_of_item_90 - 1):
            data_row.append(CalculateResult.objects.get(city_or_area=False,area=area,year=year,month=month))
        for area in range(0,21):
            data_row[area].trade_volume = int(data[4*area][i])
            if area < 4:
                data_row[area].index_value = float(data[4*area + 1][i])
            else:
                data_row[area].index_value_base09 = float(data[4*area + 1][i])
            data_row[area].year_on_year_index = float(data[4*area + 2][i])
            data_row[area].chain_index = float(data[4*area + 3][i])
        for row in data_row:
            row.save()
    return True   
