from django.contrib import admin

# Register your models here.

from city.models import City
from django.core.exceptions import ObjectDoesNotExist
from index.models import CalculateResult
from city.enums import CityArea
from index.admin.addinfo import AddNewMonth

def UpdataDatabase(year:int, month:int):
    for i in range(20006,year):
        for j in range(1,13):
            AddNewMonth(i,j)
    for i in range(1,month):
        AddNewMonth(year,i)
    return True

def UpdataCityIndex():
    data_file = open('media/init_data/cityindexdata.csv',encoding='utf-8')
    volumn_file = open('media/init_data/cityvolumndata.csv',encoding='utf-8')
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
            datarow.index_value_base09 = float(data_list[i])/float(data_list[38])
            datarow.save()
    return True

def UpdataTotalData():
    data_file = open('media/init_data/totaldata.csv',encoding='utf-8')
    data = []
    for i in range(0,18):
        data.append(data_file.readline().split(','))
    for i in range(1,len(data[0])):
        year = (i-1)//12 + 2006
        month = (i-1)%12 + 1 
        try:
            dataline = CalculateResult.objects.get(year=year,month=month,city_or_area=False,area=CityArea.QUNGUO)
        except ObjectDoesNotExist:
            return APIResponse.create_fail(code=404,msg='对应月份未被上载')
        dataline.index_value = float(data[6][i])
        dataline.index_value_under90 = float(data[9][i])
        dataline.index_value_above144 = float(data[11][i])
        dataline.index_value_90144 = float(data[10][i])
        dataline.trade_volume = int(data[0][i])
        dataline.trade_volume_under_90 = int(data[1][i])
        dataline.trade_volume_above_144 = int(data[3][i])
        dataline.trade_volume_90_144 = int(data[2][i])
        dataline.index_value_base09 = float(data[18][i])
        dataline.index_value_under90_base09 = float(data[19][i])
        dataline.index_value_90144_base09 = float(data[20][i])
        dataline.index_value_above144_base09 = float(data[21][i])
        
        if i < 2:
            pass
        else:
            dataline.volume_chain = float(data[5][i])
            dataline.chain_index = float(data[8][i])
            dataline.chain_index_above144 = float(data[17][i])
            dataline.chain_index_under90 = float(data[13][i])
            dataline.chain_index_90144 = float(data[15][i])
        
        if i < 13:
            pass
        else:
            dataline.volume_year_on_year = float(data[4][i])
            dataline.year_on_year_index = float(data[7][i])
            dataline.year_on_year_index_above144 = float(data[16][i])
            dataline.year_on_year_index_under90 = float(data[12][i])
            dataline.year_on_year_index_90144 = float(data[14][i])
        dataline.save()
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
    for i in range(0,13):
        data.append(data_file.readline().split(','))
    for i in range(1,len(data[0])):
        year = (i-1)//12 + 2006
        month = (i-1)%12 + 1
        east = CalculateResult.objects.get(city_or_area=False,area=CityArea.DONGBU,year=year,month=month)
        west = CalculateResult.objects.get(city_or_area=False,area=CityArea.XIBU,year=year,month=month)
        mid = CalculateResult.objects.get(city_or_area=False,area=CityArea.ZHONGBU,year=year,month=month)
        csj = CalculateResult.objects.get(city_or_area=False,area=CityArea.CHANGSANJIAO,year=year,month=month)
        csj.index_value = float(data[12][i])
        east.trade_volume = int(data[0][i])
        east.index_value = float(data[3][i])
        west.trade_volume = int(data[2][i])
        west.index_value = float(data[5][i])
        mid.index_value = float(data[4][i])
        mid.trade_volume = int(data[1][i])
        if i < 13:
            pass
        else:
            east.year_on_year_index = float(data[6][i])
            
            west.year_on_year_index = float(data[10][i])
            mid.year_on_year_index = float(data[8][i])
        if i < 2:
            pass
        else:
            east.chain_index = float(data[7][i])
            west.chain_index = float(data[11][i])
            mid.chain_index = float(data[9][i])
        csj.save()
        east.save()
        west.save()
        mid.save()
    return True   
