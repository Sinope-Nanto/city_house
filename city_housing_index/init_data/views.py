from django.shortcuts import render
from city.models import City
from django.core.exceptions import ObjectDoesNotExist
from index.models import CalculateResult
from city.enums import CityArea
from index.admin.addinfo import AddNewMonth

from rest_framework.views import APIView
from utils.api_response import APIResponse

from init_data.admin import UpdataDatabase
from init_data.admin import UpdataCityIndex
from init_data.admin import UpdataTotalData
from init_data.admin import UpdataCityList
from init_data.admin import UpdataAreaIndex
# Create your views here.

class InitCityViews(APIView):
    
    def get(self,request):
        if UpdataCityList():
            return APIResponse.create_success()
        else:
            return APIResponse.create_fail(code=500,msg='Unknowed error')


class InitTotalViews(APIView):
    def get(self,request):
        if UpdataTotalData():
            return APIResponse.create_success()
        else:
            return APIResponse.create_fail(code=500,msg='Unknowed error')

class InitAreaViews(APIView):
    
    def get(self,request):
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
        return APIResponse.create_success()

class InitCityIndexViews(APIView):

    def get(self,request):
        if UpdataCityIndex():
            return APIResponse.create_success()
        else:
            return APIResponse.create_fail(code=500,msg='Unknowed error')

class InitDatabaseViews(APIView):

    def post(self,request):
        year = request.data['year']
        month = request.data['month']
        if UpdataDatabase(year,month):
            return APIResponse.create_success()
        else:
            return APIResponse.create_fail(code=500,msg='Unknowed error')


class InitSystemViews(APIView):

    def post(self,request):
        year = request.data['year']
        month = request.data['month']
        UpdataCityList()
        UpdataDatabase(year,month)
        UpdataTotalData()
        UpdataCityIndex()
        UpdataAreaIndex()

        return APIResponse.create_success()



            

        


