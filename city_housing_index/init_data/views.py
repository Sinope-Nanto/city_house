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

        if UpdataAreaIndex():
            return APIResponse.create_success()
        else:
            return APIResponse.create_fail(code=500,msg='unkowned error')

class InitCityIndexViews(APIView):

    def get(self,request):
        if UpdataCityIndex():
            return APIResponse.create_success()
        else:
            return APIResponse.create_fail(code=500,msg='Unknowed error')

class InitDatabaseViews(APIView):

    def post(self,request):
        year = int(request.data['year'])
        month = int(request.data['month'])
        if UpdataDatabase(year,month):
            return APIResponse.create_success()
        else:
            return APIResponse.create_fail(code=500,msg='Unknowed error')


class InitSystemViews(APIView):

    def post(self,request):
        year = int(request.data['year'])
        month = int(request.data['month'])
        UpdataCityList()
        UpdataDatabase(year,month)
        UpdataTotalData()
        UpdataCityIndex()
        UpdataAreaIndex()

        return APIResponse.create_success()



            

        


