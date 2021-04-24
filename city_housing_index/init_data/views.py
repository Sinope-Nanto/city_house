from django.shortcuts import render

from rest_framework.views import APIView
from utils.api_response import APIResponse

from init_data.admin import UpdataDatabase
from init_data.admin import UpdataCityIndex
from init_data.admin import UpdataTotalData
from init_data.admin import UpdataCityList
from init_data.admin import UpdataAreaIndex
# Create your views here.

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



            

        


