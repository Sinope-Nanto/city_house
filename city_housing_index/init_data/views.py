from django.shortcuts import render

from rest_framework.views import APIView
from utils.api_response import APIResponse

from init_data.domains import init_database
from init_data.domains import init_city_index
from init_data.domains import init_total_data
from init_data.domains import init_city_list
from init_data.domains import init_area_index
from init_data.domains import init_base_price_06


# Create your views here.

class InitSystemViews(APIView):

    def post(self, request):
        year = int(request.data['year'])
        month = int(request.data['month'])


        return APIResponse.create_success()
