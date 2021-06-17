from django.shortcuts import render

from rest_framework.views import APIView
from utils.api_response import APIResponse

from init_data.domains import init_database
from init_data.domains import init_city_index
from init_data.domains import init_total_data
from init_data.domains import init_city_list
from init_data.domains import init_area_index
from init_data.domains import init_base_price_06
from init_data.domains import init_base_price_09
from init_data.domains import init_city_complex_info


# Create your views here.

class InitSystemViews(APIView):

    def post(self, request):
        year = int(request.data['year'])
        month = int(request.data['month'])
        init_city_list()
        init_database(year,month)
        init_city_index()
        init_total_data()
        init_area_index()
        init_base_price_06()
        init_base_price_09()
        init_city_complex_info()

        return APIResponse.create_success()
