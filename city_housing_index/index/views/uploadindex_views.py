from rest_framework.views import APIView
from utils.api_response import APIResponse

from index.models import CityIndex
from index.models import CalculateResult
from city.models import City

from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission

from index.domains.addinfo import add_new_month
from index.domains.addinfo import upload_city_info_to_database

from index.domains.indexcul import calculate_city_index

def id_to_code(city_id):
    city = City.objects.get(id=city_id)
    return int(city.code)


def code_to_id(city_code):
    city = City.objects.get(code=city_code)
    return city.id



class AddNewMonthColumnView(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def post(self, request):
        try:
            add_new_month(int(request.data['year']), int(request.data['month']))
            return APIResponse.create_success()
        except ValueError:
            return APIResponse.create_fail(code=400, msg='bad request')


class UpdateAllCityIndexView(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def post(self, request):
        city_list = City.objects.filter(ifin90=True)
        uploaded_city_list = []
        for city in city_list:
            if upload_city_info_to_database(int(request.data['year']), int(request.data['month']), city.code):
                pass
            else:
                uploaded_city_list.append(city.name)
        if len(uploaded_city_list) == 0:
            return APIResponse.create_success(data='所有城市均已上传')
        else:
            return APIResponse.create_success(data='未上传城市有:' + str(uploaded_city_list))


class CalculateCityInfoView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def post(self, request):
        city_code = id_to_code(int(request.data['code']))
        if upload_city_info_to_database(int(request.data['year']), int(request.data['month']), city_code):
            data = calculate_city_index(int(request.data['code']), int(request.data['year']),city_code)
            return APIResponse.create_success(data=data)
        else:
            return APIResponse.create_fail(code=400, msg='bad request')


class GetCityIndexInfoView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def post(self, request):
        city_code = id_to_code(int(request.data['code']))
        city = CalculateResult.objects.get(year=int(request.data['year']), month=int(request.data['month']),
                                           city_or_area=True, city=city_code)
        if city is None:
            return APIResponse.create_fail(code=400, msg='当月城市数据未计算')
        else:
            return APIResponse.create_success(
                data={'index': city.index_value, 'chain': city.chain_index, 'year_on_year': city.year_on_year_index,
                      'volumn': city.trade_volume})


class ListCityIndexInfoView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def get(self, request):
        city_code = request.GET['code']
        city_code = id_to_code(city_code)
        calculate_results = CalculateResult.objects.filter(city=city_code, city_or_area=True).order_by('year', 'month')
        result = []
        for calculate_result in calculate_results:
            result.append({
                'index': calculate_result.index_value,
                'chain': calculate_result.chain_index,
                'year_on_year': calculate_result.year_on_year_index,
                'volumn': calculate_result.trade_volume,
                'year': calculate_result.year,
                'month': calculate_result.month
            })
        return APIResponse.create_success(result)
