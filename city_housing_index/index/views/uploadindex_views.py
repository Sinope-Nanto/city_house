from rest_framework.views import APIView
from utils.api_response import APIResponse

from index.models import CityIndex
from index.models import CalculateResult
from city.models import City

from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission

from index.admin.addinfo import add_new_month
from index.admin.addinfo import upload_city_info_to_database

from index.admin.indexcul import CalculateCityIndex


class UpLoadIndexView(APIView):
    # authentication_classes = [CityIndexAuthentication]
    # permission_classes = [CityIndexAdminPermission]

    def post(self, request):
        try:
            year = int(request.data['year'])
            month = int(request.data['month'])
            city_id = int(request.data['city_id'])
            index = float(request.data['index'])
        except ValueError:
            return APIResponse.create_fail(code=400, msg='bad Request')
        newdata = CityIndex(year=year, month=month, city=city_id, value=index)
        newdata.save()
        return APIResponse.create_success()


class AddNewMonthLine(APIView):
    # authentication_classes = [CityIndexAuthentication]
    # permission_classes = [CityIndexAdminPermission]

    def post(self, request):
        try:
            add_new_month(int(request.data['year']), int(request.data['month']))
            return APIResponse.create_success()
        except ValueError:
            return APIResponse.create_fail(code=400, msg='bad request')


class UpdataCityInfoView(APIView):
    def post(self, request):
        city_list = City.objects.filter(ifin90=True)
        unloadcitylist = []
        for city in city_list:
            if upload_city_info_to_database(int(request.data['year']), int(request.data['month']), city.id):
                pass
            else:
                unloadcitylist.append(city.name)
        if len(unloadcitylist) == 0:
            return APIResponse.create_success(data='所有城市均已上传')
        else:
            return APIResponse.create_success(data='未上传城市有:' + str(unloadcitylist))

class CalculateCityInfoView(APIView):

    def post(self, request):
        if upload_city_info_to_database(int(request.data['year']), int(request.data['month']), int(request.data['code'])):
            date = CalculateCityIndex(int(request.data['code']), int(request.data['year']), int(request.data['month']))
            return APIResponse.create_success(data = date)
        else:
            return APIResponse.create_fail(code=400, msg='bad request')

class FindCityInfo(APIView):

    def post(self, request):
        city = CalculateResult.objects.get(year = int(request.data['year']), month = int(request.data['month']), city_or_area=True, city=int(request.data['code']))
        if city is None:
            return APIResponse.create_fail(code=404, msg='当月城市数据未计算')
        else:
            return APIResponse.create_success(data = {'index': city.index_value, 'chain': city.chain_index, 'year_on_year':city.year_on_year_index, 'volumn': city.trade_volume})
            

