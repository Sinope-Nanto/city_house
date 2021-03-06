from rest_framework.views import APIView
from utils.api_response import APIResponse

from index.models import CityIndex
from city.models import City

from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission

from index.admin.addinfo import AddNewMonth
from index.admin.addinfo import UploadCityInfoToDatabase


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
            AddNewMonth(int(request.data['year']), int(request.data['month']))
            return APIResponse.create_success()
        except ValueError:
            return APIResponse.create_fail(code=400, msg='bad request')


class UpdataCityInfoView(APIView):
    def post(self, request):
        city_list = City.objects.filter(ifin90=True)
        unloadcitylist = []
        for city in city_list:
            if UploadCityInfoToDatabase(int(request.data['year']), int(request.data['month']), city.id):
                pass
            else:
                unloadcitylist.append(city.name)
        if len(unloadcitylist) == 0:
            return APIResponse.create_success(data='所有城市均已上传')
        else:
            return APIResponse.create_success(data='未上传城市有:' + str(unloadcitylist))
