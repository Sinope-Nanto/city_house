from calculate.models import DataFile
from django.http import HttpResponse
from city.models import City
# from django.views.decorators.csrf import csrf_exempt
from rest_framework.views import APIView
from utils.api_response import APIResponse
from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission


class UpLoadFilesViews(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def post(self, request):
        year = request.data['year']
        month = request.data['month']
        try:
            year_int = int(year)
            month_int = int(month)
        except ValueError:
            return APIResponse.create_fail(code=400, msg='Bad Request')
        response_data = get_file_list(year_int, month_int)
        return APIResponse.create_success(data=response_data)


def get_file_list(year: int, month: int):
    city_list = City.objects.all()
    response_data = []
    for city in city_list:
        file_id = 0
        file_list = DataFile.objects.filter(city_id=city.id)
        for f in file_list:
            data = f.start.split('-')
            if len(data) < 2:
                continue
            if int(data[0]) == year and int(data[1]) == month:
                file_id = f.id
                break
        response_data.append({'city_id': city.id, 'city_name': city.name, 'file_id': int(file_id)})
    return response_data
