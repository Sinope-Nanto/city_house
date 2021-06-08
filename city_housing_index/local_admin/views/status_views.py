from calculate.models import DataFile
from city.models import City
from rest_framework.views import APIView
from utils.api_response import APIResponse
from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission


def parse_year_month(date_str):
    year_month_list = date_str.split('-')
    if len(year_month_list) < 2:
        return -1, -1
    return int(year_month_list[0]), int(year_month_list[1])


def get_city_upload_status(year, month):
    city_list = City.objects.all()
    data_list = []
    for city in city_list:
        upload_files = DataFile.objects.filter(city_id=city.id)
        exist = False
        for upload_file in upload_files:
            upload_year, upload_month = parse_year_month(upload_file.start)
            if upload_year == year and upload_month == month:
                exist = True
                break
        data_list.append(
            {
                'city_id': city.id,
                'city_code': city.code,
                'city_name': city.name,
                'upload_status': exist
            }
        )
    return data_list


class UploadStatusViews(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def get(self, request):
        year = request.data['year']
        month = request.data['month']
        data = get_city_upload_status(year, month)
        return APIResponse.create_success(data=data)
