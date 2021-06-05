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

    def post(self,request):
        year = request.data['year']
        month = request.data['month']
        year_int = 0
        month_int = 0 
        try:
            year_int = int(year)
            month_int = int(month)
        except ValueError:
            return APIResponse.create_fail(code=400,msg='Bad Request')
        city_list = City.objects.all()
        responsedata = getfilelist(year_int,month_int)
        return APIResponse.create_success(data=responsedata)

def getfilelist(year:int , month:int):
    city_list = City.objects.all()
    responsedata = []
    for i in city_list:
        file_id = 0
        file_list = DataFile.objects.filter(city_id = i.code)
        for f in file_list:
            data = f.start.split('-')
            if len(data) < 2:
                continue
            if int(data[0]) == year and int(data[1]) == month:
                file_id = f.id
                responsedata.append({'city_id':i.code,'city_name':i.name,'file_id':int(file_id)})
    return responsedata
