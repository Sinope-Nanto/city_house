from django.http import HttpResponse
from calculate.models import DataFile
from city.models import City
import datetime
from local_auth.models import UserProfile
from rest_framework.views import APIView
from utils.api_response import APIResponse
from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission

def GetUploadStatus():
    text = City.objects.all()
    data_list = []
    for i in text:
        upload = DataFile.objects.filter(city_id = i.code)
        exist = False
        for j in upload:
            data = j.start.split('-')
            if len(data) < 2:
                continue
            if int(data[0]) == datetime.datetime.now().year and int(data[1]) == datetime.datetime.now().month:
                exist = True
                break
        data_list.append({'id':i.code,'name':i.name,'upload_status':exist})
    return data_list
class UploadStatusViews(APIView):
    # authentication_classes = [CityIndexAuthentication]
    # permission_classes = [CityIndexAdminPermission]
    def get(self,request):
        data = []
        data = GetUploadStatus()
        return APIResponse.create_success(data=data)
