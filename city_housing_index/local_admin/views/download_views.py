from wsgiref.util import FileWrapper
from django.http import HttpResponse
from calculate.models import DataFile
from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission
import io
import zipfile
import datetime
from rest_framework.views import APIView
from utils.api_response import APIResponse
from local_admin.views.files_views import getfilelist

def getIP():
    f = open('setting.txt')
    re = f.readline()
    f.close()
    return re

# 将用户请求的文件打包为zip并提供下载
class DownLoadViews(APIView):
    # authentication_classes = [CityIndexAuthentication]
    # permission_classes = [CityIndexAdminPermission]
    def get(self,request,filenames):
        fileidlist = filenames.split('+')
        filelist = []
        for fileid in fileidlist:
            try:
                newfile = DataFile.objects.get(id = fileid)
            except DataFile.DoesNotExist:
                return APIResponse.create_fail(code=404,msg='file not exist')
            filelist.append(str(DataFile.objects.get(id = fileid).file))
        s = io.BytesIO()
        zip = zipfile.ZipFile(s, 'w')
        for i in filelist:
            f_name = 'media/' + i
            zip.write(f_name,i.split('/')[1])
        zip.close()
        s.seek(0)
        wrapper = FileWrapper(s)
        response = HttpResponse(wrapper, content_type='application/zip')
        response['Content-Disposition'] = 'attachment; filename={}.zip'.format(datetime.datetime.now().strftime("%Y-%m-%d"))
        return response  

# 返回用户要求的数据的下载地址 by id
class DownLoadbyIDViews(APIView):
    # authentication_classes = [CityIndexAuthentication]
    # permission_classes = [CityIndexAdminPermission]
    def get(self,request,fileid):
        return APIResponse.create_success(data = 'http://' + getIP()+ '/v1/admin/download_files/' + str(fileid))

# 返回用户要求的数据的下载地址 by time
class DownLoadbyTime(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]
    def post(self,request):
        year = request.data['year']
        month = request.data['month']
        try:
            year_int = int(year)
            month_int = int(month)
        except ValueError:
            return APIResponse.create_fail(code=400,msg='Bad Request')
        url = 'http://' + getIP()+ '/v1/admin/download_files/'
        filelist = []
        all_file = DataFile.objects.all()
        for i in all_file:
            data = i.start
            dataf = data.split('-')
            if len(dataf) < 2:
                continue
            if int(dataf[0]) == year_int and int(dataf[1]) == month_int:
                filelist.append(i.id)
        first = True
        for i in filelist:
            if not first:
                url += '+'
            url += str(i)
            first = False
        return APIResponse.create_success(data=url)

# 返回用户要求的数据的下载地址 by city
class DownLoadbyCity(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]
    def post(self,request):
        city = request.data['city']
        try:
            city_int = int(city)
        except ValueError:
            return APIResponse.create_fail(code=400,msg='Bad Request')
        url = 'http://' + getIP()+ '/v1/admin/download_files/'
        all_file = DataFile.objects.filter(city_id=city_int)
        first = True
        for i in all_file:
            if not first:
                url += '+'
            url += str(i.id)
            first = False
        return APIResponse.create_success(data=url)

# 返回用户要求的数据的下载地址 by city and time
class DownLoadbyCityTime(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]
    def post(self,request):
        city = request.data['city']
        year = request.data['year']
        month = request.data['month']
        try:
            city_int = int(city)
            year_int = int(year)
            month_int = int(month)
        except ValueError:
            return APIResponse.create_fail(code=400,msg='Bad Request')
        url = 'http://' + getIP()+ '/v1/admin/download_files/'
        all_file = DataFile.objects.filter(city_id=city_int)
        filelist = []
        for i in all_file:
            data = i.start
            dataf = data.split('-')
            if len(dataf) < 2:
                continue
            if int(dataf[0]) == year_int and int(dataf[1]) == month_int:
                filelist.append(i.id)
        first = True
        for i in filelist:
            if not first:
                url += '+'
            url += str(i)
            first = False
        return APIResponse.create_success(data=url)
        
