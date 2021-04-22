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

class DownloadReoprt(APIView):
    # authentication_classes = [CityIndexAuthentication]
    # permission_classes = [CityIndexAdminPermission]

    def post(self,request):
        year = int(request.data['year'])
        month = int(request.data['month'])
        filename = 'media/report/'+'40_city_report_'+str(year)+'_'+str(month)+'.xlsx'
        try:
            report = open(filename,'rb')
        except FileNotFoundError:
            return APIResponse.create_fail(code=404,msg='report not found')
        response = HttpResponse(report)
        response['Content-Type'] = 'application/octet-stream'
        response['Content-Disposition'] = 'attachment;filename={year}年{month}月40城市指数汇总及图表.xlsx'.format(year = str(year), month=str(month))
        report.close()
        return response
