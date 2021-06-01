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

    def post(self, request):
        year = int(request.data['year'])
        month = int(request.data['month'])
        filename = []
        filename.append('media/report/' + '40_city_report_' + str(year) + '_' + str(month) + '.docx')
        filename.append('media/report/' + '40_city_report_' + str(year) + '_' + str(month) + '.xlsx')
        filename.append('media/report/' + '90_city_report_' + str(year) + '_' + str(month) + '.docx')
        filename.append('media/report/' + '90_city_report_' + str(year) + '_' + str(month) + '.xlsx')
        filename.append('media/report/' + '40_city_picture_' + str(year) + '_' + str(month) + '.docx')
        filename.append('media/report/' + '90_city_picture_' + str(year) + '_' + str(month) + '.docx')
        s = io.BytesIO()
        zip = zipfile.ZipFile(s, 'w')
        for f in filename:
            try:
                zip.write(f, f.split('/')[2])
            except:
                return APIResponse.create_fail(code=404, msg='report not found')
        zip.close()
        s.seek(0)
        wrapper = FileWrapper(s)
        response = HttpResponse(wrapper, content_type='application/zip')
        response['Content-Disposition'] = 'attachment; filename={}.zip'.format(
            datetime.datetime.now().strftime("%Y-%m-%d"))
        return response
