from rest_framework.views import APIView
from utils.api_response import APIResponse

from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission

from index.admin.getreport import getReport

class ExcelReportViews(APIView):

    def post(self,request):
        year = int(request.data['year'])
        month = int(request.data['month'])
        if getReport(year,month):
            return APIResponse.create_success()
        else:
            return APIResponse.create_fail(code=404,msg="some sourence does not exist")
