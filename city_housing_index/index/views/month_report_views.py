from rest_framework.views import APIView
from utils.api_response import APIResponse

from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission

from index.admin.getreport import getWordReport
from index.admin.getreport import getReport90
from index.admin.getreport import getReport
from index.admin.getreport import getWordReport90
from index.admin.getreport import getWordPicture
from index.admin.getreport import getWordPicture90
from index.views.plot import plot


class GenReportViews(APIView):

    def post(self, request):
        year = int(request.data['year'])
        month = int(request.data['month'])
        plot(year, month)
        if (
                getReport(year, month)
                and getReport90(year, month)
                and getWordReport(year, month)
                and getWordReport90(year, month)
                and getWordPicture(year, month)
                and getWordPicture90(year, month)
        ):
            return APIResponse.create_success()
        else:
            return APIResponse.create_fail(code=404, msg="some sourence does not exist")
