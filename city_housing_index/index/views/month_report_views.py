from rest_framework.views import APIView
from utils.api_response import APIResponse

from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission

from index.domains.generate_report import get_word_report_40
from index.domains.generate_report import get_report_90
from index.domains.generate_report import get_report_40
from index.domains.generate_report import get_word_report_90
from index.domains.generate_report import get_word_picture_40
from index.domains.generate_report import get_word_picture_90
from index.domains.plot import plot


class GenReportViews(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def post(self, request):
        year = int(request.data['year'])
        month = int(request.data['month'])
        plot(year, month)
        if (
                get_report_40(year, month)
                and get_report_90(year, month)
                and get_word_report_40(year, month)
                and get_word_report_90(year, month)
                and get_word_picture_40(year, month)
                and get_word_picture_90(year, month)
        ):
            return APIResponse.create_success()
        else:
            return APIResponse.create_fail(code=404, msg="some sourence does not exist")
