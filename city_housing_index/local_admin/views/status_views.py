from rest_framework.views import APIView

from local_admin.domains.status_domain import send_city_sms, get_city_upload_status
from utils.api_response import APIResponse
from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission


class UploadStatusViews(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def get(self, request):
        year = request.data['year']
        month = request.data['month']
        data = get_city_upload_status(year, month)
        return APIResponse.create_success(data=data)


class SendSMSViews(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def post(self, request):
        city_ids = request.data['city_ids']
        year = request.data['year']
        month = request.data['month']

        result = send_city_sms(city_ids, year, month)
        return APIResponse.create_success(result)
