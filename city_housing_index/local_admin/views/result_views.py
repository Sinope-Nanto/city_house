from rest_framework.views import APIView

from local_admin.domains.result_domain import list_all_city_calculate_result, list_warn_city_calculate_result
from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission

from utils.api_response import APIResponse


class ListCityCalculateResultView(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def get(self, request):
        year = request.GET['year']
        month = request.GET['month']

        result = list_all_city_calculate_result(year, month)
        return APIResponse.create_success(result)


class ListCityCalculateWarnResultView(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def get(self, request):
        year = request.GET['year']
        month = request.GET['month']

        result = list_warn_city_calculate_result(year, month)
        return APIResponse.create_success(result)
