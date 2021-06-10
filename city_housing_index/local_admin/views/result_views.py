from rest_framework.views import APIView

from local_admin.domains.result_domain import list_all_city_calculate_result
from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission

from utils.api_response import APIResponse


class ListCityCalculateResultView(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def get(self, request):
        year = request.data['year']
        month = request.data['month']

        result = list_all_city_calculate_result(year, month)
        return APIResponse.create_success(result)
