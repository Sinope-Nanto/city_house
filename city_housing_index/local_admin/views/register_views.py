from rest_framework.views import APIView
from local_admin.domains.register_domain import *
from local_admin.permissions import CityIndexAdminPermission

from utils.api_response import APIResponse
from local_auth.authentication import CityIndexAuthentication
from django.contrib.auth.models import User


class GetWaitingRegisterUsersView(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def get(self, request):
        response_data = get_waiting_register_list()

        return APIResponse.create_success(data=response_data)


class GetWaitingRegisterUserDetailView(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def get(self, request, user_id=None):
        response_data = get_waiting_register_detail(user_id)

        return APIResponse.create_success(data=response_data)


class AcceptRegisterUserView(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def post(self, request, user_id=None):
        accepted, msg = accept_register_user(user_id)
        if not accepted:
            return APIResponse.create_fail(400, msg)
        return APIResponse.create_success()


class RefuseRegisterUserView(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def post(self, request, user_id=None):
        accepted, msg = refuse_register_user(user_id)

        if not accepted:
            return APIResponse.create_fail(400, msg)
        return APIResponse.create_success()
