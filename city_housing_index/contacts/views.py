from rest_framework.views import APIView
from .domain import *

from utils.api_response import APIResponse
from local_auth.authentication import CityIndexAuthentication


# Create your views here.
class GetContactsView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def get(self, request):
        user = request.user
        response_data = get_contact_list(user)
        return APIResponse.create_success(data=response_data)


class GetContactDetailView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def get(self, request, id=None):
        user = request.user
        try:
            response_data = get_contact_detail(user, id)
        except:
            return APIResponse.create_fail(401, "您没有查看该通讯录的权限")
        return APIResponse.create_success(data=response_data)


class CreateContactView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def post(self, request):
        user = request.user
        post_data = request.data
        name = post_data["name"]
        mobile = post_data["mobile"]
        position = post_data["position"]
        department = post_data["department"]
        leader = post_data["leader"]
        tel = post_data["tel"]
        qq = post_data["qq"]
        email = post_data["email"]
        fax = post_data["fax"]

        created, msg = create_contact(user, name, mobile, position, department, leader, tel, qq, email, fax)

        if not created:
            return APIResponse.create_fail(401, msg)
        return APIResponse.create_success()


class UpdateContactView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def post(self, request, id=None):
        user = request.user
        post_data = request.data
        name = post_data["name"]
        mobile = post_data["mobile"]
        position = post_data["position"]
        department = post_data["department"]
        leader = post_data["leader"]
        tel = post_data["tel"]
        qq = post_data["qq"]
        email = post_data["email"]
        fax = post_data["fax"]

        updated, msg = update_contact(user, id, name, mobile, position, department, leader, tel, qq, email, fax)
        if not updated:
            return APIResponse.create_fail(401, msg)
        return APIResponse.create_success()


class DeleteContactView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def post(self, request, id=None):
        user = request.user
        deleted, msg = delete_contact(user, id)
        if not deleted:
            return APIResponse.create_fail(401, msg)
        return APIResponse.create_success()
