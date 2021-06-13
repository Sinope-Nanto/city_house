from rest_framework.views import APIView
from rest_framework.viewsets import ViewSet

from .domain import *

from utils.api_response import APIResponse


class SendSMSCodeView(APIView):

    def post(self, request):
        post_data = request.data
        mobile = post_data["mobile"]

        if check_login_mobile(mobile):
            send_sms_code(mobile)
        else:
            return APIResponse.create_fail(403, "invalid mobile")
        return APIResponse.create_success()


class LoginView(APIView):

    def post(self, request):
        post_data = request.data
        mobile = post_data["mobile"]
        verify_code = post_data["verify"]

        check_status, msg = check_auth(mobile, verify_code)
        if not check_status:
            return APIResponse.create_fail(403, msg)
        from city.serializers import CitySerializer
        token, user_profile = create_token(mobile)
        return APIResponse.create_success(data={
            "token": token,
            "user_id": user_profile.user_id.id,
            "name": user_profile.name,
            "admin": user_profile.is_admin(),
            "city": CitySerializer(user_profile.city).data
        })


class RegisterView(ViewSet):
    queryset = UserProfile.objects.all()

    def create(self, request):
        mobile = request.data['mobile']
        name = request.data['name']
        city = int(request.data['city'])
        identity = request.data['identity']
        identity_image = request.FILES.get('identity_image')

        if check_login_mobile(mobile):
            return APIResponse.create_fail(400, "该手机号已被注册")

        created, msg = register(mobile=mobile, name=name, city=city, identity=identity, identity_image=identity_image)

        if not created:
            return APIResponse.create_fail(400, msg)
        return APIResponse.create_success()
