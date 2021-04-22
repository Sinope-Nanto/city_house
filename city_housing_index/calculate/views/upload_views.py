from rest_framework.viewsets import ViewSet
from rest_framework.views import APIView
from rest_framework.response import Response
from calculate.serializers import DataFileSerializer
from calculate.models import DataFile
from local_auth.authentication import CityIndexAuthentication
from utils.api_response import APIResponse
from calculate.domains.upload_domain import create_upload_file, list_upload_files, check_file_permission, \
    review_file_content, delete_data_file


class DataFileUploadView(ViewSet):
    # authentication_classes = [CityIndexAuthentication]
    serializer_class = DataFileSerializer
    queryset = DataFile.objects.all()

    def list(self, request):
        return Response("GET API")

    def create(self, request):
        user = request.user
        file = request.FILES.get('file')
        name = request.data["name"]
        code = request.data["code"]
        city_id = request.data["city_id"]
        start = request.data["start"]
        end = request.data["end"]
        print(type(file))
        print(file.content_type)

        ret, msg = create_upload_file(user, file, name, code, city_id, start, end)
        if not ret:
            return APIResponse.create_fail(code=400, msg=msg)
        return APIResponse.create_success()


class DataFileListView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def get(self, request):
        user = request.user
        data_file_list = list_upload_files(user)
        return APIResponse.create_success(data=data_file_list)


class ReviewFileContentView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def get(self, request, file_id=None):
        user = request.user
        if not check_file_permission(user, file_id):
            return APIResponse.create_fail(403, "您没有预览该文件的权限")

        ret_data = review_file_content(file_id)
        return APIResponse.create_success(data=ret_data)


class DeleteDataFileView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def post(self, request, file_id):
        user = request.user

        if not check_file_permission(user, file_id):
            return APIResponse.create_fail(403, "您没有预览该文件的权限")
        delete_data_file(file_id)
        return APIResponse.create_success()
