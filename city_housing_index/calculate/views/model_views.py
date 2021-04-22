from rest_framework.views import APIView
from local_auth.authentication import CityIndexAuthentication
from utils.api_response import APIResponse
from calculate.domains.model_calculate_domain import execute_model_calculate, list_model_calculate_tasks, \
    list_model_calculate_result, get_model_calculate_result_detail


class ExecuteModelCalculateView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def post(self, request):
        file_id = request.data['file_id']
        user = request.user

        success, ret = execute_model_calculate(user, file_id)

        if not success:
            return APIResponse.create_fail(403, "您没有使用此文件的权限")

        return APIResponse.create_success(data=ret)


class ListModelCalculateTaskView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def get(self, request):
        user = request.user
        ret = list_model_calculate_tasks(user)
        return APIResponse.create_success(data=ret)


class ListModelCalculateResultView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def get(self, request):
        user = request.user
        ret = list_model_calculate_result(user)
        return APIResponse.create_success(data=ret)


class GetModelCalculateResultDetailView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def get(self, request, result_id):
        user = request.user
        success, ret = get_model_calculate_result_detail(user, result_id)

        if not success:
            return APIResponse.create_fail(400, "计算结果不存在，或您没有权限查看")
        return APIResponse.create_success(data=ret)
