from rest_framework.views import APIView
from utils.api_response import APIResponse

from index.models import CalculateResult, CityIndexCalculateTaskRecord
from city.models import City

from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission

from index.domains.addinfo import add_new_month, upload_city_info_to_database

from index.domains.indexcul import calculate_city_index
from index.tasks import city_calculate


def id_to_code(city_id):
    city = City.objects.get(id=city_id)
    return int(city.code)


class AddNewMonthColumnView(APIView):
    authentication_classes = [CityIndexAuthentication]
    permission_classes = [CityIndexAdminPermission]

    def post(self, request):
        try:
            add_new_month(int(request.data['year']), int(request.data['month']))
            return APIResponse.create_success()
        except ValueError:
            return APIResponse.create_fail(code=400, msg='bad request')


class UpdateAllCityIndexView(APIView):
    # authentication_classes = [CityIndexAuthentication]
    # permission_classes = [CityIndexAdminPermission]

    def post(self, request):
        running_task = None
        if running_task:
            result = {
                "task_id": ""
            }
        else:
            task_record = CityIndexCalculateTaskRecord(
                kwargs={"year": int(request.data['year']), "month": int(request.data['month'])})
            task_record.code = task_record.generate_code()
            task_record.save()

            print("start delay")
            city_calculate.delay(year=int(request.data['year']), month=int(request.data['month']),
                                 task_id=task_record.id)

            print("delay over")
            result = {
                "task_id": task_record.id
            }
        return APIResponse.create_success(result)
        # city_list = City.objects.filter(ifin90=True)
        # uploaded_city_list = []
        # for city in city_list:
        #     if upload_city_info_to_database(int(request.data['year']), int(request.data['month']), city.code):
        #         pass
        #     else:
        #         uploaded_city_list.append(city.name)
        # if len(uploaded_city_list) == 0:
        #     return APIResponse.create_success(data='所有城市均已上传')
        # else:
        #     return APIResponse.create_success(data='未上传城市有:' + str(uploaded_city_list))


class CalculateCityInfoView(APIView):
    # authentication_classes = [CityIndexAuthentication]

    def post(self, request):
        city_code = int(request.data['code'])
        if upload_city_info_to_database(int(request.data['year']), int(request.data['month']), city_code):
            data = calculate_city_index(city_code, int(request.data['year']), int(request.data['month']))
            return APIResponse.create_success(data=data)
        else:
            return APIResponse.create_fail(code=400, msg='bad request')


class GetCityIndexInfoView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def post(self, request):
        from index.serializers import CalculateResultSimpleSerializer
        city_code = int(request.data['code'])
        result = CalculateResult.objects.get(year=int(request.data['year']), month=int(request.data['month']),
                                             city_or_area=True, city=city_code)

        if result is None:
            return APIResponse.create_fail(code=400, msg='当月城市数据未计算')
        else:
            result_data = {'index': result.index_value, 'chain': result.chain_index,
                           'year_on_year': result.year_on_year_index,
                           'volumn': result.trade_volume}
            result_data.update(CalculateResultSimpleSerializer(result).data)
            return APIResponse.create_success(
                data=result_data)


class ListCityIndexInfoView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def get(self, request):
        from index.serializers import CalculateResultSimpleSerializer
        city_code = request.GET['code']
        calculate_results = CalculateResult.objects.filter(city=city_code, city_or_area=True).order_by('year', 'month')
        result = []
        for calculate_result in calculate_results:
            item_data = {
                'index': calculate_result.index_value,
                'chain': calculate_result.chain_index,
                'year_on_year': calculate_result.year_on_year_index,
                'volumn': calculate_result.trade_volume,
                'year': calculate_result.year,
                'month': calculate_result.month
            }
            item_data.update(CalculateResultSimpleSerializer(calculate_result).data)
            result.append(item_data)
        return APIResponse.create_success(result)


class CalculateTaskView(APIView):
    # authentication_classes = [CityIndexAuthentication]
    # permission_classes = [CityIndexAdminPermission]

    def get(self, request):
        task_id = request.GET['task_id']
        task_record = CityIndexCalculateTaskRecord.objects.get(id=task_id)
        data = {
            "id": task_id,
            "code": task_record.code,
            "progress": task_record.progress,
            "current_task": task_record.current_task,
            "finished": task_record.finished,
            "task_id": task_id
        }
        return APIResponse.create_success(data)
