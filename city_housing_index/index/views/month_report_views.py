from rest_framework.views import APIView
from utils.api_response import APIResponse

from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission

from index.tasks import generate_report
from ..models import GenReportTaskRecord


class GenReportViews(APIView):
    # authentication_classes = [CityIndexAuthentication]
    # permission_classes = [CityIndexAdminPermission]

    def post(self, request):
        year = int(request.data['year'])
        month = int(request.data['month'])
        user = request.user
        running_task = None
        if running_task:
            result = {
                "task_id": ""
            }
        else:
            task_record = GenReportTaskRecord(kwargs={"year": year, "month": month})
            task_record.code = task_record.generate_code()
            task_record.save()

            print("start delay")
            generate_report.delay(year=year, month=month, task_id=task_record.id)

            print("delay over")
            result = {
                "task_id": task_record.id
            }
        return APIResponse.create_success(result)


class QueryReportTaskView(APIView):
    # authentication_classes = [CityIndexAuthentication]
    # permission_classes = [CityIndexAdminPermission]

    def get(self, request):
        from ..serializers import GenReportTaskSerializer
        print(request.data)
        task_id = request.GET['task_id']
        task_record = GenReportTaskRecord.objects.get(id=task_id)
        result = GenReportTaskSerializer(task_record).data
        return APIResponse.create_success(result)
