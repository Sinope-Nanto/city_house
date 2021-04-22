from calculate.tasks import ModelCalculateTask
from calculate.models import CalculateTask, DataFile, ModelCalculateResult
from calculate.enums import CalculateTaskType, CalculateTaskStatus
from calculate.serializers import ModelCalculateResultSerializer, CalculateTaskSerializer
import json


def check_file_permission(user, file_id):
    data_file = DataFile.objects.get(id=file_id)
    return data_file.user_id == user.id


def execute_model_calculate(user, file_id):
    if not check_file_permission(user, file_id):
        return False, {}

    task_code = CalculateTask.generate_code()
    task_record = CalculateTask(user=user, code=task_code, type=CalculateTaskType.MODEL_CALCULATE,
                                kwargs={"file_id": file_id}, status=CalculateTaskStatus.START)
    task_record.save()

    ModelCalculateTask().delay(file_id=file_id, task_id=task_record.id)
    return True, {"task_id": task_record.id, "task_code": task_code}


def list_model_calculate_tasks(user):
    tasks_queryset = CalculateTask.objects.filter(user=user, type=CalculateTaskType.MODEL_CALCULATE)
    return CalculateTaskSerializer(tasks_queryset, many=True).data


def list_model_calculate_result(user):
    task_id_list = list(
        CalculateTask.objects.filter(user=user, type=CalculateTaskType.MODEL_CALCULATE).values_list('id', flat=True))
    result_queryset = ModelCalculateResult.objects.filter(task_id__in=task_id_list)
    return ModelCalculateResultSerializer(result_queryset, many=True).data


def get_model_calculate_result_detail(user, result_id):
    try:
        result_obj = ModelCalculateResult.objects.get(id=result_id)
    except ModelCalculateResult.DoesNotExist:
        return False, {}

    if result_obj.task.user_id != user.id:
        return False, {}

    ret_data = ModelCalculateResultSerializer(result_obj).data

    ret_data['details'] = result_obj.details
    return True, ret_data
