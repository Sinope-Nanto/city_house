from celery import Task

from calculate.task_utils import ModelCalculateUtils, PriceSequenceUtils
from calculate.models import DataFile, FileContent, CalculateTask, ModelCalculateResult, PriceSequenceCalculateResult


class ModelCalculateTask(Task):
    max_retries = 3
    time_limit = 600

    def run(self, *args, **kwargs):
        task_id = kwargs.get("task_id")
        file_id = kwargs.get("file_id")

        calculate_task = CalculateTask.objects.get(id=task_id)
        calculate_task.execute()

        try:
            input_data = FileContent.objects.get(file_id=file_id).content
            data_file = DataFile.objects.get(id=file_id)
            calculate_utils = ModelCalculateUtils(input=input_data)
            raw_result = calculate_utils.get_result()

            task_result = ModelCalculateResult(task=calculate_task, file_name=data_file.name, data_file=data_file,
                                               details=raw_result)
            task_result.save()

            calculate_task.finish()
            return True
        except:
            calculate_task.fail()
            return False


class PriceSequenceTask(Task):
    max_retries = 3
    time_limit = 600

    def run(self, *args, **kwargs):
        task_id = kwargs.get("task_id")
        current_file_id = kwargs.get("current_file_id")
        last_month_file_id = kwargs.get("last_month_file_id")
        last_year_file_id = kwargs.get("last_year_file_id")

        calculate_task = CalculateTask.objects.get(id=task_id)
        calculate_task.execute()

        try:
            current_input_data = FileContent.objects.get(file_id=current_file_id).content
            last_month_input_data = FileContent.objects.get(file_id=last_month_file_id).content
            last_year_input_data = FileContent.objects.get(file_id=last_year_file_id).content

            calculate_utils = PriceSequenceUtils(input_current=current_input_data,
                                                 input_last_month=last_month_input_data,
                                                 input_last_year=last_year_input_data)
            raw_result = calculate_utils.get_result()

            task_result = PriceSequenceCalculateResult(task=calculate_task,
                                                       file_current_id=current_file_id,
                                                       file_last_month_id=last_month_file_id,
                                                       file_last_year_id=last_year_file_id,
                                                       details=raw_result)
            task_result.save()

            calculate_task.finish()
            return True
        except:
            calculate_task.fail()
            return False
