from django.db import models
from jsonfield.fields import JSONField
from django.contrib.auth.models import User
from django.utils import timezone
from calculate.enums import CalculateTaskStatus, CalculateTaskType


# Create your models here.

class DataFile(models.Model):
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    name = models.CharField(default="", max_length=200, blank=True)
    file = models.FileField(upload_to="origin_data")
    code = models.CharField(default="", max_length=100, blank=True)
    city_code = models.IntegerField(default=0)
    start = models.CharField(default="", max_length=100, blank=True)
    end = models.CharField(default="", max_length=100, blank=True)
    upload_date = models.DateTimeField(auto_now_add=True, null=True)
    deleted = models.BooleanField(default=False)


class FileContent(models.Model):
    file = models.ForeignKey(DataFile, on_delete=models.SET_NULL, null=True)
    title = JSONField(default=[], null=False)
    content = JSONField(default=[], null=False)


class TemplateFiles(models.Model):
    city_code = models.IntegerField(default=0, verbose_name="城市编号")
    file = models.FileField(upload_to="template_file")


class CalculateTask(models.Model):
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    code = models.CharField(default="", max_length=200, blank=True)
    type = models.CharField(default=CalculateTaskType.MODEL_CALCULATE, max_length=50, blank=True)
    kwargs = JSONField(default={}, null=False)
    start = models.DateTimeField(auto_now_add=True, null=True)
    end = models.DateTimeField(null=True)
    status = models.CharField(default=CalculateTaskStatus.START, max_length=200, blank=True)

    @classmethod
    def generate_code(cls):
        import uuid
        return uuid.uuid4().hex

    def execute(self):
        self.start = timezone.now()
        self.status = CalculateTaskStatus.EXECUTING
        self.save()

    def finish(self):
        self.end = timezone.now()
        self.status = CalculateTaskStatus.SUCCESS
        self.save()

    def fail(self):
        self.end = timezone.now()
        self.status = CalculateTaskStatus.FAIL
        self.save()


class ModelCalculateResult(models.Model):
    task = models.ForeignKey(CalculateTask, on_delete=models.SET_NULL, null=True)
    file_name = models.CharField(default="", max_length=100, blank=True)
    data_file = models.ForeignKey(DataFile, on_delete=models.SET_NULL, null=True)
    record_date = models.DateTimeField(auto_now_add=True, null=True)
    details = JSONField(default={}, null=False)


class PriceSequenceCalculateResult(models.Model):
    task = models.ForeignKey(CalculateTask, on_delete=models.SET_NULL, null=True)
    file_current_id = models.IntegerField(default=0)
    file_last_month_id = models.IntegerField(default=0)
    file_last_year_id = models.IntegerField(default=0)
    record_date = models.DateTimeField(auto_now_add=True, null=True)
    details = JSONField(default={}, null=False)
