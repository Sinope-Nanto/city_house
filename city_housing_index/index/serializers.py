from rest_framework import serializers
from .models import GenReportTask, GenReportTaskStatus


class GenReportTaskSerializer(serializers.ModelSerializer):
    task_id = serializers.SerializerMethodField()
    download_url = serializers.SerializerMethodField()
    errmsg = serializers.SerializerMethodField()

    class Meta:
        model = GenReportTask
        fields = ["id", "code", "progress", "current_task", "finished"]

    def get_task_id(self, obj: GenReportTask):
        return obj.id

    def get_download_url(self, obj: GenReportTask):
        if obj.finished and obj.status == GenReportTaskStatus.SUCCESS:
            return obj.result.get("download_url")
        return ""

    def get_errmsg(self, obj: GenReportTask):
        if obj.finished and obj.status == GenReportTaskStatus.ERROR:
            return obj.result.get("errmsg")
        return ""
