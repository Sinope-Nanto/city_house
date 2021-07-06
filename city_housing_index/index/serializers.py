from rest_framework import serializers
from .models import GenReportTaskRecord, GenReportTaskStatus, CalculateResult


class GenReportTaskSerializer(serializers.ModelSerializer):
    task_id = serializers.SerializerMethodField()
    download_url = serializers.SerializerMethodField()
    errmsg = serializers.SerializerMethodField()

    class Meta:
        model = GenReportTaskRecord
        fields = ["id", "code", "progress", "current_task", "finished", "task_id", "download_url", "errmsg"]

    def get_task_id(self, obj: GenReportTaskRecord):
        return obj.id

    def get_download_url(self, obj: GenReportTaskRecord):
        if obj.finished and obj.status == GenReportTaskStatus.SUCCESS:
            return obj.result.get("download_url")
        return ""

    def get_errmsg(self, obj: GenReportTaskRecord):
        if obj.finished and obj.status == GenReportTaskStatus.ERROR:
            return obj.result.get("errmsg")
        return ""


class CalculateResultSimpleSerializer(serializers.ModelSerializer):
    warn = serializers.SerializerMethodField()

    class Meta:
        model = CalculateResult
        fields = ['year_on_year_index', 'chain_index', 'year_on_year_index_above144', 'chain_index_above144',
                  'year_on_year_index_under90', 'chain_index_under90', 'year_on_year_index_90144', 'chain_index_90144',
                  'warn']

    def get_warn(self, obj: CalculateResult):
        yoy, mom = obj.is_warn()
        return {
            "yoy": yoy,
            "mom": mom
        }


class CalculateResultYoYWarnSerializer(serializers.ModelSerializer):
    class Meta:
        model = CalculateResult
        fields = ['year_on_year_index', 'year_on_year_index_above144',
                  'year_on_year_index_under90', 'year_on_year_index_90144']


class CalculateResultChainWarnSerializer(serializers.ModelSerializer):
    class Meta:
        model = CalculateResult
        fields = ['chain_index', 'chain_index_above144', 'chain_index_under90', 'chain_index_90144']
