from rest_framework import serializers
from .models import DataFile, FileContent, ModelCalculateResult, PriceSequenceCalculateResult, CalculateTask


class DataFileSerializer(serializers.ModelSerializer):
    class Meta:
        model = DataFile
        fields = ['id', "name", "file", "code", "city_id", "start", "end"]


class FileContentSerializer(serializers.ModelSerializer):
    class Meta:
        model = FileContent
        fields = ['file_id', 'title', 'content']


class ModelCalculateResultSerializer(serializers.ModelSerializer):
    class Meta:
        model = ModelCalculateResult
        fields = ['id', 'task_id', 'record_date', 'file_name', 'details']


class CalculateTaskSerializer(serializers.ModelSerializer):
    class Meta:
        model = CalculateTask
        fields = ['id', 'code', 'start', 'end', 'status']


class PriceSequenceCalculateResultSerializer(serializers.ModelSerializer):
    class Meta:
        model = PriceSequenceCalculateResult
        fields = ['id', 'task_id', 'record_date', 'details']
