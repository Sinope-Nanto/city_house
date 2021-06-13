from rest_framework import serializers
from .models import DataFile, FileContent, ModelCalculateResult, PriceSequenceCalculateResult, CalculateTask
from city.models import City


class DataFileSerializer(serializers.ModelSerializer):
    city = serializers.SerializerMethodField()

    class Meta:
        model = DataFile
        fields = ['id', "name", "file", "code", "city_code", "start", "end", "city"]

    def get_city(self, obj: DataFile):
        city = City.objects.get(code=obj.city_code)
        return {
            "id": city.id,
            "name": city.name,
            "code": city.code
        }


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
