from calculate.models import DataFile, FileContent
from city.models import City
from city.serializers import CitySerializer
import traceback
from utils.excel_utils import read_xls, read_xlsx
import json


def create_upload_file(user, file, name, code, city_id, start, end):
    try:
        data_file = DataFile(user=user, file=file, name=name, code=code, city_id=city_id, start=start, end=end)
        data_file.save()

        file_type = data_file.file.url.split(".")[1]
        print(file_type)
        if file_type == "xls":
            title, content = read_xls(data_file.file.path)
        elif file_type == "xlsx":
            title, content = read_xlsx(data_file.file.path)
        else:
            return False, "文件格式异常"

        file_content = FileContent(file_id=data_file.id, title=title, content=content)
        file_content.save()
        return True, ""
    except Exception as ex:
        traceback.print_exc()
        return False, "储存失败"


def list_upload_files(user):
    from calculate.serializers import DataFileSerializer
    data_file_queryset = DataFile.objects.filter(user=user, deleted=False)
    raw_data = DataFileSerializer(data_file_queryset, many=True).data

    city_id_list = [item["city_id"] for item in raw_data]
    # city_queryset = City.objects.filter(code__in=city_id_list)
    city_queryset = City.objects.filter(id__in=city_id_list)

    city_data_list = CitySerializer(city_queryset, many=True).data

    city_dict = {item["id"]: item for item in city_data_list}

    for item in raw_data:
        item.update({"city": city_dict[item["city_id"]]})

    return raw_data


def check_file_permission(user, file_id):
    data_file = DataFile.objects.get(id=file_id)
    return data_file.user_id == user.id


def review_file_content(file_id):
    file_content = FileContent.objects.get(file_id=file_id)
    ret_data = {
        "id": file_content.id,
        "title": file_content.title,
        "content": file_content.content,
    }

    return ret_data


def delete_data_file(file_id):
    data_file = DataFile.objects.get(id=file_id)
    data_file.deleted = True
    data_file.save()
    return True
