from calculate.models import DataFile, FileContent, TemplateFiles
import traceback
from utils.excel_utils import read_xls, read_xlsx
from local_auth.models import UserProfile


def create_upload_file(user, file, name, code, city_code, start, end):
    try:
        data_file = DataFile(user=user, file=file, name=name, code=code, city_code=city_code, start=start, end=end)
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
    return raw_data


def check_file_permission(user, file_id):
    if UserProfile.is_user_admin(user):
        return True
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


def get_template_file():
    template_file = TemplateFiles.objects.all().first()
    return template_file.file.url
