from calculate.models import DataFile, FileContent, TemplateFiles
import traceback
from utils.excel_utils import read_xls, read_xlsx, write_xlsx
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


def get_template_file(user_id):
    user_profile = UserProfile.get_by_user_id(user_id)
    city = user_profile.city
    base_template = TemplateFiles.objects.get(city_code=0)
    generated_file_path = generate_new_template_file(base_template.file.path, city.info_block["info"], city.code)
    generated_file = open(generated_file_path, 'rb')
    city_template = TemplateFiles.objects.filter(city_code=city.code).first()
    if city_template:
        city_template.file.delete()

    else:
        city_template = TemplateFiles(city_code=city.code)
    city_template.file.save("{}模板数据.xlsx".format(city.name), generated_file)
    return city_template.file.url

def generate_new_template_file(base_file, blocks, city_code):
    title, content = read_xlsx(base_file)
    for block in blocks:
        title.append(block["code"])

    from django.conf import settings
    import os
    file_path = os.path.join(settings.MEDIA_ROOT, "template_file/city-{}.xlsx".format(city_code))
    write_xlsx(file_path, "origin-data", title, content)
    return file_path
