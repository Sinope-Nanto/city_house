from local_auth.models import UserProfile
from city.models import City
from utils.sms_utils import SMSSender

CITY_SMS_CONTENT_TEMPLATE = "【城房指数】请您及时上传【{}】{}年{}月的城房指数数据"


def parse_year_month(date_str):
    year_month_list = date_str.split('-')
    if len(year_month_list) < 2:
        return -1, -1
    return int(year_month_list[0]), int(year_month_list[1])


def get_city_upload_status(year, month):
    city_list = City.objects.all()
    data_list = []
    for city in city_list:
        from calculate.models import DataFile

        upload_files = DataFile.objects.filter(city_code=city.code)
        exist = False
        for upload_file in upload_files:
            upload_year, upload_month = parse_year_month(upload_file.start)
            if upload_year == int(year) and upload_month == int(month):
                exist = True
                break
        data_list.append(
            {
                'city_id': city.id,
                'city_code': city.code,
                'city_name': city.name,
                'upload_status': exist
            }
        )
    return data_list

def send_city_sms(city_ids, year, month):
    send_list = []
    not_send_list = []
    cities = City.objects.filter(id__in=city_ids)
    for city in cities:
        user_profile = UserProfile.objects.filter(city_id=city.id, mobile__startswith='1').first()
        if user_profile:
            sms_content = CITY_SMS_CONTENT_TEMPLATE.format(city.name, year, month)
            SMSSender(mobile=user_profile.mobile).send_sms(sms_content)
            send_list.append(city.name)
        else:
            not_send_list.append(city.name)
    return {
        "send": send_list,
        "not_send": not_send_list
    }
