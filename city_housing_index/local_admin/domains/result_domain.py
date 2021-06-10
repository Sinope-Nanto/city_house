from index.models import CalculateResult
from city.models import City

from index.serializers import CalculateResultSimpleSerializer


def list_all_city_calculate_result(year, month):
    cities = City.objects.all().values('id', 'name', 'code')
    city_dict = {item['code']: item for item in cities}
    result_list = []

    for city_code in city_dict:
        try:
            calculate_result = CalculateResult.objects.get(city=city_code, year=year, month=month)
            calculate_result_data = CalculateResultSimpleSerializer(calculate_result).data
            calculate_result_data['uploaded'] = True
        except:
            calculate_result_data = CalculateResultSimpleSerializer(CalculateResult(city=city_code)).data
            calculate_result_data['uploaded'] = False
        calculate_result_data['city_id'] = city_dict[city_code]['id']
        calculate_result_data['city_name'] = city_dict[city_code]['name']
        calculate_result_data['city_code'] = city_code
        result_list.append(calculate_result_data)
    return result_list
