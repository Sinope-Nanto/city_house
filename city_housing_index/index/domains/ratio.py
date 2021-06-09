from django.contrib import admin
from index.models import CityIndex
from index.models import CalculateResult
from django.core.exceptions import ObjectDoesNotExist
from city.models import City
from city.enums import CityArea


def CalculateRelative(year: int, month: int):
    row_list = CalculateResult.objects.filter(year=year, month=month)
    for now_value in row_list:
        last_value = CalculateResult.objects.get(city_or_area=now_value.city_or_area,
                                                 city=now_value.city,
                                                 area=now_value.area,
                                                 year=year if month > 1 else year - 1,
                                                 month=month if month > 1 else 12)
        now_value.volume_chain = 0 if last_value.trade_volume == 0 else (
                now_value.trade_volume / last_value.trade_volume - 1)
        now_value.chain_index = 0 if last_value.index_value == 0 else (
                now_value.index_value / last_value.index_value - 1)
        now_value.chain_index_above144 = 0 if last_value.index_value_above144 == 0 else (
                now_value.index_value_above144 / last_value.index_value_above144 - 1)
        now_value.chain_index_90144 = 0 if last_value.index_value_90144 == 0 else (
                now_value.index_value_90144 / last_value.index_value_90144 - 1)
        now_value.chain_index_under90 = 0 if last_value.index_value_under90 == 0 else (
                now_value.index_value_under90 / last_value.index_value_under90 - 1)

        last_value = CalculateResult.objects.get(city_or_area=now_value.city_or_area,
                                                 city=now_value.city,
                                                 area=now_value.area,
                                                 year=year - 1,
                                                 month=month
                                                 )

        now_value.volume_chain = 0 if last_value.trade_volume == 0 else (
                now_value.trade_volume / last_value.trade_volume - 1)
        now_value.chain_index = 0 if last_value.index_value == 0 else (
                now_value.index_value / last_value.index_value - 1)
        now_value.chain_index_above144 = 0 if last_value.index_value_above144 == 0 else (
                now_value.index_value_above144 / last_value.index_value_above144 - 1)
        now_value.chain_index_90144 = 0 if last_value.index_value_90144 == 0 else (
                now_value.index_value_90144 / last_value.index_value_90144 - 1)
        now_value.chain_index_under90 = 0 if last_value.index_value_under90 == 0 else (
                now_value.index_value_under90 / last_value.index_value_under90 - 1)

        now_value.save()
    return True
