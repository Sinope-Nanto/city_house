from django.contrib import admin
from index.models import CityIndex
from index.models import CalculateResult
from django.core.exceptions import ObjectDoesNotExist
from city.models import City
from city.enums import CityArea


def CalculateCityIndex(city_id:int,year:int,month:int):
    index = 0
    index_under_90 = 0
    index_above_144 = 0
    index_90_144 = 0
    result = CalculateResult.objects.get(city_or_area=True, city=city_id, year=year, month=month)
    try:
        now = CityIndex.objects.get(city_id=city_id, year=year, month=month)
        result.area_volume = now.area_volume
        result.area_volume_under_90 = now.area_volume_under_90
        result.area_volume_90_144 = now.area_volume_90_144
        result.area_volume_above_144 = now.area_volume_above_144
        now_price = now.price
        now_price_under_90 = now.price_under_90
        now_price_above_144 = now.price_above_144
        now_price_90_144 = now.price_90_144
        now_trade_volumn = now.trade_volume
        now_trade_volumn_under_90 = now.trade_volume_under_90
        now_trade_volumn_90_144 = now.trade_volume_90_144
        now_trade_volumn_above_144 = now.trade_volume_above_144
    except ObjectDoesNotExist:
        result.area_volume = 0
        now_price = 0
        now_price_under_90 = 0
        now_price_above_144 = 0
        now_price_90_144 = 0
        result.area_volume_under_90 = 0
        result.area_volume_90_144 = 0
        result.area_volume_above_144 = 0
        now_trade_volumn = 0
        now_trade_volumn_under_90 = 0
        now_trade_volumn_90_144 = 0
        now_trade_volumn_above_144 = 0
    
    base = CalculateResult.objects.get(city_or_area=True, city=city_id, year=2006, month=1)
    try: 
        index = now_price/base.price
    except ZeroDivisionError:
        index = 0
    try:
        index_90_144 = now_price_90_144/base.price_90_144
    except ZeroDivisionError:
        index_90_144 = 0
    try:
        index_above_144 = now_price_above_144/base.price_above_144
    except ZeroDivisionError:
        index_above_144 = 0
    try:
        index_under_90 = now_price_under_90/base.price_under_90
    except ZeroDivisionError:
        index_under_90 = 0

    result.price = now_price
    result.price_under_90 = now_price_under_90
    result.price_90_144 = now_price_90_144
    result.price_above_144 = now_price_above_144

    result.index_value = index
    result.index_value_90144 = index_90_144
    result.index_value_above144 = index_above_144
    result.index_value_under90 = index_under_90

    result.trade_volume = now_trade_volumn
    result.trade_volume_under_90 = now_trade_volumn_under_90
    result.trade_volume_90_144 = now_trade_volumn_90_144
    result.trade_volume_above_144 = now_trade_volumn_above_144

    result.save()
    return True

def CalculateAreaIndex(areaType:str,areaID:int, year:int, month:int):
    if areaType == 'block':
        city_list = City.objects.filter(block=areaID, ifin40=True)
    elif areaType == 'area':
        city_list = City.objects.filter(area=areaID, ifin40=True)
    elif areaType == 'all':
        city_list = City.objects.filter(ifin40=True)
    elif areaType == 's_area':
        city_list = City.objects.filter(special_area=areaID, ifin40=True)
    elif areaType == 'line':
        city_list = City.objects.filter(line=areaID, ifin40=True)
    else:
        return False
    total_number = 0.0
    total_price = 0.0
    total_number_under_90 = 0.0
    total_price_under_90 = 0.0
    total_number_above_144 = 0.0
    total_price_above_144 = 0.0
    total_number_90_144 = 0.0
    total_price_90_144 = 0.0

    total_trade_volumn = 0
    total_trade_volumn_under_90 = 0
    total_trade_volumn_90_144 = 0
    total_trade_volumn_above_144 = 0
    for city in city_list:
        city_id = int(city.code)
        try:
            nowdata = CityIndex.objects.get(city_id=city_id, year=year, month=month)
            total_number += nowdata.area_volume
            total_price += (nowdata.area_volume * nowdata.price)
            total_number_under_90 += nowdata.area_volume_under_90
            total_price_under_90 += (nowdata.area_volume_under_90 * nowdata.price_under_90)
            total_number_90_144 += nowdata.area_volume_under_90
            total_price_90_144 += (nowdata.area_volume_90_144 * nowdata.price_90_144)
            total_number_above_144 = nowdata.area_volume_above_144
            total_price_above_144 = (nowdata.area_volume_above_144 * nowdata.price_above_144)

            total_trade_volumn += nowdata.trade_volume
            total_trade_volumn_under_90 += nowdata.trade_volume_under_90
            total_trade_volumn_90_144 += nowdata.trade_volume_90_144
            total_trade_volumn_above_144 += nowdata.trade_volume_above_144
        except ObjectDoesNotExist:
            pass
    try:
        bar_price = total_price/total_number
    except ZeroDivisionError:
        bar_price = 0
    try:
        bar_price_under_90 = total_price_under_90/total_number_under_90
    except ZeroDivisionError:
        bar_price_under_90 = 0
    try:
        bar_price_90_144 = total_price_90_144/total_number_90_144
    except ZeroDivisionError:
        bar_price_90_144 = 0
    try:
        bar_price_above_144 = total_price_above_144/total_number_above_144
    except ZeroDivisionError:
        bar_price_above_144 = 0
    result = CalculateResult.objects.get(city_or_area=False, area=areaID, year=year, month=month)
    result.price = bar_price
    result.price_under_90 = bar_price_under_90
    result.price_90_144 = bar_price_90_144
    result.price_above_144 = bar_price_above_144

    result.trade_volume = total_trade_volumn
    result.trade_volume_under_90 = total_trade_volumn_under_90
    result.trade_volume_90_144 = total_trade_volumn_90_144
    result.trade_volume_above_144 = total_trade_volumn_above_144

    result.area_volume = total_number
    result.area_volume_under_90 = total_number_under_90
    result.area_volume_above_144 = total_number_90_144
    result.area_volume_90_144 = total_number_above_144
    
    base = CalculateResult.objects.get(city_or_area=False, area=areaID, year=2006, month=1)
    try:
        result.index_value = bar_price/base.price
    except ZeroDivisionError:
        result.index_value = 0 
    try:
        result.index_value_under90 = bar_price_under_90/base.price_under_90
    except ZeroDivisionError:
        result.index_value_under90 = 0
    try:
        result.index_value_90144 = bar_price_90_144/base.price_90_144
    except ZeroDivisionError:
        result.index_value_90144 = 0
    try:
        result.index_value_above144 = bar_price_above_144/base.price_above_144
    except ZeroDivisionError:
        result.index_value_above144 = 0
    result.save()
    return True



