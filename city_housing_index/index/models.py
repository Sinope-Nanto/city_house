from django.db import models
from city.models import City
from city.enums import CityArea

# Create your models here.
class CityIndex(models.Model):
    #某个城市某月的信息
    city = models.IntegerField(default=0,blank=False,null=False,help_text='城市代码')
    year = models.IntegerField(default=0,null=False)
    month = models.IntegerField(default=0,null=False)
    price = models.FloatField(default=0,null=False,help_text="总平均价格")

    trade_volume = models.IntegerField(default=0,null=False,help_text="总交易额")
    trade_volume_under_90 = models.FloatField(default=0,help_text="90平米以下总交易额")
    trade_volume_above_144 = models.FloatField(default=0)
    trade_volume_90_144 = models.FloatField(default=0)

    area_volume = models.FloatField(default=0,help_text="总成交面积")
    area_volume_under_90 = models.FloatField(default=0,help_text="总成交面积u90")
    area_volume_above_144 = models.FloatField(default=0,help_text="总成交面积a144")
    area_volume_90_144 = models.FloatField(default=0,help_text="总成交面积90-144")

    price_under_90 = models.FloatField(default=0,null=False,help_text="90平米以下平均价格")
    price_above_144 = models.FloatField(default=0,null=False)
    price_90_144 = models.FloatField(default=0,null=False)

class CalculateResult(models.Model):
    #表头基本信息字段，创建行时即要填入
    year = models.IntegerField(default=0,null=False)
    month = models.IntegerField(default=0,null=False)
    #代表此行代表一个城市(True)，或者一个城市集合(False)
    #如东部、西部、全国、一线城市、二线城市、长三角、珠三角等均为一个城市集合
    city_or_area = models.BooleanField(default=True)
    #城市与地区代码，地区代码见city.enums
    area = models.IntegerField(default=CityArea.NULL)
    city = models.IntegerField(null=True)

    #在执行CalculateXXXIndex时填入的字段
    #该段中字段的含义与CityIndex中同名字段相同
    price = models.FloatField(default=0)
    price_under_90 = models.FloatField(default=0)
    price_above_144 = models.FloatField(default=0)
    price_90_144 = models.FloatField(default=0)

    index_value = models.FloatField(default=0,help_text='指数值')
    index_value_under90 = models.FloatField(default=0,help_text='90平米以下指数值')
    index_value_above144 = models.FloatField(default=0,help_text='144平米以上指数值')
    index_value_90144 = models.FloatField(default=0)
    area_volume = models.FloatField(default=0,help_text="总成交面积")
    area_volume_under_90 = models.FloatField(default=0,help_text="总成交面积u90")
    area_volume_above_144 = models.FloatField(default=0,help_text="总成交面积a144")
    area_volume_90_144 = models.FloatField(default=0,help_text="总成交面积90-144")
    trade_volume = models.IntegerField(default=0)
    trade_volume_under_90 = models.FloatField(default=0)
    trade_volume_above_144 = models.FloatField(default=0)
    trade_volume_90_144 = models.FloatField(default=0)

    #在执行CalculateXXXIndex_base09时填入的字段
    #该段中字段的含义与CityIndex中同名字段相同
    index_value_base09 = models.FloatField(default=0,help_text='指数值')
    index_value_under90_base09 = models.FloatField(default=0,help_text='90平米以下指数值')
    index_value_above144_base09 = models.FloatField(default=0,help_text='144平米以上指数值')
    index_value_90144_base09 = models.FloatField(default=0)

    #在计算同比、环比时填入的字段
    volume_year_on_year = models.FloatField(default=0)
    volume_chain = models.FloatField(default=0)
    year_on_year_index = models.FloatField(default=0)
    chain_index = models.FloatField(default=0)
    year_on_year_index_above144 = models.FloatField(default=0)
    chain_index_above144 = models.FloatField(default=0)
    year_on_year_index_under90 = models.FloatField(default=0)
    chain_index_under90 = models.FloatField(default=0)
    year_on_year_index_90144 = models.FloatField(default=0)
    chain_index_90144 = models.FloatField(default=0)





