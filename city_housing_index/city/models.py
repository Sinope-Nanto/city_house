from django.db import models
from city.enums import CityArea


# Create your models here.
class City(models.Model):
    name = models.CharField(default="", blank=True, null=False, max_length=100, help_text="城市名称")
    code = models.CharField(default="", blank=True, null=False, max_length=200, help_text="城市编码")
    area = models.IntegerField(default=CityArea.NULL, blank=True, null=False, help_text="城市所属区域")
    line = models.IntegerField(default=CityArea.NULL, blank=True, null=True, help_text="几线城市")
    special_area = models.IntegerField(default=CityArea.NULL, blank=True, null=True, help_text="是否属于长三角，珠三角，环渤海")
    block = models.IntegerField(default=CityArea.NULL, blank=True, null=False, help_text="中部、东部与西部")
    ifin40 = models.BooleanField(default=False, help_text="是否属于40城市")
    ifin90 = models.BooleanField(default=False, help_text="是否属于90城市")
    num_block = models.IntegerField(default=1, help_text="城市区块数")
    info_block = models.JSONField(help_text="城市区块信息")