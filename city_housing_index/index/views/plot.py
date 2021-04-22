from rest_framework.views import APIView
from utils.api_response import APIResponse

from local_auth.authentication import CityIndexAuthentication
from local_admin.permissions import CityIndexAdminPermission

from index.plot import PlotReport
from index.models import CalculateResult
from city.models import City
from city.enums import CityArea


def plot(year,month):
    # 绘制全国图
    data = []
    for _year in range(2006,year):
        for _month in range(1,13):
            data.append(CalculateResult.objects.get(city_or_area=False,area=CityArea.QUNGUO,year=_year,month=_month))
    for _month in range(1,month+1):
        data.append(CalculateResult.objects.get(city_or_area=False,area=CityArea.QUNGUO,year=year,month=_month))
    lenth = len(data)
    # 绘制全国同比图
    year_on_year_index = []
    for i in range(12,lenth):
        year_on_year_index.append(data[i].year_on_year_index)
    p = PlotReport()
    p.plot_year_on_year_radio(data=year_on_year_index,saveurl=('media/image/'+str(year)+'_'+str(month)+'yearonyearplot.png'))
    # 绘制全国环比图
    chain_index = []
    for i in range(1,lenth):
        chain_index.append(data[i].chain_index)
    p.plot_chain_radio(data=chain_index,saveurl=('media/image/'+str(year)+'_'+str(month)+'chainplot.png'))
    # 绘制指数图
    index = []
    for i in range(0,lenth):
        index.append(data[i].index_value)
    p.plot_index(data=index,saveurl=('media/image/'+str(year)+'_'+str(month)+'index.png'))
    # 绘制销量-指数图
    vol = []
    for i in range(0,lenth):
        vol.append(data[i].trade_volume)
    p.plot_vol_index(volume=vol,index=index,saveurl=('media/image/'+str(year)+'_'+str(month)+'volindex.png'))
    
    # 绘制全国指数图 - 按面积分片 -简版与完整版
    index_area = [[],[],[]]
    vol_area = [[],[],[]]
    for i in range(0,lenth):
        index_area[0].append(data[i].index_value_under90)
        index_area[1].append(data[i].index_value_90144)
        index_area[2].append(data[i].index_value_above144)
        vol_area[0].append(data[i].trade_volume_under_90)
        vol_area[1].append(data[i].trade_volume_90_144)
        vol_area[2].append(data[i].trade_volume_above_144)
    p.plot_index_by_area(index_area,saveurl=('media/image/'+str(year)+'_'+str(month)+'index_by_buildarea.png'))
    p.plot_vol_index_by_area_complex(vol_area,index_area,'media/image/'+str(year)+'_'+str(month)+'index_by_buildarea_c.png')


    # 绘制全国指数图 - 按区域划分 -简版与完整版
    index_block = [[],[],[]]
    vol_block = [[],[],[]]
    block = [CityArea.DONGBU,CityArea.ZHONGBU,CityArea.XIBU]
    for i in range(0,3):
        data = []
        for _year in range(2006,year):
            for _month in range(1,13):
                data.append(CalculateResult.objects.get(city_or_area=False,area=block[i],year=_year,month=_month))
        lenth = len(data)
        for j in range(0,lenth):
            index_block[i].append(data[j].index_value)
            vol_block[i].append(data[j].trade_volume)
    p.plot_index_by_block(index_block,saveurl=('media/image/'+str(year)+'_'+str(month)+'index_by_block.png'))
    p.plot_vol_index_by_block_complex(vol_block,index_block,'media/image/'+str(year)+'_'+str(month)+'index_by_block_c.png')
    
    # 绘制城市销量-指数图
    city_list = City.objects.filter(ifin40=True)
    for city in city_list:
        data = []
        for _year in range(2006,year):
            for _month in range(1,13):
                data.append(CalculateResult.objects.get(city_or_area=True,city=int(city.code),year=_year,month=_month))
        for _month in range(1,month+1):
            data.append(CalculateResult.objects.get(city_or_area=True,city=int(city.code),year=year,month=_month))
        lenth = len(data)
        city_index = []
        city_vol = []
        if data[0].index_value == 0:
            data[0].index_value = 100
        last_value = 0
        for i in range(0,lenth):
            if data[i].index_value==0:
                city_index.append(data[last_value].index_value)
            else:
                city_index.append(data[i].index_value)
                last_value = i
            city_vol.append(data[i].trade_volume)
        p.plot_vol_index(volume=city_vol,index=city_index,saveurl=('media/image/'+str(year)+'_'+str(month)+'volindex_'+city.code+'.png'))
    return True




class PlotViews(APIView):

    def post(self,request):
        plot(int(request.data['year']),int(request.data['month']))
        return APIResponse.create_success()      