from index.plot import PlotReport
from index.models import CalculateResult
from city.models import City
from city.enums import CityArea


def plot(year, month):
    # 绘制全国图
    data = []
    data.extend(
        list(CalculateResult.objects.filter(city_or_area=False, area=CityArea.QUNGUO, year__in=range(2006, year),
                                            month__in=range(1, 13)).order_by('year', 'month')))
    data.extend(list(CalculateResult.objects.filter(city_or_area=False, area=CityArea.QUNGUO, year=year,
                                                    month__in=range(1, month + 1)).order_by('year', 'month')))
    lenth = len(data)
    # 绘制全国同比图
    year_on_year_index = []
    for i in range(12, lenth):
        year_on_year_index.append(data[i].year_on_year_index)
    p = PlotReport()
    p.plot_year_on_year_radio(data=year_on_year_index,
                              saveurl=('media/image/' + str(year) + '_' + str(month) + 'yearonyearplot.png'))
    # 绘制全国环比图
    chain_index = []
    for i in range(1, lenth):
        chain_index.append(data[i].chain_index)
    p.plot_chain_radio(data=chain_index, saveurl=('media/image/' + str(year) + '_' + str(month) + 'chainplot.png'))
    # 绘制指数图
    index = []
    for i in range(0, lenth):
        index.append(data[i].index_value)
    p.plot_index(data=index, saveurl=('media/image/' + str(year) + '_' + str(month) + 'index.png'))
    # 绘制销量-指数图
    vol = []
    for i in range(0, lenth):
        vol.append(data[i].trade_volume)
    p.plot_vol_index(volume=vol, index=index, saveurl=('media/image/' + str(year) + '_' + str(month) + 'volindex.png'))

    # 绘制全国图 90
    data_90 = []
    data_90.extend(
        list(CalculateResult.objects.filter(city_or_area=False, area=CityArea.QUANGUO_90, year__in=range(2009, year),
                                            month__in=range(1, 13)).order_by('year', 'month')))
    data_90.extend((list(CalculateResult.objects.filter(city_or_area=False, area=CityArea.QUANGUO_90, year=year,
                                                        month__in=range(1, month + 1)).order_by('year', 'month'))))
    lenth_90 = len(data_90)
    # 绘制全国同比图
    year_on_year_index_90 = []
    for i in range(12, lenth_90):
        year_on_year_index_90.append(data_90[i].year_on_year_index)
    p.plot_year_on_year_radio_90(data=year_on_year_index_90,
                                 saveurl=('media/image/' + str(year) + '_' + str(month) + 'yearonyearplot_90.png'))
    # 绘制全国环比图
    chain_index_90 = []
    for i in range(1, lenth_90):
        chain_index_90.append(data_90[i].chain_index)
    p.plot_chain_radio_90(data=chain_index_90,
                          saveurl=('media/image/' + str(year) + '_' + str(month) + 'chainplot_90.png'))
    # 绘制指数图
    index_90 = []
    for i in range(0, lenth_90):
        index_90.append(data_90[i].index_value_base09)
    p.plot_index_90(data=index_90, saveurl=('media/image/' + str(year) + '_' + str(month) + 'index_90.png'))
    # 绘制销量-指数图
    vol_90 = []
    for i in range(0, lenth_90):
        vol_90.append(data_90[i].trade_volume)
    p.plot_vol_index_90(volume=vol, index=index,
                        saveurl=('media/image/' + str(year) + '_' + str(month) + 'volindex_90.png'))

    # 绘制全国指数图 - 按面积分片 -简版与完整版
    index_area = [[], [], []]
    vol_area = [[], [], []]
    for i in range(0, lenth):
        index_area[0].append(data[i].index_value_under90)
        index_area[1].append(data[i].index_value_90144)
        index_area[2].append(data[i].index_value_above144)
        vol_area[0].append(data[i].trade_volume_under_90)
        vol_area[1].append(data[i].trade_volume_90_144)
        vol_area[2].append(data[i].trade_volume_above_144)
    p.plot_index_by_area(index_area, saveurl=('media/image/' + str(year) + '_' + str(month) + 'index_by_buildarea.png'))
    p.plot_vol_index_by_area_complex(vol_area, index_area,
                                     'media/image/' + str(year) + '_' + str(month) + 'index_by_buildarea_c.png')

    # 绘制全国指数图 - 按面积分片 -简版与完整版
    index_area_90 = [[], [], []]
    vol_area_90 = [[], [], []]
    for i in range(0, lenth_90):
        index_area_90[0].append(data_90[i].index_value_under90_base09)
        index_area_90[1].append(data_90[i].index_value_90144_base09)
        index_area_90[2].append(data_90[i].index_value_above144_base09)
        vol_area_90[0].append(data_90[i].trade_volume_under_90)
        vol_area_90[1].append(data_90[i].trade_volume_90_144)
        vol_area_90[2].append(data_90[i].trade_volume_above_144)
    p.plot_index_by_area_90(index_area_90,
                            saveurl=('media/image/' + str(year) + '_' + str(month) + 'index_by_buildarea_90.png'))
    p.plot_vol_index_by_area_complex_90(vol_area_90, index_area_90,
                                        'media/image/' + str(year) + '_' + str(month) + 'index_by_buildarea_c_90.png')

    # 绘制全国指数图 - 按区域划分 -简版与完整版
    index_block = [[], [], []]
    vol_block = [[], [], []]
    block = [CityArea.DONGBU, CityArea.ZHONGBU, CityArea.XIBU]
    for i in range(0, 3):
        data = list(CalculateResult.objects.filter(city_or_area=False, area=block[i], year__in=range(2006, year),
                                                   month__in=range(1, 13)).order_by('year', 'month'))
        lenth = len(data)
        for j in range(0, lenth):
            index_block[i].append(data[j].index_value)
            vol_block[i].append(data[j].trade_volume)
    p.plot_index_by_block(index_block, saveurl=('media/image/' + str(year) + '_' + str(month) + 'index_by_block.png'))
    p.plot_vol_index_by_block_complex(vol_block, index_block,
                                      'media/image/' + str(year) + '_' + str(month) + 'index_by_block_c.png')

    # 绘制全国指数图 - 按区域划分 -简版与完整版 90
    index_block_90 = [[], [], []]
    vol_block_90 = [[], [], []]
    block_90 = [CityArea.DONGBU_90, CityArea.ZHONGBU_90, CityArea.XIBU_90]
    for i in range(0, 3):
        data = list(CalculateResult.objects.filter(city_or_area=False, area=block_90[i], year__in=(2009, year),
                                                   month__in=(1, 13)).order_by('year', 'month'))
        for j in range(0, len(data)):
            index_block_90[i].append(data[j].index_value_base09)
            vol_block_90[i].append(data[j].trade_volume)
    p.plot_index_by_block_90(index_block_90,
                             saveurl=('media/image/' + str(year) + '_' + str(month) + 'index_by_block_90.png'))
    p.plot_vol_index_by_block_complex_90(vol_block_90, index_block_90,
                                         'media/image/' + str(year) + '_' + str(month) + 'index_by_block_c_90.png')

    # 绘制全国指数图 - 按线划分 -简版与完整版 90
    index_block_90 = [[], [], [], []]
    vol_block_90 = [[], [], [], []]
    block_90 = [CityArea.FIRSTLINE, CityArea.SECONDLINE, CityArea.THIRDLINE, CityArea.FORTHLINE]
    for i in range(0, 4):
        data = list(CalculateResult.objects.filter(city_or_area=False, area=block_90[i], year__in=range(2009, year),
                                                   month__in=range(1, 13)).order_by('year', 'month'))
        for j in range(0, len(data)):
            index_block_90[i].append(data[j].index_value_base09)
            vol_block_90[i].append(data[j].trade_volume)
    p.plot_index_by_line(index_block_90, saveurl=('media/image/' + str(year) + '_' + str(month) + 'index_by_line.png'))
    p.plot_vol_index_by_line_complex(vol_block_90, index_block_90,
                                     'media/image/' + str(year) + '_' + str(month) + 'index_by_line_c.png')

    # 绘制特殊指数图 -简版与完整版 90
    index_block_90 = [[], [], []]
    vol_block_90 = [[], [], []]
    block_90 = [CityArea.ZHUSANJIAO, CityArea.CHANGSANJIAO_90, CityArea.HUANBOHAI]
    for i in range(0, 3):
        data = list(CalculateResult.objects.filter(city_or_area=False, area=block_90[i], year__in=range(2009, year),
                                                   month__in=range(1, 13)).order_by('year', 'month'))
        for j in range(0, len(data)):
            index_block_90[i].append(data[j].index_value_base09)
            vol_block_90[i].append(data[j].trade_volume)
    p.plot_index_by_block_90(index_block_90,
                             saveurl=('media/image/' + str(year) + '_' + str(month) + 'index_by_s_area.png'))
    p.plot_vol_index_by_block_complex_90(vol_block_90, index_block_90,
                                         'media/image/' + str(year) + '_' + str(month) + 'index_by_s_area_c.png')

    # 绘制地区 -简版与完整版 90
    index_block_90 = [[] for i in range(0, 7)]
    vol_block_90 = [[] for i in range(0, 7)]
    block_90 = [CityArea.DONGBEI, CityArea.HUABEI, CityArea.HUADONG, CityArea.HUAZHONG, CityArea.HUANAN, CityArea.XINAN,
                CityArea.XIBEI]
    for i in range(0, 7):
        data = list(CalculateResult.objects.filter(city_or_area=False, area=block_90[i], year__in=range(2009, year),
                                                   month__in=range(1, 13)).order_by('year', 'month'))
        lenth = len(data)
        for j in range(0, lenth):
            index_block_90[i].append(data[j].index_value_base09)
            vol_block_90[i].append(data[j].trade_volume)
    p.plot_index_by_7area(index_block_90,
                          saveurl=('media/image/' + str(year) + '_' + str(month) + 'index_by_7area.png'))
    p.plot_vol_index_by_7area_complex(vol_block_90, index_block_90,
                                      'media/image/' + str(year) + '_' + str(month) + 'index_by_7area_c.png')

    # 绘制城市销量-指数图 40 and 90
    city_list = City.objects.filter(ifin40=True)
    city_list_90 = City.objects.filter(ifin90=True)
    for city in city_list:
        data = list(CalculateResult.objects.filter(city_or_area=True, city=int(city.code), year__in=range(2006, year),
                                                   month__in=range(1, 13)).order_by('year', 'month'))
        data.extend(list(CalculateResult.objects.filter(city_or_area=True, city=int(city.code), year=year,
                                                        month__in=range(1, month + 1)).order_by('year', 'month')))
        city_index = []
        city_vol = []
        if data[0].index_value == 0:
            data[0].index_value = 100
        last_value = 0
        for i in range(0, len(data)):
            if data[i].index_value == 0:
                city_index.append(data[last_value].index_value)
            else:
                city_index.append(data[i].index_value)
                last_value = i
            city_vol.append(data[i].trade_volume)
        p.plot_vol_index(volume=city_vol, index=city_index,
                         saveurl=('media/image/' + str(year) + '_' + str(month) + 'volindex_' + city.code + '.png'))
    for city in city_list_90:
        data = list(CalculateResult.objects.filter(city_or_area=True, city=int(city.code), year__in=range(2009, year),
                                                   month__in=range(1, 13)).order_by('year', 'month'))
        data.extend(list(CalculateResult.objects.filter(city_or_area=True, city=int(city.code), year=year,
                                                        month__in=range(1, month + 1)).order_by('year', 'month')))
        city_index = []
        city_vol = []
        if data[0].index_value == 0:
            data[0].index_value = 100
        last_value = 0
        for i in range(0, len(data)):
            if data[i].index_value == 0:
                city_index.append(data[last_value].index_value_base09)
            else:
                city_index.append(data[i].index_value_base09)
                last_value = i
            city_vol.append(data[i].trade_volume)
        p.plot_vol_index_90(volume=city_vol, index=city_index, saveurl=(
                'media/image/' + str(year) + '_' + str(month) + 'volindex_' + city.code + '_90.png'))
    return True
