from index.models import CalculateResult
from city.models import City
from city.enums import CityArea

from index.domains.genreport import GenExcelReport
from index.domains.genreport import GenWordReport
from index.domains.genreport90 import GenExcelReport90
from index.domains.genreport90 import GenWordReport90
from index.domains.genreport import GenWordPicture
from index.domains.genreport90 import GenWordPicture90
from index.domains.genreport_origindata import GenOriginalReport


def getLastNMonth(year, month, num):
    month -= num
    if month <= 0:
        month += 12
        year -= 1
    return year, month


def get_report_40(year: int, month: int):
    report = GenExcelReport('media/report/' + '40_city_report_' + str(year) + '_' + str(month) + '.xlsx')
    report.create_report()
    # 参数列表:['volumn','east_volumn','mid_volumn','west_volumn',
    # 'under_90_volumn','90_144_volumn','above_144_volumn','volumn_year_on_year','volumn_chain',
    # 'index','east_index','mid_index','west_index',
    # 'year_on_year','chain','east_year_on_year','east_chain',
    # 'mid_year_on_year','mid_chain','west_year_on_year','west_chain',
    # 'index_under_90','index_90_144','index_above_144','year_on_year_under_90',
    # 'chain_under_90','year_on_year_90_144','chain_90_144',
    # 'year_on_year_above_144','chain_above_144','index_CSJ']

    kwargs = {}

    keylist = ['volumn', 'east_volumn', 'mid_volumn', 'west_volumn',
               'under_90_volumn', '90_144_volumn', 'above_144_volumn', 'volumn_year_on_year', 'volumn_chain',
               'index', 'east_index', 'mid_index', 'west_index',
               'year_on_year', 'chain', 'east_year_on_year', 'east_chain',
               'mid_year_on_year', 'mid_chain', 'west_year_on_year', 'west_chain',
               'index_under_90', 'index_90_144', 'index_above_144', 'year_on_year_under_90',
               'chain_under_90', 'year_on_year_90_144', 'chain_90_144',
               'year_on_year_above_144', 'chain_above_144', 'index_CSJ']
    for key in keylist:
        kwargs[key] = []
    # 全国数据
    totalDataList = []
    eastDataList = []
    westDataList = []
    midDataList = []
    CSJDataList = []
    for y in range(2006, year):
        for m in range(1, 13):
            totalDataList.append(CalculateResult.objects.get(city_or_area=False, area=CityArea.QUNGUO, year=y, month=m))
            eastDataList.append(CalculateResult.objects.get(city_or_area=False, area=CityArea.DONGBU, year=y, month=m))
            westDataList.append(CalculateResult.objects.get(city_or_area=False, area=CityArea.XIBU, year=y, month=m))
            midDataList.append(CalculateResult.objects.get(city_or_area=False, area=CityArea.ZHONGBU, year=y, month=m))
            CSJDataList.append(
                CalculateResult.objects.get(city_or_area=False, area=CityArea.CHANGSANJIAO, year=y, month=m))
    for m in range(1, month + 1):
        totalDataList.append(CalculateResult.objects.get(city_or_area=False, area=CityArea.QUNGUO, year=year, month=m))
        eastDataList.append(CalculateResult.objects.get(city_or_area=False, area=CityArea.DONGBU, year=year, month=m))
        westDataList.append(CalculateResult.objects.get(city_or_area=False, area=CityArea.XIBU, year=year, month=m))
        midDataList.append(CalculateResult.objects.get(city_or_area=False, area=CityArea.ZHONGBU, year=year, month=m))
        CSJDataList.append(
            CalculateResult.objects.get(city_or_area=False, area=CityArea.CHANGSANJIAO, year=year, month=m))
    for i in range(0, len(totalDataList)):
        datarow = totalDataList[i]
        eastdatarow = eastDataList[i]
        westdatarow = westDataList[i]
        middatarow = midDataList[i]
        CSJdatarow = CSJDataList[i]
        kwargs['volumn'].append(datarow.trade_volume)
        kwargs['east_volumn'].append(eastdatarow.trade_volume)
        kwargs['mid_volumn'].append(middatarow.trade_volume)
        kwargs['west_volumn'].append(westdatarow.trade_volume)
        kwargs['under_90_volumn'].append(datarow.trade_volume_under_90)
        kwargs['90_144_volumn'].append(datarow.trade_volume_90_144)
        kwargs['above_144_volumn'].append(datarow.trade_volume_above_144)
        kwargs['index'].append(datarow.index_value)
        kwargs['east_index'].append(eastdatarow.index_value)
        kwargs['mid_index'].append(middatarow.index_value)
        kwargs['west_index'].append(westdatarow.index_value)
        kwargs['index_under_90'].append(datarow.index_value_under90)
        kwargs['index_90_144'].append(datarow.index_value_90144)
        kwargs['index_above_144'].append(datarow.index_value_above144)
        kwargs['index_CSJ'].append(CSJdatarow.index_value)

        if i > 11:
            kwargs['volumn_year_on_year'].append(datarow.volume_year_on_year)
            kwargs['year_on_year'].append(datarow.year_on_year_index)
            kwargs['east_year_on_year'].append(eastdatarow.year_on_year_index)
            kwargs['mid_year_on_year'].append(middatarow.year_on_year_index)
            kwargs['west_year_on_year'].append(westdatarow.year_on_year_index)
            kwargs['year_on_year_under_90'].append(datarow.year_on_year_index_under90)
            kwargs['year_on_year_90_144'].append(datarow.year_on_year_index_90144)
            kwargs['year_on_year_above_144'].append(datarow.year_on_year_index_above144)

        if i > 0:
            kwargs['volumn_chain'].append(datarow.volume_chain)
            kwargs['chain'].append(datarow.chain_index)
            kwargs['east_chain'].append(eastdatarow.chain_index)
            kwargs['west_chain'].append(westdatarow.chain_index)
            kwargs['mid_chain'].append(middatarow.chain_index)
            kwargs['chain_under_90'].append(datarow.chain_index_under90)
            kwargs['chain_90_144'].append(datarow.chain_index_90144)
            kwargs['chain_above_144'].append(datarow.chain_index_above144)

    report.IndexSummary(**kwargs)

    # 参数列表:['index','year_on_year','chain','year_on_year_plot','chain_plot']
    kwargs_2 = {}
    kwargs_2['index'] = kwargs['index']
    kwargs_2['year_on_year'] = kwargs['year_on_year']
    kwargs_2['chain'] = kwargs['chain']
    kwargs_2['year_on_year_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'yearonyearplot.png'
    kwargs_2['chain_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'chainplot.png'

    try:
        report.RadioPlot(**kwargs_2)
    except FileNotFoundError:
        return False, ""

    # 参数列表:['url_index_plot','url_block_plot','url_line_plot']
    kwargs_3 = {}
    kwargs_3['url_index_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index.png'
    kwargs_3['url_block_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_block.png'
    kwargs_3['url_line_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_buildarea.png'

    try:
        report.SimplyPlot(**kwargs_3)
    except FileNotFoundError:
        return False, ""

    # 参数列表:['url_index_plot','url_block_plot','url_line_plot']
    kwargs_4 = {}
    kwargs_4['url_index_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'volindex.png'
    kwargs_4['url_block_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_block_c.png'
    kwargs_4['url_line_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_buildarea_c.png'

    try:
        report.ComplexPlot(**kwargs_4)
    except FileNotFoundError:
        return False, ""

    # list每项的参数列表:{'city_name','index_this_month','index_last_month','index_last_year',
    # 'chain_radio','year_on_year','volumn_chain','volumn_year'}

    kwargs_5 = []
    cityList = City.objects.filter(ifin40=True)
    for i in cityList:
        cityindex = CalculateResult.objects.get(city_or_area=True, city=int(i.code), year=year, month=month)
        newcity = {}
        newcity['city_name'] = i.name
        newcity['index_this_month'] = cityindex.index_value
        newcity['chain_radio'] = cityindex.chain_index
        newcity['year_on_year'] = cityindex.year_on_year_index
        try:
            newcity['volumn_chain'] = cityindex.trade_volume / CalculateResult.objects.get(city_or_area=True,
                                                                                           city=int(i.code),
                                                                                           year=year if month > 1 else year - 1,
                                                                                           month=month - 1 if month > 1 else 12).trade_volume - 1

        except ZeroDivisionError:
            newcity['volumn_chain'] = 0
        try:
            newcity['volumn_year'] = cityindex.trade_volume / CalculateResult.objects.get(city_or_area=True,
                                                                                          city=int(i.code),
                                                                                          year=year - 1,
                                                                                          month=month).trade_volume - 1
        except ZeroDivisionError:
            newcity['volumn_year'] = 0

        try:
            newcity['index_last_month'] = newcity['index_this_month'] / (1 + newcity['chain_radio'])
        except ZeroDivisionError:
            newcity['index_last_month'] = CalculateResult.objects.get(city_or_area=True,
                                                                      city=int(i.code),
                                                                      year=year if month > 1 else year - 1,
                                                                      month=month if month > 1 else 12).index_value
        try:
            newcity['index_last_year'] = newcity['index_this_month'] / (1 + newcity['year_on_year'])
        except ZeroDivisionError:
            newcity['index_last_year'] = CalculateResult.objects.get(city_or_area=True, city=int(i.code), year=year - 1,
                                                                     month=month).index_value
        kwargs_5.append(newcity)
    report.CitySummary(kwargs_5)

    # 列表中每个元素应有参数:{'city_name','chain0','chain1' ,'chain2',
    # 'chain3' 'year0' ,'year1' ,'year2','year3'}
    kwargs_6 = []
    for city in cityList:
        newcity = {}
        newcity['city_name'] = city.name
        for i in range(0, 4):
            y, m = getLastNMonth(year, month, i)
            monthData = CalculateResult.objects.get(city_or_area=True, city=int(city.code), year=y, month=m)
            newcity['chain' + str(i)] = monthData.chain_index
            newcity['year' + str(i)] = monthData.year_on_year_index
        kwargs_6.append(newcity)
    report.CitySort(month, kwargs_6)

    citynamelist = []
    citycodelist = []
    kwargs_7 = {}
    kwargs_8 = {}
    for city in cityList:
        citynamelist.append(city.name)
        citycodelist.append(int(city.code))
    for i in range(0, len(citycodelist)):
        kwargs_7[citynamelist[i]] = []
        kwargs_8[citynamelist[i]] = []
        for y in range(2006, year):
            for m in range(1, 13):
                data = CalculateResult.objects.get(city_or_area=True, city=citycodelist[i], year=y, month=m)
                kwargs_7[citynamelist[i]].append(data.index_value)
                kwargs_8[citynamelist[i]].append(data.trade_volume)
        for m in range(1, month + 1):
            data = CalculateResult.objects.get(city_or_area=True, city=citycodelist[i], year=year, month=m)
            kwargs_7[citynamelist[i]].append(data.index_value)
            kwargs_8[citynamelist[i]].append(data.trade_volume)

    report.CityIndex(citynamelist, **kwargs_7)
    report.CityVolumn(citynamelist, **kwargs_8)

    ploturlList = []
    try:
        for i in citycodelist:
            ploturlList.append('media/image/' + str(year) + '_' + str(month) + 'volindex_' + str(i) + '.png')
        report.CityPlot(citynamelist, ploturlList)
    except FileNotFoundError:
        return False, ""
    report.EndReport()

    return True, report.url


def get_word_report_40(year: int, month: int):
    city_list = []
    city_code_list = []
    for city in City.objects.filter(ifin40=True):
        city_list.append(city.name)
        city_code_list.append(city.code)
    report = GenWordReport('media/report/' + '40_city_report_' + str(year) + '_' + str(month) + '.docx')
    try:
        # 取全国数据

        kwargs = {}
        kwargs['index'] = []
        kwargs['chain'] = []
        kwargs['year_on_year'] = []
        for y in range(2008, year):
            for m in range(1, 13):
                total_index = CalculateResult.objects.get(city_or_area=False, area=CityArea.QUNGUO, year=y, month=m)
                kwargs['index'].append(total_index.index_value)
                kwargs['chain'].append(total_index.chain_index)
                kwargs['year_on_year'].append(total_index.year_on_year_index)
        for m in range(1, month + 1):
            total_index = CalculateResult.objects.get(city_or_area=False, area=CityArea.QUNGUO, year=year, month=m)
            kwargs['index'].append(total_index.index_value)
            kwargs['chain'].append(total_index.chain_index)
            kwargs['year_on_year'].append(total_index.year_on_year_index)

        total_index = CalculateResult.objects.get(city_or_area=False, area=CityArea.QUNGUO, year=year, month=month)
        kwargs['index_under_90'] = [total_index.index_value_under90]
        kwargs['chain_under_90'] = [total_index.chain_index_under90]
        kwargs['year_on_year_under_90'] = [total_index.year_on_year_index_under90]

        kwargs['index_90_144'] = [total_index.index_value_90144]
        kwargs['chain_90_144'] = [total_index.chain_index_90144]
        kwargs['year_on_year_90_144'] = [total_index.year_on_year_index_90144]

        kwargs['index_above_144'] = [total_index.index_value_above144]
        kwargs['chain_above_144'] = [total_index.chain_index_above144]
        kwargs['year_on_year_above_144'] = [total_index.year_on_year_index_above144]

        city_chain = []
        city_year_on_year = []
        city_index = []
        city_index_1 = []
        city_index_2 = []

        city_volume = []
        city_volume_1 = []
        city_volume_2 = []

        for city_code in city_code_list:
            city_info = CalculateResult.objects.get(city_or_area=True, city=city_code, year=year, month=month)

            last_year, last_month = getLastNMonth(year, month, 1)
            last_city_info = CalculateResult.objects.get(city_or_area=True, city=city_code, year=last_year,
                                                         month=last_month)
            city_index_1.append(last_city_info.index_value)
            city_volume_1.append(last_city_info.trade_volume)

            last_year, last_month = getLastNMonth(year, month, 2)
            last_city_info = CalculateResult.objects.get(city_or_area=True, city=city_code, year=last_year,
                                                         month=last_month)
            city_index_2.append(last_city_info.index_value)
            city_volume_2.append(last_city_info.trade_volume)

            city_chain.append(city_info.chain_index)
            city_volume.append(city_info.trade_volume)
            city_year_on_year.append(city_info.year_on_year_index)
            city_index.append(city_info.index_value)

        kwargs['city_chain'] = city_chain
        kwargs['city_year_on_year'] = city_year_on_year
        kwargs['city_index'] = city_index
        kwargs['city_index_1'] = city_index_1
        kwargs['city_index_2'] = city_index_2
        kwargs['city_volume'] = city_volume
        kwargs['city_volume_1'] = city_volume_1
        kwargs['city_volume_2'] = city_volume_2

        east_data = CalculateResult.objects.get(city_or_area=False, area=CityArea.DONGBU, year=year, month=month)
        mid_data = CalculateResult.objects.get(city_or_area=False, area=CityArea.ZHONGBU, year=year, month=month)
        west_data = CalculateResult.objects.get(city_or_area=False, area=CityArea.XIBU, year=year, month=month)
        kwargs['east_index'] = [east_data.index_value]
        kwargs['east_chain'] = [east_data.chain_index]
        kwargs['east_year_on_year'] = [east_data.year_on_year_index]

        kwargs['mid_index'] = [mid_data.index_value]
        kwargs['mid_chain'] = [mid_data.chain_index]
        kwargs['mid_year_on_year'] = [mid_data.year_on_year_index]

        kwargs['west_index'] = [west_data.index_value]
        kwargs['west_chain'] = [west_data.chain_index]
        kwargs['west_year_on_year'] = [west_data.year_on_year_index]

        kwargs['index_image_url'] = 'media/image/' + str(year) + '_' + str(month) + 'index.png'
        kwargs['chain_image_url'] = 'media/image/' + str(year) + '_' + str(month) + 'chainplot.png'
        kwargs['yearonyear_image_url'] = 'media/image/' + str(year) + '_' + str(month) + 'yearonyearplot.png'
        kwargs['index_by_block_image_url'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_block.png'
        kwargs['index_by_buildarea_image_url'] = 'media/image/' + str(year) + '_' + str(
            month) + 'index_by_buildarea.png'
        report.Firstpage(year, month, **kwargs)
        report.Secondpage(year, month, city_list, **kwargs)
        report.Charts(year, month, city_list, **kwargs)
        report.Attach(year, month, city_list, **kwargs)
        report.EndReport()
    except:
        return False, ""
    return True, report.url


def get_report_90(year: int, month: int):
    report = GenExcelReport90('media/report/' + '90_city_report_' + str(year) + '_' + str(month) + '.xlsx')

    report.create_report()
    keylist = ['volumn', 'east_volumn', 'mid_volumn', 'west_volumn',
               'under_90_volumn', '90_144_volumn', 'above_144_volumn',
               'DONGBEI_volumn', 'HUABEI_volumn', 'HUADONG_volumn', 'HUAZHONG_volumn', 'HUANAN_volumn', 'XINAN_volumn',
               'XIBEI_volumn',
               'firstline_volumn', 'secondline_volumn', 'thirdline_volumn', 'forth_volumn', 'ZSJ_volumn', 'CSJ_volumn',
               'HBH_volumn',
               'index', 'east_index', 'mid_index', 'west_index',
               'index_under_90', 'index_90_144', 'index_above_144',
               'DONGBEI_index', 'HUABEI_index', 'HUADONG_index', 'HUAZHONG_index', 'HUANAN_index', 'XINAN_index',
               'XIBEI_index',
               'firstline_index', 'secondline_index', 'thirdline_index', 'forth_index', 'ZSJ_index', 'CSJ_index',
               'HBH_index',
               'volumn_year_on_year',
               'year_on_year', 'east_year_on_year', 'mid_year_on_year', 'west_year_on_year',
               'year_on_year_under_90', 'year_on_year_90_144', 'year_on_year_above_144',
               'DONGBEI_year_on_year', 'HUABEI_year_on_year', 'HUADONG_year_on_year', 'HUAZHONG_year_on_year',
               'HUANAN_year_on_year', 'XINAN_year_on_year', 'XIBEI_year_on_year',
               'firstline_year_on_year', 'secondline_year_on_year', 'thirdline_year_on_year', 'forth_year_on_year',
               'ZSJ_year_on_year', 'CSJ_year_on_year', 'HBH_year_on_year',
               'volumn_chain',
               'chain', 'east_chain', 'mid_chain', 'west_chain',
               'chain_under_90', 'chain_90_144', 'chain_above_144',
               'DONGBEI_chain', 'HUABEI_chain', 'HUADONG_chain', 'HUAZHONG_chain', 'HUANAN_chain', 'XINAN_chain',
               'XIBEI_chain',
               'firstline_chain', 'secondline_chain', 'thirdline_chain', 'forth_chain', 'ZSJ_chain', 'CSJ_chain',
               'HBH_chain']
    kwargs = {}
    for key in keylist:
        kwargs[key] = []
    # 取区域数据
    datarow = {}
    dataList = ['QUANGUO', 'east', 'mid', 'west', 'DONGBEI', 'HUABEI', 'HUADONG', 'HUAZHONG', 'HUANAN', 'XINAN', 'XIBEI'
        , 'firstline', 'secondline', 'thirdline', 'forth', 'ZSJ', 'CSJ', 'HBH']
    for i in dataList:
        datarow[i] = []
    areaList = [CityArea.QUANGUO_90] + [i for i in range(CityArea.DONGBU_90, CityArea.QUANGUO_90)]
    for i in range(0, len(dataList)):
        for y in range(2009, year):
            for m in range(1, 13):
                datarow[dataList[i]].append(
                    CalculateResult.objects.get(city_or_area=False, area=areaList[i], year=y, month=m))
        for m in range(1, month + 1):
            datarow[dataList[i]].append(
                CalculateResult.objects.get(city_or_area=False, area=areaList[i], year=year, month=m))
    # 填入全国数据
    for i in range(0, len(datarow[dataList[0]])):
        kwargs['volumn'].append(datarow['QUANGUO'][i].trade_volume)
        kwargs['under_90_volumn'].append(datarow['QUANGUO'][i].trade_volume_under_90)
        kwargs['90_144_volumn'].append(datarow['QUANGUO'][i].trade_volume_90_144)
        kwargs['above_144_volumn'].append(datarow['QUANGUO'][i].trade_volume_above_144)
        kwargs['index'].append(datarow['QUANGUO'][i].index_value_base09)
        kwargs['index_under_90'].append(datarow['QUANGUO'][i].index_value_under90_base09)
        kwargs['index_90_144'].append(datarow['QUANGUO'][i].index_value_90144_base09)
        kwargs['index_above_144'].append(datarow['QUANGUO'][i].index_value_above144_base09)
        if i > 11:
            kwargs['volumn_year_on_year'].append(datarow['QUANGUO'][i].volume_year_on_year)
            kwargs['year_on_year'].append(datarow['QUANGUO'][i].year_on_year_index)
            kwargs['year_on_year_under_90'].append(datarow['QUANGUO'][i].year_on_year_index_under90)
            kwargs['year_on_year_90_144'].append(datarow['QUANGUO'][i].year_on_year_index_90144)
            kwargs['year_on_year_above_144'].append(datarow['QUANGUO'][i].year_on_year_index_above144)
        if i > 0:
            kwargs['volumn_chain'].append(datarow['QUANGUO'][i].volume_chain)
            kwargs['chain'].append(datarow['QUANGUO'][i].chain_index)
            kwargs['chain_under_90'].append(datarow['QUANGUO'][i].chain_index_under90)
            kwargs['chain_90_144'].append(datarow['QUANGUO'][i].chain_index_90144)
            kwargs['chain_above_144'].append(datarow['QUANGUO'][i].chain_index_above144)

            # 填入地区数据
    for i in range(1, len(dataList)):
        for j in range(0, len(datarow[dataList[0]])):
            kwargs[dataList[i] + '_volumn'].append(datarow[dataList[i]][j].trade_volume)
            kwargs[dataList[i] + '_index'].append(datarow[dataList[i]][j].index_value_base09)
            if j > 11:
                kwargs[dataList[i] + '_year_on_year'].append(datarow[dataList[i]][j].year_on_year_index)
            if j > 0:
                kwargs[dataList[i] + '_chain'].append(datarow[dataList[i]][j].chain_index)

    report.IndexSummary(**kwargs)

    # 参数列表:['index','year_on_year','chain','year_on_year_plot','chain_plot']
    kwargs_2 = {}
    kwargs_2['index'] = kwargs['index']
    kwargs_2['year_on_year'] = kwargs['year_on_year']
    kwargs_2['chain'] = kwargs['chain']
    kwargs_2['year_on_year_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'yearonyearplot_90.png'
    kwargs_2['chain_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'chainplot_90.png'

    try:
        report.RadioPlot(**kwargs_2)
    except FileNotFoundError:
        return False
    # 参数列表:['url_index_plot','url_block_plot','url_area_plot', 'url_7area_plot', 'url_line_plot', 'url_s_area_plot']
    kwargs_3 = {}
    kwargs_3['url_index_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_90.png'
    kwargs_3['url_block_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_block_90.png'
    kwargs_3['url_area_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_buildarea_90.png'
    kwargs_3['url_7area_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_7area.png'
    kwargs_3['url_line_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_line.png'
    kwargs_3['url_s_area_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_s_area.png'
    report.SimplyPlot(**kwargs_3)

    kwargs_3 = {}
    kwargs_3['url_index_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_90.png'
    kwargs_3['url_block_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_block_90.png'
    kwargs_3['url_area_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_buildarea_90.png'
    kwargs_3['url_7area_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_7area.png'
    kwargs_3['url_line_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_line.png'
    kwargs_3['url_s_area_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_s_area.png'
    report.SimplyPlot(**kwargs_3)

    kwargs_4 = {}
    kwargs_4['url_index_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'volindex_90.png'
    kwargs_4['url_block_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_block_c_90.png'
    kwargs_4['url_area_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_buildarea_c_90.png'
    kwargs_4['url_7area_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_7area_c.png'
    kwargs_4['url_line_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_line_c.png'
    kwargs_4['url_s_area_plot'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_s_area_c.png'
    report.ComplexPlot(**kwargs_4)

    # list每项的参数列表:{'city_name','index_this_month','index_last_month','index_last_year',
    # 'chain_radio','year_on_year','volumn_chain','volumn_year'}

    kwargs_5 = []
    city_list = []
    cityList = City.objects.filter(ifin90=True)
    for i in cityList:
        cityindex = CalculateResult.objects.get(city_or_area=True, city=int(i.code), year=year, month=month)
        newcity = {}
        newcity['name'] = i.name
        newcity['line'] = i.line - CityArea.FIRSTLINE + 1
        newcity['city_name'] = i.name
        newcity['index_this_month'] = cityindex.index_value_base09
        newcity['index'] = cityindex.index_value_base09
        newcity['chain'] = cityindex.chain_index
        newcity['chain_radio'] = cityindex.chain_index
        newcity['year_on_year'] = cityindex.year_on_year_index
        city_list.append(newcity)
        try:
            newcity['volumn_chain'] = cityindex.trade_volume / CalculateResult.objects.get(city_or_area=True,
                                                                                           city=int(i.code),
                                                                                           year=year if month > 1 else year - 1,
                                                                                           month=month - 1 if month > 1 else 12).trade_volume - 1

        except ZeroDivisionError:
            newcity['volumn_chain'] = 0
        try:
            newcity['volumn_year'] = cityindex.trade_volume / CalculateResult.objects.get(city_or_area=True,
                                                                                          city=int(i.code),
                                                                                          year=year - 1,
                                                                                          month=month).trade_volume - 1
        except ZeroDivisionError:
            newcity['volumn_year'] = 0

        try:
            newcity['index_last_month'] = newcity['index_this_month'] / (1 + newcity['chain_radio'])
        except ZeroDivisionError:
            newcity['index_last_month'] = CalculateResult.objects.get(city_or_area=True,
                                                                      city=int(i.code),
                                                                      year=year if month > 1 else year - 1,
                                                                      month=month if month > 1 else 12).index_value_base09
        try:
            newcity['index_last_year'] = newcity['index_this_month'] / (1 + newcity['year_on_year'])
        except ZeroDivisionError:
            newcity['index_last_year'] = CalculateResult.objects.get(city_or_area=True, city=int(i.code), year=year - 1,
                                                                     month=month).index_value_base09
        kwargs_5.append(newcity)
    report.CitySummary(kwargs_5)

    # 列表中每个元素应有参数:{'city_name','chain0','chain1' ,'chain2',
    # 'chain3' 'year0' ,'year1' ,'year2','year3'}
    kwargs_6 = []
    for city in cityList:
        newcity = {}
        newcity['city_name'] = city.name
        for i in range(0, 4):
            y, m = getLastNMonth(year, month, i)
            monthData = CalculateResult.objects.get(city_or_area=True, city=int(city.code), year=y, month=m)
            newcity['chain' + str(i)] = monthData.chain_index
            newcity['year' + str(i)] = monthData.year_on_year_index
        kwargs_6.append(newcity)
    report.CitySort(month, kwargs_6)
    citynamelist = []
    citycodelist = []
    kwargs_7 = {}
    kwargs_8 = {}
    for city in cityList:
        citynamelist.append(city.name)
        citycodelist.append(int(city.code))
    for i in range(0, len(citycodelist)):
        kwargs_7[citynamelist[i]] = []
        kwargs_8[citynamelist[i]] = []
        for y in range(2009, year):
            for m in range(1, 13):
                data = CalculateResult.objects.get(city_or_area=True, city=citycodelist[i], year=y, month=m)
                kwargs_7[citynamelist[i]].append(data.index_value_base09)
                kwargs_8[citynamelist[i]].append(data.trade_volume)
        for m in range(1, month + 1):
            data = CalculateResult.objects.get(city_or_area=True, city=citycodelist[i], year=year, month=m)
            kwargs_7[citynamelist[i]].append(data.index_value_base09)
            kwargs_8[citynamelist[i]].append(data.trade_volume)

    report.CityIndex(citynamelist, **kwargs_7)
    report.CityVolumn(citynamelist, **kwargs_8)

    ploturlList = []
    try:
        for i in citycodelist:
            ploturlList.append('media/image/' + str(year) + '_' + str(month) + 'volindex_' + str(i) + '_90.png')
        report.CityPlot(citynamelist, ploturlList)
    except FileNotFoundError:
        return False

    linelist = []
    # list每个项的参数：{'index', 'chain', 'year_on_year'}
    for i in range(CityArea.FIRSTLINE, CityArea.FORTHLINE + 1):
        linedata = CalculateResult.objects.get(city_or_area=False, area=i, year=y, month=m)
        linelist.append({'index': linedata.index_value_base09, 'chain': linedata.chain_index,
                         'year_on_year': linedata.year_on_year_index})
    report.LineSummary(linelist)
    # citylist中参数: {'name', 'line', 'index', 'chain', 'year_on_year'}
    # linelist中参数: {'index', 'chain', 'year_on_year'}

    report.LineCity(city_list, linelist)

    report.EndReport()
    return True, report.url


def get_word_report_90(year: int, month: int):
    city_list = []
    city_code_list = []
    for city in City.objects.filter(ifin90=True):
        city_list.append(city.name)
        city_code_list.append(city.code)
    report = GenWordReport90('media/report/' + '90_city_report_' + str(year) + '_' + str(month) + '.docx')
    kwargs = {}
    kwargs['index_image_url_90'] = 'media/image/' + str(year) + '_' + str(month) + 'index_90.png'
    kwargs['yearonyear_image_url_90'] = 'media/image/' + str(year) + '_' + str(month) + 'yearonyearplot_90.png'
    kwargs['chain_image_url_90'] = 'media/image/' + str(year) + '_' + str(month) + 'chainplot_90.png'
    report.Firstpage(year, month, **kwargs)
    # 参数列表:['index_90', 'chain_90', 'year_on_year_90', 'city_name_90', 'city_chain_90', 'city_year_on_year_90']
    kwargs['index_90'] = []
    kwargs['chain_90'] = []
    kwargs['year_on_year_90'] = []
    for y in range(2009, year):
        for m in range(1, 13):
            total_index = CalculateResult.objects.get(city_or_area=False, area=CityArea.QUANGUO_90, year=y, month=m)
            kwargs['index_90'].append(total_index.index_value_base09)
            kwargs['chain_90'].append(total_index.chain_index)
            kwargs['year_on_year_90'].append(total_index.year_on_year_index)
    for m in range(1, month + 1):
        total_index = CalculateResult.objects.get(city_or_area=False, area=CityArea.QUANGUO_90, year=year, month=m)
        kwargs['index_90'].append(total_index.index_value_base09)
        kwargs['chain_90'].append(total_index.chain_index)
        kwargs['year_on_year_90'].append(total_index.year_on_year_index)
    kwargs['city_name_90'] = city_list
    kwargs['city_chain_90'] = []
    kwargs['city_year_on_year_90'] = []
    kwargs['city_index_90'] = []
    kwargs['city_volume_90'] = []
    kwargs['city_volume_1_90'] = []
    kwargs['city_volume_2_90'] = []

    strlist = [1, 2, 3, 4]
    for s in strlist:
        kwargs['city_index_90_' + str(s)] = []
        kwargs['city_chain_90_' + str(s)] = []
        kwargs['city_year_on_year_90_' + str(s)] = []

    kwargs['cityname_1'] = ['一线城市']
    kwargs['cityname_2'] = ['二线城市']
    kwargs['cityname_3'] = ['三线城市']
    kwargs['cityname_4'] = ['四线城市']

    for area_line in range(CityArea.FIRSTLINE, CityArea.FORTHLINE + 1):
        line = CalculateResult.objects.get(city_or_area=False, area=area_line, year=year, month=month)

        kwargs['city_index_90_' + str(area_line - CityArea.FIRSTLINE + 1)].append(line.index_value_base09)
        kwargs['city_chain_90_' + str(area_line - CityArea.FIRSTLINE + 1)].append(line.chain_index)
        kwargs['city_year_on_year_90_' + str(area_line - CityArea.FIRSTLINE + 1)].append(line.year_on_year_index)

    for city_code in city_code_list:
        city = City.objects.get(code=city_code)

        kwargs['cityname_' + str(city.line - CityArea.FIRSTLINE + 1)].append(city.name)

        city_info = CalculateResult.objects.get(city_or_area=True, city=city_code, year=year, month=month)
        kwargs['city_chain_90'].append(city_info.chain_index)
        kwargs['city_year_on_year_90'].append(city_info.year_on_year_index)
        kwargs['city_index_90'].append(city_info.index_value_base09)
        kwargs['city_volume_90'].append(city_info.trade_volume)

        kwargs['city_index_90_' + str(city.line - CityArea.FIRSTLINE + 1)].append(city_info.index_value_base09)
        kwargs['city_chain_90_' + str(city.line - CityArea.FIRSTLINE + 1)].append(city_info.chain_index)
        kwargs['city_year_on_year_90_' + str(city.line - CityArea.FIRSTLINE + 1)].append(city_info.year_on_year_index)

        last_year, last_month = getLastNMonth(year, month, 1)
        last_city_info = CalculateResult.objects.get(city_or_area=True, city=city_code, year=last_year,
                                                     month=last_month)
        kwargs['city_volume_1_90'].append(last_city_info.trade_volume)
        last_year, last_month = getLastNMonth(year, month, 2)
        last_city_info = CalculateResult.objects.get(city_or_area=True, city=city_code, year=last_year,
                                                     month=last_month)
        kwargs['city_volume_2_90'].append(last_city_info.trade_volume)
    report.Secondpage(year, month, city_list, **kwargs)
    # kwargs['city_volume_90']:当月城市“城房指数”值的列表，按城市顺序排列，'city_chain_90'以及'city_year_on_year_90'以此类推
    # kwargs['city_volume_1_90']:上月城市“城房指数”值的列表，按城市顺序排列
    # kwargs['city_volume_2_90']:两个月前城市“城房指数”值的列表，按城市顺序排列
    report.Charts(year, month, city_list, **kwargs)

    east_data = CalculateResult.objects.get(city_or_area=False, area=CityArea.DONGBU_90, year=year, month=month)
    mid_data = CalculateResult.objects.get(city_or_area=False, area=CityArea.ZHONGBU_90, year=year, month=month)
    west_data = CalculateResult.objects.get(city_or_area=False, area=CityArea.XIBU_90, year=year, month=month)
    kwargs['east_index_90'] = [east_data.index_value_base09]
    kwargs['east_chain_90'] = [east_data.chain_index]
    kwargs['east_year_on_year_90'] = [east_data.year_on_year_index]

    kwargs['mid_index_90'] = [mid_data.index_value_base09]
    kwargs['mid_chain_90'] = [mid_data.chain_index]
    kwargs['mid_year_on_year_90'] = [mid_data.year_on_year_index]

    kwargs['west_index_90'] = [west_data.index_value_base09]
    kwargs['west_chain_90'] = [west_data.chain_index]
    kwargs['west_year_on_year_90'] = [west_data.year_on_year_index]

    kwargs['index_image_url_90'] = 'media/image/' + str(year) + '_' + str(month) + 'index_90.png'
    kwargs['chain_image_url_90'] = 'media/image/' + str(year) + '_' + str(month) + 'chainplot_90.png'
    kwargs['yearonyear_image_url_90'] = 'media/image/' + str(year) + '_' + str(month) + 'yearonyearplot_90.png'
    kwargs['index_by_block_image_url_90_90'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_block_90.png'
    kwargs['index_by_buildarea_image_url_90'] = 'media/image/' + str(year) + '_' + str(
        month) + 'index_by_buildarea_90.png'

    kwargs['index_by_sevenareas_image_url_90'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_7area.png'
    kwargs['index_by_line_image_url_90'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_line.png'
    kwargs['index_by_focus_image_url_90_90'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_s_area.png'

    total_index = CalculateResult.objects.get(city_or_area=False, area=CityArea.QUANGUO_90, year=year, month=month)
    kwargs['index_under_90_90'] = [total_index.index_value_under90_base09]
    kwargs['chain_under_90_90'] = [total_index.chain_index_under90]
    kwargs['year_on_year_under_90_90'] = [total_index.year_on_year_index_under90]

    kwargs['index_90_144_90'] = [total_index.index_value_90144_base09]
    kwargs['chain_90_144_90'] = [total_index.chain_index_90144]
    kwargs['year_on_year_90_144_90'] = [total_index.year_on_year_index_90144]

    kwargs['index_above_144_90'] = [total_index.index_value_above144_base09]
    kwargs['chain_above_144_90'] = [total_index.chain_index_above144]
    kwargs['year_on_year_above_144_90'] = [total_index.year_on_year_index_above144]

    area_list = ['_northeast_90', '_north_90', '_east_90', '_mid_90', '_south_90', '_southwest_90', '_northwest_90']
    kw_list = ['index', 'year_on_year', 'chain']
    for i in range(0, len(area_list)):
        for j in kw_list:
            kwargs[j + area_list[i]] = []
        area = CalculateResult.objects.get(city_or_area=False, area=CityArea.DONGBEI + i, year=year, month=month)
        kwargs['index' + area_list[i]] = [area.index_value_base09]
        kwargs['year_on_year' + area_list[i]] = [area.year_on_year_index]
        kwargs['chain' + area_list[i]] = [area.chain_index]

    s_area_list = ['zhu', 'chang', 'bo']
    kw_list = ['_index_90', '_chain_90', '_year_on_year_90']
    for i in range(0, len(s_area_list)):
        for j in kw_list:
            kwargs[s_area_list[i] + j] = []
        area = CalculateResult.objects.get(city_or_area=False, area=CityArea.ZHUSANJIAO + i, year=year, month=month)
        kwargs[s_area_list[i] + '_index_90'] = [area.index_value_base09]
        kwargs[s_area_list[i] + '_year_on_year_90'] = [area.year_on_year_index]
        kwargs[s_area_list[i] + '_chain_90'] = [area.chain_index]

    report.Attach(year, month, city_list, **kwargs)

    report.EndReport()
    return True, report.url


def get_word_picture_90(year: int, month: int):
    report = GenWordPicture90('media/report/' + '90_city_picture_' + str(year) + '_' + str(month) + '.docx')
    city_list = []
    city_code_list = []
    kwargs = {}
    for city in City.objects.filter(ifin90=True):
        city_list.append(city.name)
        city_code_list.append(city.code)
    kwargs['city_index_url_90'] = []
    for code in city_code_list:
        kwargs['city_index_url_90'].append(
            'media/image/' + str(year) + '_' + str(month) + 'volindex_' + str(code) + '_90.png')
    kwargs['total_index_url_90'] = 'media/image/' + str(year) + '_' + str(month) + 'index_90.png'
    kwargs['total_year_on_year_url_90'] = 'media/image/' + str(year) + '_' + str(month) + 'yearonyearplot_90.png'
    kwargs['total_chain_url_90'] = 'media/image/' + str(year) + '_' + str(month) + 'chainplot_90.png'
    kwargs['index_block_90'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_block_c_90.png'
    kwargs['index_area_90'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_buildarea_c_90.png'
    report.genpicdoc(year, month, city_list, **kwargs)
    report.EndReport()
    return True, report.url


def get_word_picture_40(year: int, month: int):
    report = GenWordPicture('media/report/' + '40_city_picture_' + str(year) + '_' + str(month) + '.docx')
    city_list = []
    city_code_list = []
    kwargs = {}
    for city in City.objects.filter(ifin40=True):
        city_list.append(city.name)
        city_code_list.append(city.code)
    kwargs['city_index_url'] = []
    for code in city_code_list:
        kwargs['city_index_url'].append(
            'media/image/' + str(year) + '_' + str(month) + 'volindex_' + str(code) + '.png')
    kwargs['total_index_url'] = 'media/image/' + str(year) + '_' + str(month) + 'index.png'
    kwargs['total_year_on_year_url'] = 'media/image/' + str(year) + '_' + str(month) + 'yearonyearplot.png'
    kwargs['total_chain_url'] = 'media/image/' + str(year) + '_' + str(month) + 'chainplot.png'
    kwargs['index_block'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_block_c.png'
    kwargs['index_area'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_buildarea_c.png'
    report.genpicdoc(year, month, city_list, **kwargs)
    report.EndReport()
    return True, report.url

def get_origindata_report(year: int, month: int):
    report = GenOriginalReport('media/report/' + 'origin_data_report_' + str(year) + '_' + str(month) + '.xlsx')
    city_list = []
    city_code_list = []
    for city in City.objects.filter(ifin90=True):
        city_list.append(city.name)
        city_code_list.append(city.code)
    cityList = []
    for i in range(0, len(city_code_list)):
        newcity = {}
        newcity['code'] = city_code_list[i]
        newcity['name'] = city_list[i]
        keys = ['trade_vol', 'trade_vol_under_90', 'trade_vol_90_144', 'trade_vol_above_144', 'index', 'index_under_90', 'index_90_144', 'index_above_144']
        data_key = ['trade_volume', 'trade_volume_under_90', 'trade_volume_90_144', 'trade_volume_above_144', 'index_value', 'index_value_under90', 'index_value_90144', 'index_value_above144']
        for key in keys:
            newcity[key] = []
        citydata = list(CalculateResult.objects.filter(city_or_area=True, city=int(city_code_list[i]), year__in=range(2006, year),
                                                   month__in=range(1, 13)).order_by('year', 'month'))
        citydata.extend(list(CalculateResult.objects.filter(city_or_area=True, city=int(city_code_list[i]), year=year,
                                                        month__in=range(1, month + 1)).order_by('year', 'month')))
        for data_row in citydata:
            for i in range(0,len(keys)):
                newcity[keys[i]].append(getattr(data_row, data_key[i]))
        
        citydata =  CalculateResult.objects.get(city_or_area=True, city=city_code_list[i], year=year, month=month)
        keys = ['chain', 'volumn', 'chain_under_90', 'vol_under_90', 'chain_90_144', 'vol_90_144', 'chain_above_144', 'vol_above_144', 'max_area', 'max_price']
        data_key = ['chain_index', 'trade_volume', 'chain_index_under90', 'trade_volume_under_90', 'chain_index_90144', 'trade_volume_90_144', 'chain_index_above144', 'trade_volume_above_144', 'max_area', 'max_price']
        for i in range(0,len(keys)):
            newcity[keys[i]] = getattr(data_row, data_key[i])
        cityList.append(newcity)
    
    report.Market(cityList)
    report.Check(cityList)
    report.EndReport()
    return True, report.url
