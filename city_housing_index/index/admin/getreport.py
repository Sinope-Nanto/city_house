from index.plot import PlotReport
from index.models import CalculateResult
from city.models import City
from city.enums import CityArea

from index.admin.genreport import GenExcelReport
from index.admin.genreport import GenWordReport


def getLastNMonth(year,month,num):
    month -= num
    if month <= 0:
        month += 12
        year -= 1
    return year,month

def getReport(year:int,month:int):
    report = GenExcelReport('media/report/'+'40_city_report_'+str(year)+'_'+str(month)+'.xlsx')
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

    keylist = ['volumn','east_volumn','mid_volumn','west_volumn',
        'under_90_volumn','90_144_volumn','above_144_volumn','volumn_year_on_year','volumn_chain',
        'index','east_index','mid_index','west_index',
        'year_on_year','chain','east_year_on_year','east_chain',
        'mid_year_on_year','mid_chain','west_year_on_year','west_chain',
        'index_under_90','index_90_144','index_above_144','year_on_year_under_90',
        'chain_under_90','year_on_year_90_144','chain_90_144',
        'year_on_year_above_144','chain_above_144','index_CSJ']
    for key in keylist:
        kwargs[key] = []
    # 全国数据
    totalDataList = []
    eastDataList = []
    westDataList = []
    midDataList = []
    CSJDataList = [] 
    for y in range(2006,year):
        for m in range(1,13):
            totalDataList.append(CalculateResult.objects.get(city_or_area=False,area=CityArea.QUNGUO,year=y,month=m))
            eastDataList.append(CalculateResult.objects.get(city_or_area=False,area=CityArea.DONGBU,year=y,month=m))
            westDataList.append(CalculateResult.objects.get(city_or_area=False,area=CityArea.XIBU,year=y,month=m))
            midDataList.append(CalculateResult.objects.get(city_or_area=False,area=CityArea.ZHONGBU,year=y,month=m))
            CSJDataList.append(CalculateResult.objects.get(city_or_area=False,area=CityArea.CHANGSANJIAO,year=y,month=m))
    for m in range(1,month+1):
        totalDataList.append(CalculateResult.objects.get(city_or_area=False,area=CityArea.QUNGUO,year=year,month=m))
        eastDataList.append(CalculateResult.objects.get(city_or_area=False,area=CityArea.DONGBU,year=year,month=m))
        westDataList.append(CalculateResult.objects.get(city_or_area=False,area=CityArea.XIBU,year=year,month=m))
        midDataList.append(CalculateResult.objects.get(city_or_area=False,area=CityArea.ZHONGBU,year=year,month=m))
        CSJDataList.append(CalculateResult.objects.get(city_or_area=False,area=CityArea.CHANGSANJIAO,year=year,month=m))
    for i in range(0,len(totalDataList)):
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
    kwargs_2['year_on_year_plot'] = 'media/image/'+str(year)+'_'+str(month)+'yearonyearplot.png'
    kwargs_2['chain_plot'] = 'media/image/'+str(year)+'_'+str(month)+'chainplot.png'

    try:
        report.RadioPlot(**kwargs_2)
    except FileNotFoundError:
        return False
    
    # 参数列表:['url_index_plot','url_block_plot','url_line_plot']
    kwargs_3 = {}
    kwargs_3['url_index_plot'] = 'media/image/'+str(year)+'_'+str(month)+'index.png'
    kwargs_3['url_block_plot'] = 'media/image/'+str(year)+'_'+str(month)+'index_by_block.png'
    kwargs_3['url_line_plot'] = 'media/image/'+str(year)+'_'+str(month)+'index_by_buildarea.png'

    try:
        report.SimplyPlot(**kwargs_3)
    except FileNotFoundError:
        return False

    # 参数列表:['url_index_plot','url_block_plot','url_line_plot']
    kwargs_4 = {}
    kwargs_4['url_index_plot'] = 'media/image/'+str(year)+'_'+str(month)+'volindex.png'
    kwargs_4['url_block_plot'] = 'media/image/'+str(year)+'_'+str(month)+'index_by_block_c.png'
    kwargs_4['url_line_plot'] = 'media/image/'+str(year)+'_'+str(month)+'index_by_buildarea_c.png'

    try:
        report.ComplexPlot(**kwargs_4)
    except FileNotFoundError:
        return False
    
    # list每项的参数列表:{'city_name','index_this_month','index_last_month','index_last_year',
    # 'chain_radio','year_on_year','volumn_chain','volumn_year'}

    kwargs_5 = []
    cityList = City.objects.filter(ifin40=True)
    for i in cityList:
        cityindex = CalculateResult.objects.get(city_or_area=True,city=int(i.code),year=year,month=month)
        newcity = {}
        newcity['city_name'] = i.name
        newcity['index_this_month'] = cityindex.index_value
        newcity['chain_radio'] = cityindex.chain_index
        newcity['year_on_year'] = cityindex.year_on_year_index
        try:
            newcity['volumn_chain'] = cityindex.trade_volume/CalculateResult.objects.get(city_or_area=True,
                city=int(i.code),
                year=year if month > 1 else year-1,
                month=month-1 if month > 1 else 12).trade_volume - 1

        except ZeroDivisionError:
            newcity['volumn_chain'] = 0
        try:
            newcity['volumn_year'] = cityindex.trade_volume/CalculateResult.objects.get(city_or_area=True, city=int(i.code),year=year-1,month=month).trade_volume - 1
        except ZeroDivisionError:
            newcity['volumn_year'] = 0

        try:
            newcity['index_last_month'] = newcity['index_this_month']/(1+newcity['chain_radio'])
        except ZeroDivisionError:
            newcity['index_last_month'] = CalculateResult.objects.get(city_or_area=True,
            city=int(i.code),
            year=year if month > 1 else year-1,
            month=month if month > 1 else 12).index_value
        try:
            newcity['index_last_year'] = newcity['index_this_month']/(1+newcity['year_on_year'])
        except ZeroDivisionError:
            newcity['index_last_year'] = CalculateResult.objects.get(city_or_area=True,city=int(i.code),year=year-1,month=month).index_value
        kwargs_5.append(newcity)
    report.CitySummary(kwargs_5)

    # 列表中每个元素应有参数:{'city_name','chain0','chain1' ,'chain2',
    # 'chain3' 'year0' ,'year1' ,'year2','year3'}
    kwargs_6 = []
    for city in cityList:
        newcity = {}
        newcity['city_name'] = city.name
        for i in range(0,4):
            y,m = getLastNMonth(year,month,i)
            monthData = CalculateResult.objects.get(city_or_area=True,city=int(city.code),year=y,month=m)
            newcity['chain'+str(i)] = monthData.chain_index
            newcity['year'+str(i)] = monthData.year_on_year_index
        kwargs_6.append(newcity)
    report.CitySort(month,kwargs_6)

    citynamelist = []
    citycodelist = []
    kwargs_7 = {}
    kwargs_8 = {}
    for city in cityList:
        citynamelist.append(city.name)
        citycodelist.append(int(city.code))
    for i in range(0,len(citycodelist)):
        kwargs_7[citynamelist[i]] = []
        kwargs_8[citynamelist[i]] = []
        for y in range(2006,year):
            for m in range(1,13):
                data = CalculateResult.objects.get(city_or_area=True,city=citycodelist[i],year=y,month=m)
                kwargs_7[citynamelist[i]].append(data.index_value)
                kwargs_8[citynamelist[i]].append(data.trade_volume)
        for m in range(1,month+1):
            data = CalculateResult.objects.get(city_or_area=True,city=citycodelist[i],year=year,month=m)
            kwargs_7[citynamelist[i]].append(data.index_value)
            kwargs_8[citynamelist[i]].append(data.trade_volume)
    
    report.CityIndex(citynamelist,**kwargs_7)
    report.CityVolumn(citynamelist,**kwargs_8)

    ploturlList = []
    try:
        for i in citycodelist:
            ploturlList.append('media/image/'+str(year)+'_'+str(month)+'volindex_'+str(i)+'.png')
        report.CityPlot(citynamelist,ploturlList)
    except FileNotFoundError:
        return False
    report.EndReport()

    return True

def getWordReport(year:int, month:int):
    city_list = []
    city_code_list = []
    for city in City.objects.filter(ifin40=True):
        city_list.append(city.name)
        city_code_list.append(city.code)
    report = GenWordReport('media/report/'+'40_city_report_'+str(year)+'_'+str(month)+'.docx')
    # try:
    # 取全国数据

    kwargs = {}
    kwargs['index'] = []
    kwargs['chain'] = []
    kwargs['year_on_year'] = []
    for y in range(2008,year):
        for m in range(1,13):
            total_index = CalculateResult.objects.get(city_or_area=False, area=CityArea.QUNGUO, year=y, month=m)
            kwargs['index'].append(total_index.index_value)
            kwargs['chain'].append(total_index.chain_index)
            kwargs['year_on_year'].append(total_index.year_on_year_index)
    for m in range(1,month + 1):
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

        last_year,last_month = getLastNMonth(year,month,1)
        last_city_info = CalculateResult.objects.get(city_or_area=True, city=city_code, year=last_year, month=last_month)
        city_index_1.append(last_city_info.index_value)
        city_volume_1.append(last_city_info.trade_volume)

        last_year,last_month = getLastNMonth(year,month,2)
        last_city_info = CalculateResult.objects.get(city_or_area=True, city=city_code, year=last_year, month=last_month)
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
    kwargs['index_by_buildarea_image_url'] = 'media/image/' + str(year) + '_' + str(month) + 'index_by_buildarea.png'
    

    report.Firstpage(year, month,**kwargs)
    report.Secondpage(year,month,city_list,**kwargs)
    report.Charts(year,month,city_list,**kwargs)
    report.Attach(year,month,city_list,**kwargs)
        

    report.EndReport()
    # except:
    #     return False
    return True