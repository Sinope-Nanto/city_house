from calculate.domains.analyse_domain import gen_html_report
from calculate.models import DataFile
from city.models import City
from rest_framework.views import APIView
from local_auth.authentication import CityIndexAuthentication
from utils.api_response import APIResponse
from django.shortcuts import render

class GetHtmlReportView(APIView):

    def post(self, request):
        type = request.data['type']
        city = int(request.data['city'])
        year = int(request.data['year'])
        month = int(request.data['month'])
        cityinfo = City.objects.get(id=city)
        block_list = []
        for block in cityinfo.info_block['info']:
            block_list.append(block['code'])

        datafile_list = DataFile.objects.filter(city_code=request.data['city'])
        url_now = ''
        url_last = ''
        exist_now = False
        exist_last = False
        for datafile in datafile_list:
            time = datafile.start.split('-')
            if len(time) < 2:
                continue
            if int(time[0]) == year and int(time[1]) == month:
                url_now = 'media/' + str(datafile.file)
                exist_now = True
        if type == 'month':
            last_year = year - 1 if month == 1 else year
            last_month = 12 if month == 1 else month - 1
        elif type == 'year':
            last_year = year - 1
            last_month = month
        else:
            return APIResponse.create_fail(code=400, msg='bad request')
        for datafile in datafile_list:
            time = datafile.start.split('-')
            if len(time) < 2:
                continue
            if int(time[0]) == last_year and int(time[1]) == last_month:
                url_now = 'media/' + str(datafile.file)
                exist_last = True
        if not (exist_now and exist_last):
            return APIResponse.create_fail(code=404, msg="文件不存在")
        re = gen_html_report(url_now, url_last, 'media/report/' + str(year) + '_' + str(month) + '_ '+ str(cityinfo.code) + 'analysereport.html',
        block_list, block_list, 'media/init_data/reporttemplate.html', year, month)
        return render(None, re)
