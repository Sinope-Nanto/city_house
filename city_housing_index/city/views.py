from rest_framework.views import APIView
from .domain import *

from utils.api_response import APIResponse
from city.models import City
from local_auth.authentication import CityIndexAuthentication


# Create your views here.
class GetCitiesView(APIView):

    def get(self, request):
        return APIResponse.create_success(data=get_city_list())

class GetCityBlocksView(APIView):

    def post(self, request):
        if request.data.__contains__('city_code'):
            city_code = request.data['city_code']
            cityinfo = City.objects.get(code=city_code)
        else:
            city_id = int(request.data['city_id'])
            cityinfo = City.objects.get(id=city_id)
        data = cityinfo.info_block
        num = cityinfo.num_block
        re_Data = {
            'num_of_block' : num,
            'block_info' : []
        }
        for block in data['info']:
            re_Data['block_info'].append({'code':block['code'],'name':block['name'], 'remark':block['remark']})
        return APIResponse.create_success(data=re_Data)

class AddCityBlockView(APIView):

    def post(self, request):
        block_name = request.data['block']
        remark = request.data['remark']
        if request.data.__contains__('city_code'):
            city_code = request.data['city_code']
            cityinfo = City.objects.get(code=city_code)
        else:
            city_id = int(request.data['city_id'])
            cityinfo = City.objects.get(id=city_id)
        for block in cityinfo.info_block['info']:
            if block['name'] == block_name:
                return APIResponse.create_fail(code=400, msg='区块已存在')
        cityinfo.num_block += 1
        blockcode = 'BLOCK' + str(cityinfo.num_block)
        cityinfo.info_block['info'].append({'code':blockcode, 'name':block_name, 'remark':remark})
        cityinfo.save()
        return APIResponse.create_success(data={'code':blockcode,'name':block_name})

class DeleteBlockView(APIView):

    def post(self, request):
        block_name = request.data['block']
        if request.data.__contains__('city_code'):
            city_code = request.data['city_code']
            cityinfo = City.objects.get(code=city_code)
        else:
            city_id = int(request.data['city_id'])
            cityinfo = City.objects.get(id=city_id)
        for block in cityinfo.info_block['info']:
            if block['name'] == block_name:
                cityinfo.info_block['delete_block'].append(block)
                cityinfo.info_block['info'].remove(block)
                cityinfo.num_block -= 1
                cityinfo.save()
                return APIResponse.create_success(data={'code':block['code'], 'name':block['name']})
        return APIResponse.create_fail(code=404, msg='区块不存在')


class GetCityInfoByUserView(APIView):
    authentication_classes = [CityIndexAuthentication]

    def get(self, request):
        user = request.user
        return APIResponse.create_success(get_city_by_user(user))





