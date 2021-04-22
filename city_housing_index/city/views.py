from rest_framework.views import APIView
from .domain import *

from utils.api_response import APIResponse


# Create your views here.
class GetCitiesView(APIView):

    def get(self, request):
        return APIResponse.create_success(data=get_city_list())
