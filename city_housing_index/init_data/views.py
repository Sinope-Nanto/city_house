from django.shortcuts import render

from rest_framework.views import APIView
from utils.api_response import APIResponse

from .tasks import initsystem


# Create your views here.

class InitSystemViews(APIView):

    def post(self, request):
        year = int(request.data['year'])
        month = int(request.data['month'])
        initsystem.delay(year, month)

        return APIResponse.create_success()
