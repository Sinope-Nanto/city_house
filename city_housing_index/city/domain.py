from .serializers import CitySerializer
from .models import City


def get_city_list():
    city_query_set = City.objects.all()
    return CitySerializer(city_query_set, many=True).data
