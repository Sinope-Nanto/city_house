from .serializers import CitySerializer
from .models import City


def get_city_list():
    city_query_set = City.objects.all()
    return CitySerializer(city_query_set, many=True).data


def get_city_by_user(user):

    from local_auth.models import UserProfile
    user_profile = UserProfile.objects.get(user_id_id=user.id)
    return CitySerializer(user_profile.city).data
