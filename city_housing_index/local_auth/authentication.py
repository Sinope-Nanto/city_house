from django.contrib.auth.models import User
from rest_framework import authentication
from rest_framework import exceptions

from local_auth.models import UserSession
from django.utils import timezone


class CityIndexAuthentication(authentication.BaseAuthentication):
    def authenticate(self, request):
        token = request.META.get("HTTP_AUTHTOKEN")
        if not token:
            raise exceptions.NotAuthenticated("not authenticated")
        try:
            user_session = UserSession.objects.get(token=token, expire__gte=timezone.now())
        except UserSession.DoesNotExist:
            raise exceptions.AuthenticationFailed("Login Expired")
        user = user_session.user
        return (user, None)

    def authenticate_header(self, request):
        return "AUTHTOKEN"
