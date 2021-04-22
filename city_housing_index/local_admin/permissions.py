from rest_framework import permissions
from local_auth.models import UserProfile


class CityIndexAdminPermission(permissions.BasePermission):

    def has_permission(self, request, view):
        user = request.user
        user_profile = UserProfile.get_by_user(user)

        return user_profile.is_admin()
