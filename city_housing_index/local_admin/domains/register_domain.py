from local_auth.models import UserProfile
from local_auth.serializers import UserProfileSerializer
from django.contrib.auth.models import User

def get_waiting_register_list():
    user_profiles = UserProfile.get_waiting_list()
    return UserProfileSerializer(user_profiles, many=True).data


def get_waiting_register_detail(user_id):
    user_profile = UserProfile.get_by_user_id(user_id)

    ret_dict = UserProfileSerializer(user_profile).data
    ret_dict["identity"] = user_profile.identity
    ret_dict["identity_image"] = user_profile.identity_image.url

    return ret_dict


def accept_register_user(user_id) -> (bool, str):
    
    try:
        user_profile = UserProfile.get_by_user_id(user_id)
        user_profile.set_to_accept()
    except UserProfile.DoesNotExist:
        return False, "该用户不存在"
    user = User.objects.get(id = user_id)
    user.is_active = True
    user.save()
    return True, ""


def refuse_register_user(user_id) -> (bool, str):
    try:
        user_profile = UserProfile.get_by_user_id(user_id)
        user_profile.set_to_refuse()
    except UserProfile.DoesNotExist:
        return False, "该用户不存在"
    return True, ""
