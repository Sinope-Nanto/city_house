import datetime
import base64

from PIL import Image
from io import BytesIO

from django.core.files.base import ContentFile

from .models import UserProfile, UserSMSCode, UserSession
from .templates import LOGIN_SMS_CONTENT_TEMPLATE
from utils.sms_utils import SMSSender
from django.utils import timezone

import random
from django.contrib.auth.models import User


def check_login_mobile(mobile) -> bool:
    return UserProfile.objects.filter(mobile=mobile).exists()


def send_sms_code(mobile):
    verify_code = update_user_sms_code(mobile)
    return SMSSender(mobile).send_sms(LOGIN_SMS_CONTENT_TEMPLATE.format(verify_code))


def update_user_sms_code(mobile) -> str:
    user_profile = UserProfile.objects.get(mobile=mobile)
    verify_code = UserSMSCode.generate_verify_code()
    expire = timezone.now() + datetime.timedelta(minutes=10)

    user_sms = UserSMSCode.objects.filter(user_id=user_profile.user_id_id).first()

    if user_sms:
        if timezone.now() < user_sms.expire:
            return user_sms.code
        user_sms.code = verify_code
        user_sms.expire = expire
    else:
        user_sms = UserSMSCode(user_id=user_profile.user_id_id, code=verify_code, expire=expire)
    user_sms.save()
    return verify_code


def check_auth(mobile, verify_code) -> (bool, str):
    try:
        user_profile = UserProfile.objects.get(mobile=mobile)
    except UserProfile.DoesNotExist:
        return False, "未找到手机号"

    user_id = user_profile.user_id
    check_status = UserSMSCode.objects.filter(user_id=user_id, code=verify_code, expire__gte=timezone.now()).exists()
    msg = "" if check_status else "验证码不正确"
    return check_status, msg


def create_token(mobile):
    
    user_profile = UserProfile.objects.get(mobile=mobile)
    token = UserSession.generate_token()
  
    user_session, created = UserSession.objects.get_or_create(user_id=user_profile.user_id_id)
    user_session.token = token
    user_session.expire = timezone.now() + datetime.timedelta(minutes=60)
    user_session.save()

    return token, user_profile


def register(**kwargs) -> (bool, str):
    mobile = kwargs.get("mobile")
    name = kwargs.get("name")
    city = kwargs.get("city")
    identity = kwargs.get("identity")
    identity_image = kwargs.get("identity_image")
    new_user = User.objects.create_user(username=name, is_active=False)
    user_id = new_user.id
    image_file = decode_image(identity_image, identity)
    try:
        user_profile = UserProfile(user_id_id=user_id, name=name, mobile=mobile, city_id=city, identity=identity, identity_image=image_file)
        user_profile.set_to_wait()
        user_profile.save()
    except Exception as ex:
        return False, str(ex)
    return True, ""


def decode_image(encoded_image, file_name):
    print('1')
    decoded_bytes = base64.decodebytes(encoded_image.encode())
    image_file = ContentFile(decoded_bytes, file_name)
    return image_file
