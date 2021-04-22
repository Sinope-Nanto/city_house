import random
import uuid
from django.db import models
from django.contrib.auth.models import User

from .enum import UserRole, UserStatus

# Create your models here.
from django.db.models import SET_NULL

from city.models import City


class UserProfile(models.Model):
    user_id = models.ForeignKey(User, on_delete=SET_NULL, null=True)
    name = models.CharField(default="", max_length=100, blank=True, null=False, help_text="姓名")
    mobile = models.CharField(default="", max_length=100, blank=True, null=False, help_text="手机号", unique=True)
    city = models.ForeignKey(City, on_delete=SET_NULL, null=True)
    identity = models.CharField(default="", max_length=100, blank=True, null=False, help_text="身份证号")
    identity_image = models.ImageField(upload_to="users")
    role = models.IntegerField(default=UserRole.CUSTOMER, help_text="身份")
    status = models.IntegerField(default=UserStatus.WAIT_CHECK)

    def set_to_wait(self):
        self.status = UserStatus.WAIT_CHECK
        self.save()

    def set_to_accept(self):
        self.status = UserStatus.REGISTERED
        self.save()

    def set_to_refuse(self):
        self.status = UserStatus.REFUSED
        self.save()

    def is_admin(self):
        return self.role == UserRole.ADMIN

    @classmethod
    def get_waiting_list(cls):
        return cls.objects.filter(status=UserStatus.WAIT_CHECK)

    @classmethod
    def get_by_user(cls, user):
        return cls.objects.get(user=user)

    @classmethod
    def get_by_user_id(cls, user_id):
        return cls.objects.get(user_id=user_id)


class UserSession(models.Model):
    user = models.ForeignKey(User, on_delete=SET_NULL, null=True)
    token = models.CharField(default="", max_length=200, null=False)
    expire = models.DateTimeField(auto_now_add=True)

    @classmethod
    def generate_token(cls):
        return uuid.uuid4().hex


class UserSMSCode(models.Model):
    user = models.ForeignKey(User, on_delete=SET_NULL, null=True)
    code = models.CharField(default="", max_length=100, null=False)
    expire = models.DateTimeField(auto_now_add=False, auto_now=False)

    @classmethod
    def generate_verify_code(cls):
        return "".join([str(random.randint(0, 9)) for i in range(6)])
