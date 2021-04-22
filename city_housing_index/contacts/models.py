from django.contrib.auth.models import User
from django.db import models


class Contact(models.Model):
    creator = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    name = models.CharField(default="", max_length=200, blank=True, null=False)
    mobile = models.CharField(default="", max_length=100, blank=True, null=False)
    position = models.CharField(default="", max_length=200, blank=True, null=False)
    department = models.CharField(default="", max_length=200, blank=True, null=True)
    leader = models.CharField(default="", max_length=200, blank=True, null=True)
    tel = models.CharField(default="", max_length=200, blank=True, null=True)
    qq = models.CharField(default="", max_length=200, blank=True, null=True)
    email = models.CharField(default="", max_length=200, blank=True, null=True)
    fax = models.CharField(default="", max_length=200, blank=True, null=True)
