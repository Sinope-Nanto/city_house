# coding: utf-8
from django.conf import settings
import datetime
import hashlib
import json

import requests


class SMSSender:
    def __init__(self, mobile):
        self.mobile = mobile

    def send_sms(self, content):
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        post_data = {
            "appId": settings.SMS_APP_ID,
            "mobiles": self.mobile,
            "content": content,
            "timestamp": timestamp,
            "sign": self.create_sign(settings.SMS_APP_ID, settings.SMS_APP_SECRET, timestamp)
        }
        print(post_data)
        response = requests.post(settings.SMS_SEND_HOST, data=post_data)
        print("response<{}>[{},{}]".format(response.status_code, response.headers, response.content))

        response_data = response.json()
        return response_data["code"] == "SUCCESS", response_data["code"]

    def create_sign(self, app_id, app_secret, timestamp):
        message = (app_id + app_secret + timestamp).encode("utf-8")
        return hashlib.md5(message).hexdigest().upper()
