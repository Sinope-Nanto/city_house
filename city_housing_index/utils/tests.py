from django.test import TestCase
from django.conf import settings

from utils.sms_utils import SMSSender


class SendSMSTest(TestCase):
    content = "【城房指数】这是一条测试短信, 验证码是123456"
    mobile = settings.SMS_TEST_MOBILE
    sms = None

    def setUp(self):
        self.sms = SMSSender(mobile=self.mobile)

    def send_sms_happy_path(self):
        success, msg = self.sms.send_sms(self.content)

        print(success, msg)
        self.assertTrue(success)
