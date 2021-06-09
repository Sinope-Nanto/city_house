from django.test import TestCase
from django.contrib.auth.models import User
from contacts.models import Contact
from local_auth.models import UserProfile

import contacts.domain as Domain
import json


# Create your tests here.
class ContactDomainTest(TestCase):

    def setUp(self):
        creator = User.objects.create_user(username="creator")
        creator_user_profile = UserProfile(user=creator, name="creator", mobile="1", role=1)
        creator_user_profile.save()

        another = User.objects.create_user(username="another")
        another_user_profile = UserProfile(user=another, name="another", mobile="2")
        another_user_profile.save()

        admin = User.objects.create_user(username="domains")
        admin_user_profile = UserProfile(user=admin, name="domains", mobile="3", role=0)
        admin_user_profile.save()

        for i in range(0, 5):
            contact = Contact(name="contact{}".format(i), mobile="1231{}".format(i), position="position",
                              creator=creator)
            contact.save()

    def test_get_contact_list_happy_path(self):
        creator = User.objects.get(username="creator")
        ret = Domain.get_contact_list(creator)
        print(ret)
        self.assertEqual(5, len(ret))

    def test_check_contact_permission_happy_path(self):
        creator = User.objects.get(username="creator")
        contact = Contact.objects.get(name="contact0")
        self.assertTrue(Domain.check_contact_permission(creator, contact))

    def test_check_contact_permission_shouldBeTrue_ifUserAdmin(self):
        admin = User.objects.get(username="domains")
        contact = Contact.objects.get(name="contact0")
        self.assertTrue(Domain.check_contact_permission(admin, contact))

    def test_check_contact_permission_shouldBeFalse_ifUserNotEqual(self):
        another = User.objects.get(username="another")
        contact = Contact.objects.get(name="contact0")
        self.assertFalse(Domain.check_contact_permission(another, contact))

    def test_get_contact_detail_happyPath(self):
        creator = User.objects.get(username="creator")
        contact = Contact.objects.get(name="contact0")
        ret = Domain.get_contact_detail(creator, contact.id)
        self.assertEqual("contact0", ret["name"])
        self.assertEqual("12310", ret["mobile"])
        self.assertEqual("position", ret["position"])

    def test_create_contact_happyPath(self):
        creator = User.objects.get(username="creator")
        ret, msg = Domain.create_contact(creator, "contactNew", "1231New", "position")
        self.assertTrue(ret)
        self.assertEqual("", msg)

    def test_create_contact_shouldBeFail_ifMobileExist(self):
        creator = User.objects.get(username="creator")
        ret, msg = Domain.create_contact(creator, "contact0", "12310", "position")
        self.assertFalse(ret)
        self.assertEqual("该手机号已被创建", msg)

    def test_update_contact_happyPath(self):
        creator = User.objects.get(username="creator")
        contact = Contact.objects.get(name="contact0")
        ret, msg = Domain.update_contact(creator, contact.id, name="contact0", mobile="newMobile",
                                         position='newPosition')
        self.assertTrue(ret)
        self.assertEqual("", msg)
        new_contact = Contact.objects.get(name="contact0")
        self.assertEqual("newMobile", new_contact.mobile)
        self.assertEqual("newPosition", new_contact.position)

    def test_update_contact_shouldFail_ifUserNoPermission(self):
        another = User.objects.get(username="another")
        contact = Contact.objects.get(name="contact0")
        ret, msg = Domain.update_contact(another, contact.id, name="contact0", mobile="newMobile",
                                         position='newPosition')
        self.assertFalse(ret)
        self.assertEqual("您没有修改该通讯录的权限", msg)
