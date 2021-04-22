from .models import Contact
from local_auth.models import UserProfile
from .serializers import ContactSerializer


def get_contact_list(user):
    contact_query_set = Contact.objects.filter(creator=user)
    return ContactSerializer(contact_query_set, many=True).data


def check_contact_permission(user, contact):
    user_profile = UserProfile.objects.get(user_id=user.id)
    return contact.creator_id == user.id or user_profile.is_admin()


def get_contact_detail(user, contact_id):
    contact = Contact.objects.get(id=contact_id)
    if not check_contact_permission(user, contact):
        raise Contact.DoesNotExist
    return ContactSerializer(contact).data


def create_contact(user, name, mobile, position, department, leader, tel, qq, email, fax) -> (bool, str):
    if Contact.objects.filter(creator=user, mobile=mobile).exists():
        return False, "该手机号已被创建"
    contact = Contact(creator=user, name=name, mobile=mobile, position=position, department=department,
                      leader=leader, tel=tel, qq=qq, email=email, fax=fax)
    contact.save()
    return True, ""


def update_contact(user, contact_id, name, mobile, position, department, leader, tel, qq, email, fax):
    contact = Contact.objects.get(id=contact_id)
    if not check_contact_permission(user, contact):
        return False, "您没有修改该通讯录的权限"
    contact.name = name
    contact.mobile = mobile
    contact.position = position
    contact.department = department
    contact.leader = leader
    contact.tel = tel
    contact.qq = qq
    contact.email = email
    contact.fax = fax
    contact.save()
    return True, ""


def delete_contact(user, contact_id):
    contact = Contact.objects.get(id=contact_id)
    if not check_contact_permission(user, contact):
        return False, "您没有删除该通讯录的权限"
    contact.delete()
    return True, ""
