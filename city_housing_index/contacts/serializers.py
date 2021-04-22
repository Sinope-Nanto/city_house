from rest_framework import serializers
from .models import Contact


class ContactSerializer(serializers.ModelSerializer):
    class Meta:
        model = Contact
        fields = ['id', 'name', 'mobile', 'position', 'department', 'leader', 'tel', 'qq', 'email', 'fax']
