from django.urls import path, include
from .views import SendSMSCodeView, LoginView, RegisterView
from rest_framework import routers

router = routers.DefaultRouter()
router.register(r'register', RegisterView)

urlpatterns = [
    path("", include(router.urls)),
    path("get_code/", SendSMSCodeView.as_view()),
    path("login/", LoginView.as_view()),
]
