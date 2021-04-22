from django.urls import path
from .views import SendSMSCodeView, LoginView, RegisterView

urlpatterns = [
    path("get_code/", SendSMSCodeView.as_view()),
    path("login/", LoginView.as_view()),
    path("register/", RegisterView.as_view())
]
