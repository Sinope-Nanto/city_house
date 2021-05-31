from django.urls import path

from init_data.views import InitSystemViews

urlpatterns = [
    path("init", InitSystemViews.as_view()),
]
