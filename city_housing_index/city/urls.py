from django.urls import path
from .views import GetCitiesView

urlpatterns = [
    path("", GetCitiesView.as_view()),
]
