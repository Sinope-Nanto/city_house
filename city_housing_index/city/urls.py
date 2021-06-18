from django.urls import path
from .views import GetCitiesView, GetCityBlocksView, AddCityBlockView

urlpatterns = [
    path("", GetCitiesView.as_view()),
    path("block", GetCityBlocksView.as_view()),
    path("addblock", AddCityBlockView.as_view())
]
