from django.urls import path
from .views import GetCitiesView, GetCityBlocksView, AddCityBlockView, DeleteBlockView, GetCityInfoByUserView

urlpatterns = [
    path("", GetCitiesView.as_view()),
    path("block", GetCityBlocksView.as_view()),
    path("addblock", AddCityBlockView.as_view()),
    path("deleteblock", DeleteBlockView.as_view()),
    path("get_by_user", GetCityInfoByUserView.as_view()),
]
