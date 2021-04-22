from django.urls import path
from init_data.views import InitCityViews
from init_data.views import InitTotalViews
from init_data.views import InitAreaViews
from init_data.views import InitCityIndexViews

urlpatterns = [
    path("init_city_list",InitCityViews.as_view()),
    path("init_total",InitTotalViews.as_view()),
    path("init_area",InitAreaViews.as_view()),
    path("init_city",InitCityIndexViews.as_view())
]