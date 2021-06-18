from django.urls import path
from index.views.uploadindex_views import AddNewMonthColumnView
from index.views.uploadindex_views import UpdateAllCityIndexView

from index.views.report_download_views import DownloadReoprt
from index.views.month_report_views import GenReportViews, QueryReportTaskView
from index.views.uploadindex_views import CalculateCityInfoView
from index.views.uploadindex_views import GetCityIndexInfoView, ListCityIndexInfoView

urlpatterns = [
    path("add_new_month", AddNewMonthColumnView.as_view()), # ok
    path("update_city_info", UpdateAllCityIndexView.as_view()),
    path("genreport", GenReportViews.as_view()), # ok
    path("query_task", QueryReportTaskView.as_view()), # ok
    path("downloadreport", DownloadReoprt.as_view()), # ok
    path("calculate_city_info", CalculateCityInfoView.as_view()), # ok
    path("get_city_info", GetCityIndexInfoView.as_view()), # ok
    path("get_city_info_list", ListCityIndexInfoView.as_view()) # ok
]
