from django.urls import path
from index.views.uploadindex_views import UpLoadIndexView
from index.views.uploadindex_views import AddNewMonthLine
from index.views.uploadindex_views import UpdataCityInfoView

from index.views.reportdownlond_views import DownloadReoprt
from index.views.month_report_views import GenReportViews
from index.views.uploadindex_views import CalculateCityInfoView
from index.views.uploadindex_views import FindCityInfo

urlpatterns = [
    path("upload_index", UpLoadIndexView.as_view()),
    path("add_new_month", AddNewMonthLine.as_view()),
    path("update_city_info", UpdataCityInfoView.as_view()),
    path("genreport", GenReportViews.as_view()),
    path("downloadreport", DownloadReoprt.as_view()),
    path("calculate_city_info", CalculateCityInfoView.as_view()),
    path("get_city_info",FindCityInfo.as_view()),
]
