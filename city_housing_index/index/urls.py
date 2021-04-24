from django.urls import path
from index.views.uploadindex_views import UpLoadIndexView
from index.views.uploadindex_views import AddNewMonthLine
from index.views.uploadindex_views import UpdataCityInfoView
from index.views.plot import PlotViews 
from index.views.excelreport_views import ExcelReportViews
from index.views.reportdownlond_views import DownloadReoprt
from index.views.wordreport_views import WordReportViews



urlpatterns = [
    path("upload_index",UpLoadIndexView.as_view()),
    path("add_new_month",AddNewMonthLine.as_view()),
    path("update_city_info",UpdataCityInfoView.as_view()),
    path("plot",PlotViews.as_view()),
    path("genreport",ExcelReportViews.as_view()),
    path("genwordreport",WordReportViews.as_view()),
    path("downloadreport",DownloadReoprt.as_view()),
]