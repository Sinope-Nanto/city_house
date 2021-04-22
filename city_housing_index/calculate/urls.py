from django.urls import path, include
import calculate.views.upload_views as upload_views
import calculate.views.model_views as model_views
from rest_framework import routers

router = routers.DefaultRouter()
router.register(r'upload_file', upload_views.DataFileUploadView)

urlpatterns = [
    path("", include(router.urls)),
    path("data_files/", upload_views.DataFileListView.as_view()),
    path("data_file/content/<int:file_id>/", upload_views.ReviewFileContentView.as_view()),
    path("data_file/delete/<int:file_id>/", upload_views.DeleteDataFileView.as_view()),

    path("model/task/execute/", model_views.ExecuteModelCalculateView.as_view()),
    path("model/tasks/", model_views.ListModelCalculateTaskView.as_view()),
    path("model/results/", model_views.ListModelCalculateResultView.as_view()),
    path("model/results/<int:result_id>/", model_views.GetModelCalculateResultDetailView.as_view())

]
