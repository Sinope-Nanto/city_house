from django.urls import path
import local_admin.views.register_views as register_view
import local_admin.views.status_views as status_view
import local_admin.views.files_views as files_views
import local_admin.views.download_views as download_views

urlpatterns = [
    path("auth/register_users/", register_view.GetWaitingRegisterUsersView.as_view()),
    path("auth/register_users/<int:user_id>", register_view.GetWaitingRegisterUserDetailView.as_view()),
    path("auth/register_users/accept/<int:user_id>", register_view.AcceptRegisterUserView.as_view()),
    path("auth/register_users/refuse/<int:user_id>", register_view.RefuseRegisterUserView.as_view()),
    path("upload_status", status_view.UploadStatusViews.as_view()),
    path("upload_files", files_views.UpLoadFilesViews.as_view()),
    path("download_files/<str:filenames>", download_views.DownLoadViews.as_view()),
    path("upload_files/download/<int:fileid>", download_views.DownLoadbyIDViews.as_view()),
    path("upload_files/download_time", download_views.DownLoadbyTime.as_view()),
    path("upload_files/download_city", download_views.DownLoadbyCity.as_view()),
    path("upload_files/download_citytime", download_views.DownLoadbyCityTime.as_view())
]
