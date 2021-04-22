from django.urls import path
from .views import GetContactsView, GetContactDetailView, CreateContactView, UpdateContactView, DeleteContactView

urlpatterns = [
    path("", GetContactsView.as_view()),
    path("<int:id>/", GetContactDetailView.as_view()),
    path("create/", CreateContactView.as_view()),
    path("update/<int:id>/", UpdateContactView.as_view()),
    path("delete/<int:id>/", DeleteContactView.as_view()),
]
