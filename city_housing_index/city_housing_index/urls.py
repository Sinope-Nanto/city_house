"""city_housing_index URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/2.1/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
                  path('v1/auth/', include('local_auth.urls')),
                  path('v1/cities/', include('city.urls')),
                  path('v1/contacts/', include('contacts.urls')),
                  path('v1/calculate/', include('calculate.urls')),
                  path('v1/admin/', include('local_admin.urls')),
                  path('v1/index/', include('index.urls')),
                  path('v1/init_data/', include('init_data.urls')),
              ] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
