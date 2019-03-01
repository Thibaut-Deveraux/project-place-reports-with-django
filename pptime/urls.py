from django.urls import path
from . import views
from django.conf.urls.static import static
from django.conf import settings

app_name = 'pptime'
urlpatterns = [
    path('', views.index, name='index'),
    path('list/', views.list, name='list'),
]+ static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)