from django.urls import path, re_path
from .views import home, download
from django.conf.urls.static import static
from django.conf import settings	

urlpatterns = [
    path('', home),
    path('download/<task_id>', download, name = 'download'),
] + static(settings.MEDIA_URL, document_root = settings.MEDIA_ROOT)
