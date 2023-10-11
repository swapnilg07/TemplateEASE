from django.contrib import admin
from word_editor import views
from django.urls import path

urlpatterns = [
    path('admin/', admin.site.urls),
    path('parse-docx-to-html', views.parse_docx_to_html, name='parse_docx_to_html'),
    path('parse-to-docx', views.parse_to_docx, name='parse_to_docx')
]
