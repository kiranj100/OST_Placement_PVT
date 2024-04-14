# pdfconverter/urls.py
from django.urls import path
from . import views

urlpatterns = [
    path('', views.upload_file, name='upload_file'),
    path('convert/', views.convert_to_excel, name='convert_to_excel'),
]
