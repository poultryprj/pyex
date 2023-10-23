from django.urls import path
from . import views

urlpatterns = [
    path('create_excel/<str:sheet_name>/', views.excel_view, name='create_excel'),
    path('read_excel/<str:sheet_name>/', views.read_excel, name='read_excel'),
    path('update_excel/<str:sheet_name>/', views.update_excel, name='update_excel'),
    path('delete_excel/<str:sheet_name>/', views.delete_excel, name='delete_excel'),
]

