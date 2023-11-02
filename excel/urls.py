from django.urls import path
from . import views

urlpatterns = [
    path('create_excel/<str:sheet_name>/', views.excel_view, name='create_excel'),
    path('create_daily_summary_sheet/<str:sheet_name>/',views.create_daily_summary_sheet, name='create_daily_summary_sheet'),

]

