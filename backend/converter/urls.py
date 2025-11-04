from django.urls import path
from . import views, auth_views

urlpatterns = [
    # Authentication URLs
    path("api/auth/login/", auth_views.login_view, name="login"),
    path("api/auth/logout/", auth_views.logout_view, name="logout"),
    path("api/auth/check/", auth_views.check_auth_view, name="check_auth"),
    
    # Converter URLs
    path("api/upload/", views.upload_files, name="upload_files"),
    path("api/convert/", views.start_convert, name="start_convert"),
    path("api/progress/", views.progress, name="progress"),
    path("api/result/", views.result_file, name="result_file"),
    path("api/reset/", views.reset_job, name="reset_job"),
    path("api/upload-excel/", views.upload_excel_sheet, name="upload_excel_sheet"),
    path("api/upload-extract-excel/", views.upload_extract_excel, name="upload_extract_excel"),
    path("api/upload-direct-excel/", views.upload_direct_excel, name="upload_direct_excel"),
    path("api/apply-mapping/", views.apply_excel_mapping, name="apply_excel_mapping"),
]
