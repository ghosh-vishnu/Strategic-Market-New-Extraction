from django.contrib.auth.models import AbstractUser
from django.db import models
import uuid

class CustomUser(AbstractUser):
    email = models.EmailField(unique=True)
    created_at = models.DateTimeField(auto_now_add=True)
    last_login = models.DateTimeField(auto_now=True)
    
    USERNAME_FIELD = 'email'
    REQUIRED_FIELDS = ['username']
    
    def __str__(self):
        return self.email

class UploadedFile(models.Model):
    """Model to store uploaded file information in database"""
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    job_id = models.CharField(max_length=100, db_index=True)
    folder_name = models.CharField(max_length=255, default="Word_Files")
    file_path = models.CharField(max_length=500)
    file_name = models.CharField(max_length=255)
    file_size = models.BigIntegerField()
    upload_date = models.DateTimeField(auto_now_add=True)
    conversion_complete = models.BooleanField(default=False)
    download_complete = models.BooleanField(default=False)
    
    class Meta:
        db_table = 'uploaded_files'
        indexes = [
            models.Index(fields=['job_id']),
            models.Index(fields=['download_complete']),
        ]
    
    def __str__(self):
        return f"{self.folder_name}/{self.file_name}"

class JobRecord(models.Model):
    """Model to track job progress in database"""
    job_id = models.CharField(max_length=100, primary_key=True)
    folder_name = models.CharField(max_length=255, default="Word_Files")
    progress = models.IntegerField(default=0)
    status = models.CharField(max_length=50, default="pending")  # pending, converting, completed, failed
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    
    class Meta:
        db_table = 'job_records'
        indexes = [
            models.Index(fields=['is_active']),
            models.Index(fields=['created_at']),
        ]
    
    def __str__(self):
        return f"Job: {self.job_id} - {self.folder_name}"

class ExcelMapping(models.Model):
    """Model to store Excel mapping data in database"""
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    job_id = models.CharField(max_length=100, db_index=True)
    title = models.CharField(max_length=500)
    category = models.CharField(max_length=255)
    subcategory = models.CharField(max_length=255)
    subcategory_url = models.CharField(max_length=1000, blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    
    class Meta:
        db_table = 'excel_mappings'
        indexes = [
            models.Index(fields=['job_id']),
            models.Index(fields=['title']),
        ]
    
    def __str__(self):
        return f"Mapping: {self.title} (Job: {self.job_id})"

class ExtractExcelData(models.Model):
    """Model to store Extract Excel data in database"""
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    job_id = models.CharField(max_length=100, db_index=True)
    row_data = models.JSONField()  # Store entire row as JSON
    created_at = models.DateTimeField(auto_now_add=True)
    
    class Meta:
        db_table = 'extract_excel_data'
        indexes = [
            models.Index(fields=['job_id']),
        ]
    
    def __str__(self):
        return f"Extract Data (Job: {self.job_id})"