#!/usr/bin/env python
import os
import sys
import django

# Add the project directory to Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Setup Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'excel_backend.settings')
django.setup()

from converter.models import CustomUser

def create_user():
    """Create the default user"""
    try:
        # Check if user already exists
        if CustomUser.objects.filter(email='abc@gmail.com').exists():
            print("User abc@gmail.com already exists!")
            return
        
        # Create user
        user = CustomUser.objects.create_user(
            username='abc_user',
            email='abc@gmail.com',
            password='abc@123',
            first_name='ABC',
            last_name='User'
        )
        
        print(f"User created successfully: {user.email}")
        print("Login credentials:")
        print("Email: abc@gmail.com")
        print("Password: abc@123")
        
    except Exception as e:
        print(f"Error creating user: {e}")

if __name__ == "__main__":
    create_user()
