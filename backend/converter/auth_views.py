from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods
from django.contrib.auth.models import User
from rest_framework.decorators import api_view
from rest_framework.response import Response
from rest_framework import status
import json

@api_view(['POST'])
@csrf_exempt
def login_view(request):
    """Login API endpoint"""
    try:
        data = json.loads(request.body)
        email = data.get('email')
        password = data.get('password')

     
        
        if not email or not password:
            return Response({
                'success': False,
                'message': 'Email and password are required'
            }, status=status.HTTP_400_BAD_REQUEST)
        
        # Authenticate user
        user = authenticate(request, username=email, password=password)
    
        # return Response({'user': user})
        
        if user is not None:
            login(request, user)
            return Response({
                'success': True,
                'message': 'Login successful',
                'user': {
                    'id': user.id,
                    'email': user.email,
                    'username': user.username,
                    'first_name': user.first_name,
                    'last_name': user.last_name
                }
            }, status=status.HTTP_200_OK)
        else:
            return Response({
                'success': False,
                'message': 'Invalid email or password'
            }, status=status.HTTP_401_UNAUTHORIZED)
            
    except Exception as e:
        return Response({
            'success': False,
            'message': 'Login failed',
            'error': str(e)
        }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

@api_view(['POST'])
@csrf_exempt
def logout_view(request):
    """Logout API endpoint - Server-side cookie deletion"""
    try:
        # Logout user & flush server session
        logout(request)
        request.session.flush()
        
        # Delete cookies via response (server-side only)
        response = Response({
            "success": True, 
            "message": "Logout successful"
        }, status=status.HTTP_200_OK)
        
        # Get domain from settings
        from django.conf import settings
        domain = getattr(settings, "SESSION_COOKIE_DOMAIN", None)
        
        # Delete session and CSRF cookies
        response.delete_cookie(
            settings.SESSION_COOKIE_NAME, 
            path='/', 
            domain=domain
        )
        response.delete_cookie(
            getattr(settings, "CSRF_COOKIE_NAME", "csrftoken"), 
            path='/', 
            domain=domain
        )
        
        print(f"Server-side logout completed - cookies deleted for domain: {domain}")
        
        return response
    except Exception as e:
        print(f"Logout error: {str(e)}")
        return Response({
            'success': False,
            'message': 'Logout failed',
            'error': str(e)
        }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

@api_view(['GET'])
def check_auth_view(request):
    """Check if user is authenticated"""
    try:
        # Add connection headers to prevent broken pipes
        response_data = {}
        
        if request.user.is_authenticated:
            response_data = {
                'success': True,
                'authenticated': True,
                'user': {
                    'id': request.user.id,
                    'email': request.user.email,
                    'username': request.user.username,
                    'first_name': request.user.first_name,
                    'last_name': request.user.last_name
                }
            }
        else:
            response_data = {
                'success': True,
                'authenticated': False
            }
        
        response = Response(response_data, status=status.HTTP_200_OK)
        response['Cache-Control'] = 'no-cache'
        return response
        
    except Exception as e:
        return Response({
            'success': False,
            'message': 'Auth check failed',
            'error': str(e)
        }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)
