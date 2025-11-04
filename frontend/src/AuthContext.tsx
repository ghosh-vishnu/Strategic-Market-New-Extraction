import React, { createContext, useContext, useState, useEffect } from 'react';
import type { ReactNode } from 'react';
import { API_ENDPOINTS } from './config';

interface User {
  id: number;
  email: string;
  username: string;
  first_name: string;
  last_name: string;
}

interface AuthContextType {
  user: User | null;
  isLoading: boolean;
  login: (user: User) => void;
  logout: () => Promise<void>;
  checkAuth: () => Promise<void>;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const useAuth = () => {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error('useAuth must be used within an AuthProvider');
  }
  return context;
};

interface AuthProviderProps {
  children: ReactNode;
}

// Helper: Get CSRF token from cookie
function getCookie(name: string) {
  let cookieValue: string | null = null;
  if (document.cookie && document.cookie !== '') {
    const cookies = document.cookie.split(';');
    for (let cookie of cookies) {
      cookie = cookie.trim();
      if (cookie.startsWith(name + '=')) {
        cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
        break;
      }
    }
  }
  return cookieValue;
}

// Helper: Get CSRF token from Django
async function getCSRFToken(): Promise<string> {
  try {
    // First try to get CSRF token from existing cookies
    const existingToken = getCookie('csrftoken');
    if (existingToken) {
      return existingToken;
    }
    
    // Only make backend call if no existing token
    const response = await fetch(API_ENDPOINTS.AUTH.CHECK, {
      method: 'GET',
      credentials: 'include',
    });
    
    // Extract CSRF token from response headers or cookies
    const csrfToken = getCookie('csrftoken');
    return csrfToken || '';
  } catch (error) {
    console.error('Error getting CSRF token:', error);
    return '';
  }
}

export const AuthProvider: React.FC<AuthProviderProps> = ({ children }) => {
  const [user, setUser] = useState<User | null>(() => {
    try {
      const savedUser = localStorage.getItem('user');
      return savedUser ? JSON.parse(savedUser) : null;
    } catch {
      return null;
    }
  });
  const [isLoading, setIsLoading] = useState(true);

  const login = (userData: User) => {
    setUser(userData);
    localStorage.setItem('user', JSON.stringify(userData));
  };

  const logout = async () => {
    try {
      // Server-side logout - no CSRF token needed
      const response = await fetch(API_ENDPOINTS.AUTH.LOGOUT, {
        method: 'POST',
        credentials: 'include',  // Important: include cookies
      });
      
      if (response.ok) {
        console.log('Server-side logout successful');
      }
    } catch (error) {
      console.error('Logout error:', error);
    } finally {
      // Clear user state and localStorage
      setUser(null);
      localStorage.removeItem('user');
      
      // Clear cookies from JavaScript (now possible since HttpOnly = False)
      document.cookie = 'sessionid=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/; domain=localhost;';
      document.cookie = 'csrftoken=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/; domain=localhost;';
      
      console.log('Frontend state and cookies cleared');
    }
  };

  const checkAuth = async () => {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 60000);
      const csrfToken = getCookie('csrftoken');

      const response = await fetch(API_ENDPOINTS.AUTH.CHECK, {
        method: 'GET',
        credentials: 'include',
        signal: controller.signal,
        headers: {
          'Content-Type': 'application/json',
          'Connection': 'keep-alive',
          'Cache-Control': 'no-cache',
          'X-Requested-With': 'XMLHttpRequest',
          'X-CSRFToken': csrfToken || '',
        },
      });

      clearTimeout(timeoutId);

      if (!response.ok) {
        if (response.status === 403) {
          setUser(null);
          localStorage.removeItem('user');
          return;
        }
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();

      if (data.success && data.authenticated) {
        setUser(data.user);
        localStorage.setItem('user', JSON.stringify(data.user));
      } else {
        setUser(null);
        localStorage.removeItem('user');
      }
    } catch (error) {
      console.error('Auth check error:', error);
      if (!user) {
        setUser(null);
        localStorage.removeItem('user');
      }
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    // Only check auth if user is not already in localStorage
    const savedUser = localStorage.getItem('user');
    if (!savedUser) {
      // No user in localStorage, no need to call backend
      setIsLoading(false);
    } else if (!user) {
      // User exists in localStorage but not in state, restore from localStorage
      try {
        const userData = JSON.parse(savedUser);
        setUser(userData);
        setIsLoading(false);
      } catch {
        localStorage.removeItem('user');
        setIsLoading(false);
      }
    } else {
      setIsLoading(false);
    }
    
    // Clear any existing sessionid on page load to prevent issues
    const existingSessionId = getCookie('sessionid');
    if (existingSessionId && !savedUser) {
      // Clear sessionid if no user in localStorage
      document.cookie = 'sessionid=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;';
    }
  }, []);

  const value = {
    user,
    isLoading,
    login,
    logout,
    checkAuth,
  };

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
};
