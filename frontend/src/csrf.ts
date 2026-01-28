import { API_ENDPOINTS } from './config';

export function getCookie(name: string): string | null {
  if (typeof document === 'undefined') {
    return null;
  }

  const cookies = document.cookie ? document.cookie.split(';') : [];
  for (let cookie of cookies) {
    cookie = cookie.trim();
    if (cookie.startsWith(`${name}=`)) {
      return decodeURIComponent(cookie.substring(name.length + 1));
    }
  }
  return null;
}

export async function ensureCsrfToken(): Promise<string> {
  const existingToken = getCookie('csrftoken');
  if (existingToken) {
    return existingToken;
  }

  try {
    await fetch(API_ENDPOINTS.AUTH.CHECK, {
      method: 'GET',
      credentials: 'include',
    });
  } catch (error) {
    console.error('Failed to ensure CSRF token', error);
  }

  return getCookie('csrftoken') ?? '';
}

