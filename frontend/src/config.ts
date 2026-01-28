// API Configuration
// To switch between Production and Local:
// 1. Comment the active URL line
// 2. Uncomment the URL you want to use

// Production URL (ACTIVE - FOR PRODUCTION)
// export const API_BASE_URL = (import.meta as any).env?.VITE_API_BASE_URL || 'http://72.60.202.207:8000';

const envApiBase = (import.meta as any).env?.VITE_API_BASE_URL as string | undefined;

function normalizeLocalBaseUrl(url: string | undefined): string | undefined {
  if (!url) return url;
  try {
    const parsed = new URL(url);
    if (parsed.hostname === '127.0.0.1') {
      parsed.hostname = 'localhost';
    }
    return parsed.toString().replace(/\/$/, '');
  } catch {
    return url;
  }
}

const defaultLocalBase =
  typeof window !== 'undefined' && window.location.hostname === 'localhost'
    ? `http://localhost:8000`
    : 'http://localhost:8000';

// Local Development URL (FOR LOCAL - UNCOMMENT TO USE)
export const API_BASE_URL = normalizeLocalBaseUrl(envApiBase) || defaultLocalBase;

export const API_ENDPOINTS = {
  AUTH: {
    LOGIN: `${API_BASE_URL}/api/auth/login/`,
    LOGOUT: `${API_BASE_URL}/api/auth/logout/`,
    CHECK: `${API_BASE_URL}/api/auth/check/`,
  },
  UPLOAD: `${API_BASE_URL}/api/upload/`,
  CONVERT: `${API_BASE_URL}/api/convert/`,
  PROGRESS: `${API_BASE_URL}/api/progress/`,
  RESULT: `${API_BASE_URL}/api/result/`,
  UPLOAD_EXCEL: `${API_BASE_URL}/api/upload-excel/`,
  UPLOAD_EXTRACT_EXCEL: `${API_BASE_URL}/api/upload-extract-excel/`,
  UPLOAD_DIRECT_EXCEL: `${API_BASE_URL}/api/upload-direct-excel/`,
  APPLY_MAPPING: `${API_BASE_URL}/api/apply-mapping/`,
};

