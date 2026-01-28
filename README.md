# Strategic Market - Word to Excel Converter

A full-stack application for converting Word documents to Excel format with advanced mapping and data extraction capabilities.

## 🌟 Features

- **Word to Excel Conversion**: Convert multiple Word documents into a single Excel file
- **Direct Excel Upload**: Upload existing Excel files for processing
- **Excel Mapping**: Apply custom mapping to categorize and organize data
- **Database Storage**: All uploaded files and data stored in SQLite database
- **Auto Cleanup**: Automatic database cleanup after download
- **Thread-Safe Operations**: Thread-safe job tracking with proper locking
- **Production Ready**: Configured for both local development and production deployment
- **Automated Testing**: Pytest-based regression tests for the extractor and conversion logic

## 🚀 Tech Stack

### Frontend
- **React** with TypeScript
- **Vite** for build tooling
- **Tailwind CSS** for styling
- **Django REST Framework** API integration

### Backend
- **Django** 5.1.7
- **SQLite3** database
- **Pandas** for Excel processing
- **openpyxl** for Excel file manipulation
- **Threading** for concurrent operations

## 📋 Prerequisites

- Python 3.13+
- Node.js 18+
- npm or yarn
- pytest (installed via `requirements.txt`)

## 🔧 Installation

### Backend Setup

```bash
# Navigate to backend directory
cd backend

# Install dependencies
pip install -r requirements.txt

# Run migrations
python manage.py migrate

# Create superuser (optional)
python manage.py createsuperuser

# Start backend server
python manage.py runserver 8000
```

### Frontend Setup

```bash
# Navigate to frontend directory
cd frontend

# Install dependencies
npm install

# Start development server
npm run dev
```

## 🧪 Testing

```bash
# From the project root
cd backend

# Run the backend/unit tests
pytest
```

The repository does not ship large sample Word documents. Bring your own `.docx` inputs when running conversions locally. Frontend unit tests are not yet configured (`npm run test` currently exits with a placeholder message).

## 🔀 Switching Between Local and Production

### For Local Development

**Frontend (`frontend/src/config.ts`):**
```typescript
// Comment production line:
// export const API_BASE_URL = 'http://72.60.202.207:8000';

// Uncomment local line:
export const API_BASE_URL = 'http://127.0.0.1:8000';
```

**Backend (`backend/excel_backend/settings.py`):**
```python
# Comment production settings:
# DEBUG = False
# ALLOWED_HOSTS = ['72.60.202.207']

# Uncomment local settings:
DEBUG = True
ALLOWED_HOSTS = ['127.0.0.1', 'localhost', '0.0.0.0']
```

### For Production

**Frontend (`frontend/src/config.ts`):**
```typescript
// Uncomment production line:
export const API_BASE_URL = 'http://72.60.202.207:8000';

// Comment local line:
// export const API_BASE_URL = 'http://127.0.0.1:8000';
```

**Backend (`backend/excel_backend/settings.py`):**
```python
# Uncomment production settings:
DEBUG = False
ALLOWED_HOSTS = ['127.0.0.1', 'localhost', '0.0.0.0', '72.60.202.207']

# Comment local settings:
# DEBUG = True
# ALLOWED_HOSTS = ['127.0.0.1', 'localhost', '0.0.0.0']
```

## 🗄️ Database Models

- **UploadedFile**: Stores uploaded Word file metadata
- **JobRecord**: Tracks job progress and status
- **ExcelMapping**: Stores mapping configuration data
- **ExtractExcelData**: Stores extracted Excel data in JSON format

## 🔑 Environment Variables

### Backend
- `DJANGO_SECRET_KEY`: Django secret key
- `DEBUG`: Debug mode (True/False)
- `ALLOWED_HOSTS`: Comma-separated list of allowed hosts
- `CORS_ALLOWED_ORIGINS`: Comma-separated list of CORS origins

### Frontend
- `VITE_API_BASE_URL`: Backend API URL

## 📝 API Endpoints

- `POST /api/upload/` - Upload Word files
- `POST /api/convert/` - Start conversion
- `GET /api/progress/` - Get conversion progress
- `GET /api/result/` - Download result file
- `POST /api/upload-direct-excel/` - Upload Excel file
- `POST /api/upload-excel/` - Upload mapping Excel
- `POST /api/apply-mapping/` - Apply mapping to data
- `POST /api/reset/` - Reset job

## 🎯 Usage

1. **Upload Word Files**: Click "Browse Folder" and select Word documents
2. **Start Conversion**: Click "Start Conversion" to process files
3. **Upload Excel Mapping** (Optional): Upload mapping sheet for categorization
4. **Apply Mapping**: Click "Apply Mapping" to organize data
5. **Download Result**: Click "Download Excel File" to get final output

## 🔐 Security Features

- Thread-safe job tracking with `JOBS_LOCK`
- Filename sanitization to prevent path traversal
- Proper exception handling with logging
- Automatic database cleanup
- Production-ready security settings

## 🛠️ Project Structure

```
Strategic-Market/
├── backend/
│   ├── converter/
│   │   ├── models.py          # Database models
│   │   ├── views.py           # API endpoints
│   │   ├── urls.py            # URL routing
│   │   └── utils/
│   │       └── extractor.py   # Word extraction logic
│   └── excel_backend/
│       └── settings.py        # Django settings
├── frontend/
│   ├── src/
│   │   ├── config.ts          # API configuration
│   │   ├── AuthContext.tsx    # Authentication
│   │   └── WordToExcel.tsx    # Main component
│   └── package.json
├── test_*.py                  # Pytest suites for extraction & conversion
└── README.md
```

> Note: Legacy market report `.docx` samples have been removed from source control to keep the repository lightweight.

## 🚀 Production Deployment

- **Frontend URL**: http://72.60.202.207:3000
- **Backend URL**: http://72.60.202.207:8000
- **Current Mode**: Production

## 📄 License

This project is proprietary software for Strategic Market Research.

## 👥 Vishnu Kumar Ghosh

Developed for Strategic Market Research

## 📞 7292992274

For issues or questions, please contact the development team.

