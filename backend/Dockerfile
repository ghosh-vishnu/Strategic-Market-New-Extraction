# Use Python image
FROM python:3.12-slim

# Set working directory
WORKDIR /app

# Copy dependency file
COPY requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir --upgrade pip
RUN pip install --no-cache-dir -r requirements.txt

# Copy project files
COPY . .

# Expose the port Django will run on
EXPOSE 8000

# Default command
CMD ["gunicorn", "excel_backend.wsgi:application", "--bind", "0.0.0.0:8000"]

