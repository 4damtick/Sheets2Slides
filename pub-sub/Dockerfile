# Use Python 3.11 slim image as base
FROM python:3.11-slim

# Set working directory in container
WORKDIR /app

# Copy requirements first (for Docker cache efficiency)
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY main.py .

# Expose port 8080 (Cloud Run default)
EXPOSE 8080

# Cloud Run settings
ENV PORT=8080
ENV BUCKET_INPUT=YOUR_INPUT_BUCKET
ENV BUCKET_OUTPUT=YOUR_OUTPUT_BUCKET

# Run the application
CMD ["python", "main.py"]
