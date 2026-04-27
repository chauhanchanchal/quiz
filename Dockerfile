FROM python:3.11-slim

# Install system dependencies (poppler for pdf2image)
RUN apt-get update && apt-get install -y \
    poppler-utils \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# Set up app directory
WORKDIR /app

# HF Spaces requires user 1000 for write access
RUN useradd -m -u 1000 user
ENV HOME=/home/user \
    PATH=/home/user/.local/bin:$PATH

# Copy requirements first (layer cache)
COPY --chown=user requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Pre-download EasyOCR models so first request isn't 60+ seconds
# Pre-download all models so first user request isn't slow
RUN python -c "import easyocr; easyocr.Reader(['en'], gpu=False)"
RUN python -m spacy download en_core_web_sm
RUN python -c "from sentence_transformers import SentenceTransformer; SentenceTransformer('all-MiniLM-L6-v2')"

# Copy app code
COPY --chown=user . .

# Switch to non-root user
USER user

# HF Spaces expects port 7860
ENV PORT=7860
EXPOSE 7860

# Use gunicorn for production (1 worker since we hold model in memory)
CMD ["gunicorn", "--bind", "0.0.0.0:7860", "--workers", "1", "--timeout", "300", "app:app"]
