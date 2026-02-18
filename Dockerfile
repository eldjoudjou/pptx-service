FROM python:3.12-slim

WORKDIR /app

# Dépendances système : LibreOffice (conversion PPTX→PDF) + poppler (PDF→images)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress \
    poppler-utils \
    fonts-liberation \
    fonts-noto-core \
    && rm -rf /var/lib/apt/lists/*

# Dépendances Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Code du service
COPY main.py .
COPY pptx_tools.py .
COPY system_prompt.md .

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
