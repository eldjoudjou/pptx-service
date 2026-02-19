FROM python:3.12-slim

WORKDIR /app

# Dépendances système (léger)
RUN apt-get update && apt-get install -y --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# Dépendances Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Code du service
COPY main.py .
COPY pptx_tools.py .
COPY pptx_validate.py .
COPY system_prompt.md .
COPY sia_config.md .

# Schemas XSD Office (pour validation PPTX)
COPY schemas/ ./schemas/

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
