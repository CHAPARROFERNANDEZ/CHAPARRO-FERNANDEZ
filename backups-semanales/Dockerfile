FROM python:3.12-slim

# postgresql-client trae pg_dump, necesario para el backup lógico completo de la base de datos.
RUN apt-get update && \
    apt-get install -y --no-install-recommends postgresql-client && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY backup.py .

CMD ["python3", "backup.py"]
