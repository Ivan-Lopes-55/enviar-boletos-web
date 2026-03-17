FROM python:3.11-slim

WORKDIR /app

COPY . .

RUN pip install --no-cache-dir fastapi uvicorn pandas openpyxl python-multipart

EXPOSE 80

CMD ["uvicorn", "backend:app", "--host", "0.0.0.0", "--port", "80"]
