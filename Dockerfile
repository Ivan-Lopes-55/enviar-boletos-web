
FROM python:3.11-slim
WORKDIR /app
COPY . .
RUN pip install fastapi uvicorn pandas openpyxl
CMD ["uvicorn", "backend:app", "--host", "0.0.0.0", "--port", "80"]
