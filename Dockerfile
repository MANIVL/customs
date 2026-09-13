FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY backend ./backend
COPY frontend ./frontend
COPY main.py .

ENV APP_HOST=0.0.0.0
ENV APP_OPEN_BROWSER=0
ENV PYTHONUNBUFFERED=1

EXPOSE 8756

CMD ["sh", "-c", "uvicorn server:app --app-dir backend --host 0.0.0.0 --port ${PORT:-8756}"]
