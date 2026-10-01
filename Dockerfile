FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 DEBUG=false
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN SECRET_KEY=build python manage.py collectstatic --no-input
EXPOSE 8000
CMD ["sh", "-c", "python manage.py migrate --no-input && python manage.py ensure_admin && gunicorn config.wsgi --bind 0.0.0.0:${PORT:-8000} --workers 3 --timeout 60"]
