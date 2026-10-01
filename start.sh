#!/usr/bin/env bash
# Запуск на сервере: миграции, администратор, импорт курсов из content/, веб-сервер.
set -o errexit
python manage.py migrate --no-input
python manage.py ensure_admin
python manage.py import_content content --quiet
exec gunicorn config.wsgi --bind 0.0.0.0:${PORT:-8000} --workers ${WEB_CONCURRENCY:-3} --timeout 120
