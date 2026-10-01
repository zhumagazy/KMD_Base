#!/usr/bin/env bash
# Сборка при деплое (Render, Railway и т.п.)
set -o errexit
pip install -r requirements.txt
python manage.py collectstatic --no-input
