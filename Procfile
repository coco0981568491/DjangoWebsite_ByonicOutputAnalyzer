release: python manage.py migrate --noinput
web: gunicorn website.wsgi
worker: celery --app=website.celery worker