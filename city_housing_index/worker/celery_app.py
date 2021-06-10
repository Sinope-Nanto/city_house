import os

import django
from celery import Celery
from django.conf import settings

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "city_housing_index.settings")

django.setup()

app = Celery("tasks")

app.config_from_object("worker.celery_config")
app.log.setup(
    loglevel="INFO",
    logfile=os.path.join(settings.BASE_DIR, "log/tasks.log"),
)
