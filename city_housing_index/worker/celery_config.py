# https://docs.celeryproject.org/en/latest/getting-started/first-steps-with-celery.html#celerytut-configuration

from importlib import import_module

from celery.schedules import crontab

broker_url = "redis://127.0.0.1:6379"
task_serializer = "json"
result_serializer = "json"
accept_content = ["json"]
timezone = "Asia/Chongqing"
enable_utc = True

imports = ["worker.tasks", "index.tasks"]

task_default_queue = 'city_housing_index_queue'
# task_routes = {}
# task_annotations = {}

beat_schedule = {
    'hello-every-30-seconds': {
        'task': 'worker.tasks.hello',
        'args': ('hello-every-30-seconds',),
        'schedule': 30,
    },
    'hello-by-crontab': {
        'task': 'worker.tasks.hello',
        'args': ('hello-by-crontab',),
        'schedule': crontab(minute='0'),  # 每次整点
    }
}

try:
    extra_settings = import_module("settings.worker")
    for setting_key in dir(extra_settings):
        if setting_key in locals():
            locals()[setting_key] = getattr(extra_settings, setting_key)
except ImportError:
    pass
