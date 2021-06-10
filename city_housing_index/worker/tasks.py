from .celery_app import app


@app.task
def hello(content='hello!'):
    print(content)
