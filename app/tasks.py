from celery import Celery

app = Celery("bom_tasks", broker="redis://redis:6379/0")

@app.task
def dummy_task():
    return "Celery is running!"

