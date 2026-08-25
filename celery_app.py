import os
from celery import Celery
from dotenv import load_dotenv

load_dotenv()

# Si existe la variable REDIS_URL en Render la usa; si estás en local, usa localhost por defecto
REDIS_URL = os.getenv("REDIS_URL", "redis://localhost:6379/0")

celery_app = Celery(
    "actapro_tasks",
    broker=REDIS_URL,
    backend=REDIS_URL,
    include=["tasks"]  # Registra las tareas de tasks.py en el worker
)

celery_app.conf.update(
    task_serializer="json",
    accept_content=["json"],
    result_serializer="json",
    timezone="UTC",
    enable_utc=True,
    task_track_started=True,
)
