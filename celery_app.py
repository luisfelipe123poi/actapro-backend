import os
from celery import Celery
from dotenv import load_dotenv

load_dotenv()

# Lee la URL de Redis desde tus variables de entorno o usa el localhost por defecto
REDIS_URL = os.getenv("REDIS_URL", "red-da6uafrncjis73f70rgg")

celery_app = Celery(
    "actapro_tasks",
    broker=REDIS_URL,
    backend=REDIS_URL
)

celery_app.conf.update(
    task_serializer="json",
    accept_content=["json"],
    result_serializer="json",
    timezone="UTC",
    enable_utc=True,
    task_track_started=True,
)