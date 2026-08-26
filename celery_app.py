import os
from celery import Celery
from dotenv import load_dotenv

load_dotenv()

# Si existe la variable REDIS_URL en Render la usa; si estás en local, usa localhost por defecto
REDIS_URL = os.getenv("REDIS_URL", "redis://red-da6uafrncjis73f70rgg:6379")

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
    
    # --- CONFIGURACIONES ADICIONALES PARA ALTA CONCURRENCIA Y RESILIENCIA ---
    result_expires=3600,             # Las tareas guardadas en Redis expiran en 1 hora para liberar RAM
    broker_connection_retry_on_startup=True, # Evita caídas al reiniciar la conexión con Redis en Render
    worker_prefetch_multiplier=1,    # Garantiza una distribución equitativa de tareas entre corrutinas
    task_acks_late=True,             # Si un worker se cae, la tarea no se pierde y regresa a la cola
)
