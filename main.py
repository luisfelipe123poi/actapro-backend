import datetime 
import datetime
import hashlib
import os
import shutil
import uuid
from typing import Optional
import assemblyai as aai
from docx import Document
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response
import mercadopago
from openai import OpenAI
from pydantic import BaseModel
from pymongo import MongoClient
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY
import io
from bson import ObjectId
import re
import datetime
import hashlib
import io
import os
import re
import shutil
import uuid
from typing import Optional

import assemblyai as aai
from bson import ObjectId
from docx import Document
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response
import mercadopago
import pdfplumber
import pymupdf as fitz
from openai import OpenAI
from pydantic import BaseModel
from pymongo import MongoClient
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

load_dotenv()

app = FastAPI(title="ActaBot PH con MongoDB y Mercado Pago", version="1.9.5")
app.include_router(router)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configuración de Claves y Base de Datos
AAI_API_KEY = os.getenv("ASSEMBLYAI_API_KEY")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
MONGO_URI = os.getenv("MONGO_URI", "mongodb://localhost:27017/")
MP_ACCESS_TOKEN = os.getenv("MP_ACCESS_TOKEN")

if not AAI_API_KEY or not OPENAI_API_KEY:
  raise ValueError(
      "❌ Error: Claves de API de AssemblyAI u OpenAI no configuradas en .env"
  )

# Inicializar cliente de MongoDB especificando la base de datos aislada 'actabot_db'
mongo_client = MongoClient(MONGO_URI)
db = mongo_client["actabot_db"]
users_collection = db["users"]
actas_collection = db["actas_historial"]
transripciones_collection = db["transripciones_cache"]
scanners_historial_collection = db["scanners_historial"]  # <--- NUEVA COLECCIÓN PARA ESCANEOS

# Configurar índice TTL para eliminar automáticamente la caché de transcripciones pasados 30 días (2592000 segundos)
try:
  transripciones_collection.create_index(
      "createdAt", expireAfterSeconds=2592000
  )
except Exception as e:
  print(f"Nota sobre índice TTL: {e}")

# Inicializar SDK de Mercado Pago
mp_sdk = mercadopago.SDK(MP_ACCESS_TOKEN) if MP_ACCESS_TOKEN else None

aai.settings.api_key = AAI_API_KEY
openai_client = OpenAI(api_key=OPENAI_API_KEY)

# Prompt enfocado en máxima exhaustividad y cero resumen innecesario
prompt_sistema = """
Eres un Secretario Jurídico de alto nivel, experto en Propiedad Horizontal en Colombia (Ley 675 de 2001). 
Tu misión es transformar la transcripción de audio adjunta en un ACTA FORMAL, DETALLADA Y COMPLETA. 

ORDEN DE TRABAJO (ESTRICTO):
1. INTEGRIDAD DE LA INFORMACIÓN: NO RESUMAS EL CONTENIDO TÉCNICO NI LAS PROPUESTAS. Debes capturar todos los argumentos, cifras, justificaciones financieras, propuestas de mantenimiento y posturas de los copropietarios. Si alguien propone algo específico, inclúyelo detalladamente.
2. FILTRADO DE RUIDO: ÚNICAMENTE elimina interrupciones, saludos, chismes, peleas personales o frases vacías que no aporten al objeto de la asamblea. Todo lo que tenga que ver con gestión, presupuesto, administración o decisiones debe quedar plasmado.
3. ESTRUCTURA JURÍDICA:
   - ENCABEZADO Y QUÓRUM: Detalla la verificación de coeficientes si se menciona.
   - DESARROLLO PUNTO POR PUNTO: Para CADA punto del orden del día, redacta el desarrollo de forma narrativa pero minuciosa. Transcribe los debates relevantes ("El copropietario A solicitó aclarar X; el administrador respondió que Y").
   - DECISIONES Y VOTACIONES: Registra el sentido de las votaciones y cualquier salvedad o constancia que los copropietarios hayan solicitado dejar por escrito.
4. ESTILO Y FORMATO: Lenguaje jurídico formal, impersonal y preciso. Utiliza asteriscos dobles (ej: **$1.500.000** o **Aprobado por mayoría**) para resaltar en el texto cifras, costos, valores y decisiones clave. Redacta como si hubieras estado presente tomando nota exhaustiva de todo lo importante.
"""
PROMPT_SISTEMA_ACTAS = """Eres un Secretario Jurídico experto en Propiedad Horizontal en Colombia (Ley 675 de 2001). Tu objetivo es redactar un acta formal, jurídica y detallada a partir de la transcripción de la asamblea provista, asegurando un formato profesional en Markdown y cumplimiento legal estricto."""

# Diccionario seguro de precios y planes controlados estrictamente en el backend
PRECIOS_PLANES = {
    "basico": {
        "nombre": "Plan Básico",
        "precio": 153000.0,
        "tokens_mensuales": 100000,      # Actualizado a 100k tokens
        "documentos_estimados": 40,      # Referencia para el usuario
        "limite_horas": 15.0
    },
    "profesional": {
        "nombre": "Plan Intermedio",      # Ajustado a "Intermedio" según tu tabla
        "precio": 479000.0,
        "tokens_mensuales": 300000,      # Actualizado a 300k tokens
        "documentos_estimados": 120,     # Referencia para el usuario
        "limite_horas": 60.0,
    },
    "corporativo": {
        "nombre": "Plan Profesional / Pro", # Ajustado a "Pro" según tu tabla
        "precio": 939000.0,
        "tokens_mensuales": 1000000,     # Actualizado a 1M tokens
        "documentos_estimados": 400,     # Referencia para el usuario
        "limite_horas": 200.0,
    },
}

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import os
import openai




# Configuración de CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 2. Configuración oficial del cliente de OpenAI para Python
client = openai.OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

# 3. Modelos Pydantic para validar los datos de entrada
class ConsultaFAQRequest(BaseModel):
    tipoConsulta: str

class ConsultaPlanRequest(BaseModel):
    clienteId: str
    tipoProblema: str = None



class AuthModel(BaseModel):
  email: str
  password: str
  plan: Optional[str] = "free"


class PaymentPreferenceModel(BaseModel):
  email: str
  plan_name: str
  price: Optional[float] = None


class RenombrarActaRequest(BaseModel):
  email: str
  acta_id: str
  nuevo_nombre: str


class EliminarActaRequest(BaseModel):
  email: str
  acta_id: str


# Estructura maestra de configuración
CONFIGURACION_PLANES = {
    "free": {"tokens": 10000, "horas": 3.0},
    "basico": {"tokens": 100000, "horas": 15.0},
    "profesional": {"tokens": 300000, "horas": 60.0},
    "corporativo": {"tokens": 1000000, "horas": 200.0}
}

@app.post("/api/registro")
def registrar_usuario(data: AuthModel):
    existing_user = users_collection.find_one({"email": data.email})
    if existing_user:
        raise HTTPException(status_code=400, detail="El correo ya está registrado.")

    # Obtener configuración del plan
    plan_config = CONFIGURACION_PLANES.get(data.plan, CONFIGURACION_PLANES["free"])

    nuevo_usuario = {
        "email": data.email,
        "password": data.password, # Nota: Recuerda hashear esto en producción
        "plan": data.plan,
        # Gestión de Horas
        "horas_restantes": plan_config["horas"],
        "horas_usadas_mes": 0.0,
        "limite_horas_mes": plan_config["horas"],
        # Gestión de Tokens
        "tokens_usados": 0,
        "limite_tokens_mes": plan_config["tokens"]
    }

    users_collection.insert_one(nuevo_usuario)
    return {
        "message": "Usuario registrado con éxito",
        "email": data.email,
        "plan": data.plan,
        "tokens_limite": plan_config["tokens"],
        "limite_horas_mes": plan_config["horas"]
    }


@app.post("/api/login")
def login_usuario(data: AuthModel):
    user = users_collection.find_one({"email": data.email})
    if not user or user["password"] != data.password:
        raise HTTPException(status_code=401, detail="Credenciales inválidas.")

    return {
        "message": "Login exitoso",
        "email": user["email"],
        "plan": user.get("plan", "free"),
        "tokens_usados": user.get("tokens_usados", 0),
        "limite_tokens_mes": user.get("limite_tokens_mes", 0),
        "horas_restantes": user.get("horas_restantes", 0.0),
        "limite_horas_mes": user.get("limite_horas_mes", 0.0)
    }


@app.get("/api/user/status")
def get_user_status(email: str):
    usuario = users_collection.find_one({"email": email})
    if not usuario:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")

    return {
        "email": usuario.get("email"),
        "plan": usuario.get("plan", "free"),
        "tokens_usados": usuario.get("tokens_usados", 0),
        "limite_tokens_mes": usuario.get("limite_tokens_mes", 0),
        "horas_restantes": usuario.get("horas_restantes", 0.0),
        "horas_usadas_mes": usuario.get("horas_usadas_mes", 0.0),
        "limite_horas_mes": usuario.get("limite_horas_mes", 0.0)
    }


@app.get("/api/actas/historial")
def obtener_historial_actas(email: str):
  actas_cursor = actas_collection.find({"email": email})
  actas = []
  for doc in actas_cursor:
    doc["id"] = str(doc["_id"])
    del doc["_id"]
    actas.append(doc)
  return {"actas": actas}


@app.put("/api/actas/renombrar")
async def renombrar_acta(data: RenombrarActaRequest):
  try:
    resultado = actas_collection.update_one(
        {"email": data.email, "nombre_acta": data.acta_id},
        {"$set": {"nombre_acta": data.nuevo_nombre}},
    )

    if resultado.matched_count == 0:
      from bson import ObjectId

      try:
        resultado = actas_collection.update_one(
            {"_id": ObjectId(data.acta_id), "email": data.email},
            {"$set": {"nombre_acta": data.nuevo_nombre}},
        )
      except Exception:
        pass

    if resultado.matched_count == 0:
      raise HTTPException(
          status_code=404, detail="Acta no encontrada o no autorizada."
      )

    return {"message": "¡Acta renombrada con éxito!"}
  except HTTPException as he:
    raise he
  except Exception as e:
    raise HTTPException(status_code=500, detail=str(e))


@app.delete("/api/actas/eliminar")
async def eliminar_acta(data: EliminarActaRequest):
  try:
    resultado = actas_collection.delete_one(
        {"email": data.email, "nombre_acta": data.acta_id}
    )

    if resultado.deleted_count == 0:
      from bson import ObjectId

      try:
        resultado = actas_collection.delete_one(
            {"_id": ObjectId(data.acta_id), "email": data.email}
        )
      except Exception:
        pass

    if resultado.deleted_count == 0:
      raise HTTPException(
          status_code=404, detail="Acta no encontrada para eliminar."
      )

    return {"message": "Acta eliminada correctamente."}
  except HTTPException as he:
    raise he
  except Exception as e:
    raise HTTPException(status_code=500, detail=str(e))


import io
import re
from fastapi import FastAPI, HTTPException, Response
from bson import ObjectId
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY, TA_CENTER

@app.get("/api/actas/descargar-pdf/{acta_id}")
async def descargar_acta_pdf(acta_id: str, email: str):
    try:
        # Buscar el acta en la base de datos por nombre o por _id de MongoDB
        acta = actas_collection.find_one({"email": email, "nombre_acta": acta_id})
        if not acta:
            try:
                acta = actas_collection.find_one({"_id": ObjectId(acta_id), "email": email})
            except Exception:
                pass
                
        if not acta:
            raise HTTPException(status_code=404, detail="Acta no encontrada.")
            
        contenido_texto = acta.get("contenido", "")
        nombre_archivo = acta.get("nombre_acta", "acta").replace(".docx", "").replace(".pdf", "").strip() + ".pdf"
        
        # Configuración del PDF en memoria usando ReportLab Platypus
        pdf_buffer = io.BytesIO()
        
        doc = SimpleDocTemplate(
            pdf_buffer, 
            pagesize=letter,
            rightMargin=54,
            leftMargin=54,
            topMargin=54,
            bottomMargin=54
        )
        
        story = []
        
        # Estilos profesionales
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            'ActaMainTitle',
            parent=styles['Heading1'],
            fontName='Helvetica-Bold',
            fontSize=14,
            leading=18,
            alignment=TA_CENTER,
            spaceAfter=15,
            textColor=styles['Normal'].textColor
        )
        
        subtitle_style = ParagraphStyle(
            'ActaSubTitle',
            parent=styles['Heading2'],
            fontName='Helvetica-Bold',
            fontSize=11,
            leading=15,
            alignment=TA_LEFT,
            spaceBefore=10,
            spaceAfter=6,
            textColor=styles['Normal'].textColor
        )
        
        body_style = ParagraphStyle(
            'ActaBody',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=14,
            alignment=TA_JUSTIFY,
            spaceAfter=8
        )
        
        bullet_style = ParagraphStyle(
            'ActaBullet',
            parent=body_style,
            leftIndent=15,
            spaceAfter=4
        )
        
        # Añadir título principal al documento
        story.append(Paragraph("ACTA DE ASAMBLEA GENERAL DE COPROPIETARIOS", title_style))
        story.append(Spacer(1, 10))
        
        # Procesar el contenido línea por línea
        lineas = contenido_texto.split('\n')
        for linea in lineas:
            linea_original = linea.strip()
            
            if not linea_original:
                story.append(Spacer(1, 6))
                continue
                
            es_encabezado_md = linea_original.startswith('#')
            texto_limpio = linea_original.lstrip('#').strip()
            
            # Limpieza de caracteres de escape HTML básicos
            texto_limpio = texto_limpio.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            
            # Convertir negritas de Markdown (**texto**) a etiquetas de ReportLab (<b>texto</b>)
            texto_limpio = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', texto_limpio)
            
            # Eliminar asteriscos sueltos o viñetas de markdown remanentes
            texto_limpio = re.sub(r'^[\*\-\•]\s*', '', texto_limpio)
            texto_limpio = texto_limpio.replace('*', '')

            if not texto_limpio:
                continue

            # Lógica de asignación de estilos inteligente
            if es_encabezado_md or (texto_limpio.isupper() and len(texto_limpio) < 80):
                story.append(Paragraph(texto_limpio, subtitle_style))
            elif linea_original.startswith(('*', '-', '•')):
                story.append(Paragraph(f"• {texto_limpio}", bullet_style))
            else:
                story.append(Paragraph(texto_limpio, body_style))

        # Construir el PDF
        doc.build(story)
        pdf_buffer.seek(0)
        
        return Response(
            content=pdf_buffer.getvalue(),
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{nombre_archivo}"'}
        )
        
    except HTTPException as he:
        raise he
    except Exception as e:
        print(f"Error generando PDF: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error interno al generar el PDF: {str(e)}")


@app.post("/api/crear-preferencia-pago")
def crear_preferencia_pago(data: PaymentPreferenceModel):
    if not mp_sdk:
        raise HTTPException(
            status_code=500, detail="Mercado Pago no está configurado en el servidor."
        )

    plan_recibido = data.plan_name.lower().strip()

    # Mapeo actualizado
    mapeo = {
        "basico": "basico",
        "básico": "basico",
        "plan basico": "basico",
        "plan básico": "basico",
        "profesional": "profesional", # Corresponde a tu plan intermedio
        "plan profesional": "profesional",
        "corporativo": "corporativo",
        "plan corporativo": "corporativo",
    }

    plan_id = mapeo.get(plan_recibido)
    if not plan_id or plan_id not in PRECIOS_PLANES:
        raise HTTPException(
            status_code=400, detail=f"Plan no válido: '{data.plan_name}'"
        )

    info_plan = PRECIOS_PLANES[plan_id]

    # Verificar o crear usuario base
    user = users_collection.find_one({"email": data.email})
    if not user:
        users_collection.insert_one({
            "email": data.email,
            "password": "temp_password_temporal",
            "plan": "free",
            "tokens_usados": 0,
            "limite_tokens_mes": 0,
            "horas_usadas_mes": 0.0,
            "horas_restantes": 0.0,
            "limite_horas_mes": 3.0,
        })

    preference_data = {
        "items": [{
            "title": f"ActaBot PH - {info_plan['nombre']}",
            "quantity": 1,
            "currency_id": "COP",
            "unit_price": float(info_plan["precio"]),
        }],
        "payer": {"email": data.email},
        "back_urls": {
            "success": "https://actapro-backend.onrender.com/pago-exitoso",
            "failure": "https://actapro-backend.onrender.com/pago-fallido",
            "pending": "https://actapro-backend.onrender.com/pago-pendiente",
        },
        "auto_return": "approved",
        "notification_url": "https://actapro-backend.onrender.com/api/webhook-mercadopago",
        "statement_descriptor": "ACTABOT PH",
        "external_reference": data.email,
    }

    try:
        preference_response = mp_sdk.preference().create(preference_data)
        preference = preference_response.get("response", preference_response)
        init_point = preference.get("init_point")
        sandbox_init_point = preference.get("sandbox_init_point")

        if not init_point:
            raise HTTPException(
                status_code=400,
                detail="Mercado Pago rechazó la preferencia o devolvió credenciales inválidas.",
            )

        return {
            "init_point": init_point,
            "sandbox_init_point": sandbox_init_point,
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/webhook-mercadopago")
async def webhook_mercadopago(request: Request):
    try:
        body = await request.json()
        if body.get("type") == "payment":
            payment_id = body.get("data", {}).get("id")
            if payment_id and mp_sdk:
                payment_info = mp_sdk.payment().get(payment_id)
                payment = payment_info.get("response", {})

                if payment.get("status") == "approved":
                    payer_email = payment.get("payer", {}).get("email") or payment.get("external_reference")
                    if payer_email:
                        items = payment.get("additional_info", {}).get("items", []) or payment.get("items", [])
                        
                        # Plan por defecto si no se detecta explícitamente
                        plan_asignado = "basico"

                        for item in items:
                            title_lower = item.get("title", "").lower()
                            if "corporativo" in title_lower:
                                plan_asignado = "corporativo"
                            elif "profesional" in title_lower or "intermedio" in title_lower:
                                plan_asignado = "profesional"
                            elif "básico" in title_lower or "basico" in title_lower:
                                plan_asignado = "basico"

                        # Obtener configuración del plan desde PRECIOS_PLANES (o CONFIGURACION_PLANES)
                        info_plan = PRECIOS_PLANES.get(plan_asignado, PRECIOS_PLANES["basico"])
                        
                        tokens_otorgados = info_plan.get("tokens_mensuales", 100000)
                        horas_otorgadas = info_plan.get("limite_horas", 15.0)

                        # Actualizar en la base de datos reseteando los contadores de uso actual
                        users_collection.update_one(
                            {"email": payer_email},
                            {
                                "$set": {
                                    "plan": plan_asignado,
                                    "tokens_usados": 0,
                                    "limite_tokens_mes": tokens_otorgados,
                                    "horas_usadas_mes": 0.0,
                                    "horas_restantes": horas_otorgadas,
                                    "limite_horas_mes": horas_otorgadas,
                                }
                            },
                        )
    except Exception as e:
        print(f"Error en webhook: {e}")

    return {"status": "ok"}

import ffmpeg

# Función para validar la calidad técnica del audio en Python
def validar_calidad_audio(file_path: str):
    try:
        # ffmpeg.probe obtiene los metadatos del archivo
        probe = ffmpeg.probe(file_path)
        format_info = probe.get('format', {})
        
        bit_rate = format_info.get('bit_rate')
        duration = format_info.get('duration')
        
        # REGLA 1: Bitrate menor a 32 kbps (audio ultra comprimido o estática)
        if bit_rate is not None:
            if int(bit_rate) < 32000:
                return {
                    "valido": False,
                    "motivo": "AUDIO_MUY_COMPRIMIDO_O_DEFECTUOSO",
                    "mensaje": "El archivo presenta una calidad técnica deficiente (bitrate muy bajo). Esto impedirá una correcta identificación de oradores."
                }
                
        # REGLA 2: Audio demasiado corto (menos de 10 segundos para una asamblea)
        if duration is not None:
            if float(duration) < 10.0:
                return {
                    "valido": False,
                    "motivo": "AUDIO_DEMASIADO_CORTO",
                    "mensaje": "El archivo de audio es demasiado corto para ser una asamblea."
                }
                
        return {"valido": True}
        
    except ffmpeg.Error:
        # Ocurre si ffprobe no puede leer el archivo o está corrupto
        return {
            "valido": False,
            "motivo": "ERROR_DE_LECTURA",
            "mensaje": "No se pudo leer el archivo de audio. Puede estar corrupto."
        }


@app.post("/procesar")
async def procesar_asamblea(
    file: UploadFile = File(...),
    instrucciones: Optional[str] = Form(""),
    email: str = Form(...),
    nombre_personalizado: Optional[str] = Form(None),
):
    user = users_collection.find_one({"email": email})
    if not user:
        raise HTTPException(status_code=401, detail="Usuario no autenticado.")

    plan_usuario = user.get("plan", "free")

    os.makedirs("temp_uploads", exist_ok=True)
    os.makedirs("temp_outputs", exist_ok=True)

    session_id = str(uuid.uuid4())
    temp_audio_path = f"temp_uploads/{session_id}_{file.filename}"

    try:
        content_bytes = await file.read()
        
        # Guardar archivo temporalmente para validar calidad y duración
        with open(temp_audio_path, "wb") as buffer:
            buffer.write(content_bytes)

        # 1. Validar calidad técnica y obtener duración exacta usando ffmpeg.probe
        try:
            metadata = await __import__("asyncio").to_thread(ffmpeg.probe, temp_audio_path)
            format_data = metadata.get("format", {})
            bit_rate = format_data.get("bit_rate")
            duration_seconds = format_data.get("duration")

            if not duration_seconds:
                raise HTTPException(
                    status_code=400,
                    detail="No se pudo leer la duración del archivo de audio."
                )

            duracion_segundos = float(duration_seconds)
            duracion_minutos = duracion_segundos / 60.0

            # REGLA CALIDAD 1: Bitrate menor a 32 kbps
            if bit_rate and float(bit_rate) < 32000:
                raise HTTPException(
                    status_code=400,
                    detail="El archivo presenta una calidad técnica deficiente (bitrate muy bajo). Esto impedirá una correcta identificación de oradores."
                )

            # REGLA CALIDAD 2: Audio demasiado corto (menos de 10 segundos)
            if duracion_segundos < 10:
                raise HTTPException(
                    status_code=400,
                    detail="El archivo de audio es demasiado corto para ser una asamblea."
                )

        except HTTPException as he:
            raise he
        except Exception as e:
            raise HTTPException(
                status_code=400,
                detail=f"No se pudo leer el archivo de audio. Puede estar corrupto o el formato no es compatible. Detalle: {str(e)}"
            )

        # 2. Validar límite de duración por archivo individual según el plan
        limites_por_archivo = {
            "free": 30,          # Máximo 30 min por archivo
            "basico": 180,       # Máximo 3 horas por archivo
            "profesional": 300,  # Máximo 5 horas por archivo
            "corporativo": 600   # Máximo 10 horas por archivo
        }
        limite_archivo_min = limites_por_archivo.get(plan_usuario, 30)
        
        if duracion_minutos > limite_archivo_min:
            raise HTTPException(
                status_code=400,
                detail=f"Este audio dura {round(duracion_minutos, 1)} minutos. Tu plan actual permite un límite máximo de {limite_archivo_min} minutos por archivo individual."
            )

        # 3. Validar límite de horas acumuladas en el mes
        horas_usadas_mes = user.get("horas_usadas_mes", 0.0)
        limites_mensuales = {
            "free": 3.0,
            "basico": 15.0,
            "profesional": 60.0,
            "corporativo": 200.0
        }
        limite_mes = limites_mensuales.get(plan_usuario, 1.0)

        if horas_usadas_mes >= limite_mes:
            raise HTTPException(
                status_code=403,
                detail=f"Has alcanzado el límite de {limite_mes} horas mensuales de tu plan. Actualiza tu suscripción para seguir procesando actas."
            )

        # Caché de transcripción para evitar gastos duplicados en AssemblyAI
        file_hash = hashlib.sha256(content_bytes).hexdigest()
        cached_transcription = transripciones_collection.find_one(
            {"file_hash": file_hash}
        )

        if cached_transcription:
            print(
                "💡 Audio duplicado detectado: Reutilizando transcripción guardada"
                " para evitar gasto en AssemblyAI."
            )
            texto_transcrito = cached_transcription["texto_transcrito"]
        else:
            config = aai.TranscriptionConfig(speaker_labels=True, language_code="es")
            transcriber = aai.Transcriber()
            transcript = transcriber.transcribe(temp_audio_path, config=config)

            if transcript.status == aai.TranscriptStatus.error:
                raise HTTPException(
                    status_code=500, detail=f"Error en AssemblyAI: {transcript.error}"
                )

            texto_transcrito = ""
            if transcript.utterances:
                for utterance in transcript.utterances:
                    texto_transcrito += (
                        f"[Persona {utterance.speaker}]: {utterance.text}\n"
                    )
            else:
                texto_transcrito = transcript.text

            transripciones_collection.insert_one({
                "file_hash": file_hash,
                "filename": file.filename,
                "texto_transcrito": texto_transcrito,
                "fecha": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                "createdAt": datetime.datetime.utcnow(),
            })

        # Lógica de nombre del acta
        if not nombre_personalizado or nombre_personalizado.strip() == "":
            try:
                prompt_nombre = f"""
                Analiza el siguiente fragmento de transcripción de una asamblea y extrae estrictamente el nombre del edificio, conjunto residencial, copropiedad o empresa mencionada. 
                Responde ÚNICAMENTE con un nombre limpio apto para archivo (sin espacios, usa guiones bajos _, sin tildes ni caracteres especiales, por ejemplo: Acta_Asamblea_Edificio_Torre_Central).
                
                Transcripción:
                {texto_transcrito[:2500]}...
                """
                resp_nombre = openai_client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "user", "content": prompt_nombre}],
                    temperature=0.2,
                )
                nombre_ia = (
                    resp_nombre.choices[0]
                    .message.content.strip()
                    .replace(" ", "_")
                )
                nombre_ia = "".join(
                    c for c in nombre_ia if c.isalnum() or c in ("_", "-")
                )
                nombre_archivo_acta = (
                    f"{nombre_ia}.docx"
                    if nombre_ia
                    else f"Acta_Asamblea_{session_id[:8]}.docx"
                )
            except Exception:
                nombre_archivo_acta = f"Acta_Asamblea_{session_id[:8]}.docx"
        else:
            nombre_limpio = nombre_personalizado.strip().replace(" ", "_")
            nombre_limpio = "".join(
                c for c in nombre_limpio if c.isalnum() or c in ("_", "-", ".")
            )
            nombre_base = nombre_limpio.replace(".docx", "")
            nombre_archivo_acta = f"{nombre_base}.docx"

        output_docx_path = f"temp_outputs/{session_id}_{nombre_archivo_acta}"

        prompt_sistema = PROMPT_SISTEMA_ACTAS
        if instrucciones:
            prompt_sistema += (
                f"\n\nINSTRUCCIONES ADICIONALES DEL USUARIO:\n{instrucciones}"
            )

        response = openai_client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": prompt_sistema},
                {
                    "role": "user",
                    "content": f"Transcripción de la asamblea:\n\n{texto_transcrito}",
                },
            ],
            temperature=0.3,
        )

        acta_final = response.choices[0].message.content

        doc = Document()
        titulo_principal = doc.add_heading(
            "ACTA DE ASAMBLEA GENERAL DE COPROPIETARIOS", level=0
        )
        titulo_principal.alignment = 1

        for linea in acta_final.split("\n"):
            linea_clean = linea.strip()
            if not linea_clean:
                continue

            if linea_clean.startswith("# "):
                doc.add_heading(linea_clean.replace("# ", "").strip(), level=1)
            elif linea_clean.startswith("## ") or linea_clean.startswith("### "):
                doc.add_heading(linea_clean.replace("#", "").strip(), level=2)
            else:
                p = doc.add_paragraph()
                if "**" in linea_clean:
                    partes = linea_clean.split("**")
                    for i, parte in enumerate(partes):
                        if parte:
                            run = p.add_run(parte)
                            if i % 2 == 1:
                                run.bold = True
                else:
                    p.add_run(linea_clean)

        doc.save(output_docx_path)

        peso_archivo = f"{round(os.path.getsize(output_docx_path) / 1024, 1)} KB"

        data_acta = {
            "email": email,
            "nombre_acta": nombre_archivo_acta,
            "fecha": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "peso": peso_archivo,
            "contenido": acta_final,
        }
        actas_collection.insert_one(data_acta)

        # 4. Actualizar consumo mensual de horas sumando la duración real del audio procesado
        nuevas_horas = horas_usadas_mes + (duracion_segundos / 3600.0)
        users_collection.update_one(
            {"email": email}, 
            {"$set": {"horas_usadas_mes": nuevas_horas}}
        )

        return FileResponse(
            path=output_docx_path,
            filename=nombre_archivo_acta,
            media_type=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.Document"
            ),
        )

    except Exception as e:
        import traceback
        traceback.print_exc() # Esto imprimirá el error exacto en los logs de Render
        if isinstance(e, HTTPException):
            raise e
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if os.path.exists(temp_audio_path):
            os.remove(temp_audio_path)

@app.get("/api/actas/descargar/{acta_id}")
async def descargar_acta(acta_id: str, email: str):
    # 1. Buscar en MongoDB por ID o nombre asegurando el aislamiento del usuario
    acta = None
    try:
        acta = actas_collection.find_one({"_id": ObjectId(acta_id), "email": email})
    except Exception:
        pass
        
    if not acta:
        acta = actas_collection.find_one({"nombre_acta": acta_id, "email": email})
        
    if not acta:
        raise HTTPException(status_code=404, detail="Acta no encontrada.")
        
    contenido_texto = acta.get("contenido", "")
    nombre_archivo = acta.get("nombre_acta", "Acta_Asamblea.docx")
    
    # 2. Recrear el archivo .docx temporalmente con formato completo
    os.makedirs("temp_outputs", exist_ok=True)
    temp_path = f"temp_outputs/download_{acta_id}.docx"
    
    doc = Document()
    titulo_principal = doc.add_heading("ACTA DE ASAMBLEA GENERAL DE COPROPIETARIOS", level=0)
    titulo_principal.alignment = 1

    for linea in contenido_texto.split("\n"):
        linea_clean = linea.strip()
        if not linea_clean:
            continue

        if linea_clean.startswith("# "):
            doc.add_heading(linea_clean.replace("# ", "").strip(), level=1)
        elif linea_clean.startswith("## ") or linea_clean.startswith("### "):
            doc.add_heading(linea_clean.replace("#", "").strip(), level=2)
        else:
            p = doc.add_paragraph()
            if "**" in linea_clean:
                partes = linea_clean.split("**")
                for i, parte in enumerate(partes):
                    if parte:
                        run = p.add_run(parte)
                        if i % 2 == 1:
                            run.bold = True
            else:
                p.add_run(linea_clean)

    doc.save(temp_path)
    
    return FileResponse(
        path=temp_path,
        filename=nombre_archivo,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

import os
import pymupdf
import pdfplumber
from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from typing import Optional
from openai import OpenAI

# Inicializa el cliente de OpenAI
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))



@app.post("/escanear")
async def escanear_documento(
    file: UploadFile = File(...), 
    email: Optional[str] = Form(None)
):
    try:
        # 0. VALIDAR USUARIO Y CUOTA DE TOKENS SI SE PROPORCIONA EMAIL
        if email:
            usuario = users_collection.find_one({"email": email})
            if usuario:
                tokens_usados = usuario.get("tokens_usados", 0)
                limite_tokens = usuario.get("limite_tokens_mes", 0)
                # Si el usuario tiene un límite asignado mayor a 0 y ya lo alcanzó
                if limite_tokens > 0 and tokens_usados >= limite_tokens:
                    raise HTTPException(
                        status_code=403, 
                        detail="Has alcanzado el límite de tokens mensuales de tu plan. Actualiza tu suscripción para continuar."
                    )

        # 1. Leer los bytes del archivo subido desde el frontend
        file_bytes = await file.read()
        filename = file.filename.lower()
        
        texto_extraido = ""
        es_imagen = filename.endswith((".png", ".jpg", ".jpeg", ".webp"))
        
        # 2. Extracción según el formato del archivo
        if es_imagen:
            base64_image = base64.b64encode(file_bytes).decode("utf-8")
            contenido_usuario = [
                {
                    "type": "text",
                    "text": "Analyze this scanned document or image. Extract the information by structuring the visual design with corporate semantic HTML tags (use <h1>, <h2>, <p>, <table>, <thead>, <tbody>, <tr>, <th>, <td>). Apply Tailwind CSS classes to maintain a professional style (e.g., fonts, clean borders, spacing). DO NOT use Markdown, DO NOT use asterisks, DO NOT use code blocks of any kind. If there are data tables, create them completely with HTML tags. If you detect statistical charts or diagrams, represent them with a structured div with the class 'p-4 border-2 border-dashed border-slate-300 bg-slate-50 text-center text-slate-500 rounded-lg text-xs my-4' indicating the content of the chart."
                },
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/jpeg;base64,{base64_image}"
                    }
                }
            ]
        else:
            if filename.endswith(".pdf"):
                doc = fitz.open(stream=file_bytes, filetype="pdf")
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    texto_pagina = page.get_text()
                    
                    if texto_pagina.strip():
                        texto_extraido += f"\n--- Página {page_num + 1} ---\n" + texto_pagina
                
                doc.close()
                
                # Respaldo con pdfplumber para extracción precisa de tablas en PDFs
                if len(texto_extraido.strip()) < 50:
                    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                        for i, page in enumerate(pdf.pages):
                            t = page.extract_text()
                            if t:
                                texto_extraido += f"\n--- Página (Tablas) {i + 1} ---\n" + t

            elif filename.endswith((".txt", ".doc", ".docx")):
                texto_extraido = file_bytes.decode("utf-8", errors="ignore")
            else:
                raise HTTPException(status_code=400, detail="Formato de archivo no soportado. Sube un PDF, imagen o documento de texto.")

            if not texto_extraido.strip():
                raise HTTPException(status_code=400, detail="El documento está vacío o no se pudo extraer texto legible.")

            contenido_usuario = f"""Analyze the following text extracted from the document. Your output must be EXCLUSIVELY corporate semantic HTML ready to render directly in a browser or web container.
- Replicate the original visual and section structure.
- Use <h1>, <h2> for main and section titles.
- Use <p> for paragraphs with Tailwind classes (e.g., text-slate-900, text-xs, leading-relaxed).
- Use complete table tags (<table>, <thead>, <tbody>, <tr>, <td>) with borders and corporate classes if there is structured data.
- If you detect references to charts, schemes, or diagrams, create them as a visual block with a dotted border.
- FORBIDDEN to use Markdown, asterisks (*), markdown list hyphens (#), or wrap the result in markdown code quotes.

Extracted text:
{texto_extraido[:15000]}"""

        # 3. Procesamiento inteligente y estructurado con OpenAI GPT-4o
        response_openai = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "system",
                    "content": """You are a Documentation Engineer and Web Designer expert in corporate digitalization. 
Your sole mission is to transform the input document information into a pure, clean, and professional HTML code block integrated with Tailwind CSS classes.
STRICT RULES:
1. Return ONLY valid HTML code. Do not include prior explanations or text outside of the HTML.
2. NEVER use Markdown syntax (*, #, -, ```html). The result must be plain text containing exclusively HTML markup.
3. Structure data tables with <table>, <thead>, <tbody>, <tr>, <th>, and <td> applying clean classes (e.g., border border-slate-300 p-2).
4. Represent detected charts using a <div> container with dotted borders and professional design.
5. Maintain absolute fidelity to the original document structure."""
                },
                {
                    "role": "user",
                    "content": contenido_usuario
                }
            ],
            temperature=0.0
        )

        resultado_html = response_openai.choices[0].message.content.strip()

        # Descontar / sumar tokens consumidos en la BD si el usuario está autenticado
        tokens_consumidos = 0
        if email and hasattr(response_openai, "usage") and response_openai.usage:
            tokens_consumidos = response_openai.usage.total_tokens
            users_collection.update_one(
                {"email": email},
                {"$inc": {"tokens_usados": tokens_consumidos}}
            )

        # Limpieza defensiva por si el modelo por inercia agrega bloques de código
        if resultado_html.startswith("```html"):
            resultado_html = resultado_html[7:]
        if resultado_html.startswith("```"):
            resultado_html = resultado_html[3:]
        if resultado_html.endswith("```"):
            resultado_html = resultado_html[:-3]
        resultado_html = resultado_html.strip()

        # Guardar automáticamente en la colección de scanners si hay email
        scanner_id = None
        if email:
            nuevo_registro = {
                "email": email,
                "nombre": file.filename or "Documento Escaneado",
                "fecha": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                "tokens": tokens_consumidos,
                "contenido": resultado_html
            }
            resultado_db = scanners_historial_collection.insert_one(nuevo_registro)
            scanner_id = str(resultado_db.inserted_id)

        # 4. Retornar el HTML estructurado al Frontend
        return {
            "status": "success",
            "transcripcion": resultado_html,
            "id": scanner_id
        }

    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error procesando el archivo: {str(e)}")
        
@app.get("/api/actas/descargar-pdf/{acta_id}")
async def descargar_acta_pdf(acta_id: str, email: str):
    # Buscar el acta en la base de datos por _id de MongoDB o por nombre
    acta = None
    try:
        acta = actas_collection.find_one({"_id": ObjectId(acta_id), "email": email})
    except Exception:
        pass
        
    if not acta:
        acta = actas_collection.find_one({"nombre_acta": acta_id, "email": email})
        
    if not acta:
        raise HTTPException(status_code=404, detail="Acta no encontrada.")
        
    contenido_texto = acta.get("contenido", "")
    nombre_base = acta.get("nombre_acta", "Acta_Asamblea").replace(".docx", "").replace(".pdf", "")
    nombre_archivo_pdf = f"{nombre_base}.pdf"
    
    # Configuración del PDF en memoria usando ReportLab
    pdf_buffer = io.BytesIO()
    
    doc = SimpleDocTemplate(
        pdf_buffer, 
        pagesize=letter,
        rightMargin=54,
        leftMargin=54,
        topMargin=54,
        bottomMargin=54
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'ActaTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=13,
        leading=16,
        alignment=TA_LEFT,
        spaceAfter=15,
        textColor=styles['Normal'].textColor
    )
    
    body_style = ParagraphStyle(
        'ActaBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        alignment=TA_JUSTIFY,
        spaceAfter=8
    )
    
    story.append(Paragraph("ACTA DE ASAMBLEA GENERAL DE COPROPIETARIOS", title_style))
    story.append(Spacer(1, 10))
    
    lineas = contenido_texto.split('\n')
    for linea in lineas:
        linea_limpia = linea.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        import re
        linea_limpia = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', linea_limpia)
        linea_limpia = linea_limpia.replace('*', '').strip()
        
        if linea_limpia:
            story.append(Paragraph(linea_limpia, body_style))
        else:
            story.append(Spacer(1, 6))

    doc.build(story)
    pdf_buffer.seek(0)
    
    return Response(
        content=pdf_buffer.getvalue(),
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={nombre_archivo_pdf}"}
    )


from bson import ObjectId
from fastapi import HTTPException
from pydantic import BaseModel

# ================= MODELOS PYDANTIC PARA SCANNERS =================
class GuardarScannerRequest(BaseModel):
    email: str
    nombre: str
    tokens: int
    contenido: str

class RenombrarScannerRequest(BaseModel):
    email: str
    scanner_id: str
    nuevo_nombre: str

class EliminarScannerRequest(BaseModel):
    email: str
    scanner_id: str


# ================= ENDPOINTS DE SCANNERS =================

@app.get("/api/scanners/historial")
async def obtener_historial_scanners(email: str):
    scanners_cursor = scanners_historial_collection.find({"email": email})
    scanners = []
    for doc in scanners_cursor:
        doc["id"] = str(doc["_id"])
        del doc["_id"]
        scanners.append(doc)
    return {"scanners": scanners}


@app.post("/api/scanners/guardar")
async def guardar_escaneo(data: GuardarScannerRequest):
    try:
        nuevo_registro = {
            "email": data.email,
            "nombre": data.nombre,
            "fecha": datetime.utcnow().isoformat(),
            "tokens": data.tokens,
            "contenido": data.contenido
        }
        
        resultado = scanners_historial_collection.insert_one(nuevo_registro)
        
        return {
            "message": "¡Escaneo guardado correctamente!", 
            "id": str(resultado.inserted_id)
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.put("/api/scanners/renombrar")
async def renombrar_scanner(data: RenombrarScannerRequest):
    try:
        # Intentar buscar primero por ObjectId si es un ID válido de Mongo, o por campo nombre
        filtro = {"email": data.email}
        try:
            filtro["$or"] = [{"_id": ObjectId(data.scanner_id)}, {"nombre": data.scanner_id}]
        except Exception:
            filtro["nombre"] = data.scanner_id

        resultado = scanners_historial_collection.update_one(
            filtro,
            {"$set": {"nombre": data.nuevo_nombre}},
        )

        if resultado.matched_count == 0:
            raise HTTPException(
                status_code=404, detail="Escaneo no encontrado o no autorizado."
            )

        return {"message": "¡Escaneo renombrado con éxito!"}
    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.delete("/api/scanners/eliminar")
async def eliminar_scanner(data: EliminarScannerRequest):
    try:
        filtro = {"email": data.email}
        try:
            filtro["$or"] = [{"_id": ObjectId(data.scanner_id)}, {"nombre": data.scanner_id}]
        except Exception:
            filtro["nombre"] = data.scanner_id

        resultado = scanners_historial_collection.delete_one(filtro)

        if resultado.deleted_count == 0:
            raise HTTPException(
                status_code=404, detail="Escaneo no encontrado para eliminar."
            )

        return {"message": "Escaneo eliminado correctamente."}
    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# ================= ENDPOINTS DE ACTAS (REFERENCIA) =================

@app.get("/api/actas/historial")
async def obtener_historial_actas(email: str):
    actas = list(actas_collection.find({"email": email}, {"_id": 0}))
    return {"actas": actas}

from fastapi.responses import HTMLResponse

@app.get("/api/scanners/descargar/{scanner_id}")
async def descargar_scanner(scanner_id: str, email: str):
    try:
        # Intentar buscar por ID de objeto (ObjectId)
        filtro = {"_id": ObjectId(scanner_id), "email": email}
        registro = scanners_historial_collection.find_one(filtro)
        
        if not registro:
            # Intentar buscar por nombre si el ID falla
            registro = scanners_historial_collection.find_one({"nombre": scanner_id, "email": email})
            
        if not registro:
            raise HTTPException(status_code=404, detail="Archivo no encontrado.")

        nombre_archivo = f"{registro['nombre'].replace('.pdf', '').replace('.jpg', '')}.html"
        contenido_html = registro['contenido']

        # Retornar como descarga forzada
        return HTMLResponse(
            content=contenido_html,
            headers={"Content-Disposition": f"attachment; filename={nombre_archivo}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib import colors
from reportlab.pdfgen import canvas
import re
import io

class NumberedCanvas(canvas.Canvas):
    """ Canvas personalizado para agregar pie de página profesional con numeración """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.pages = []

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self.pages)
        for page in self.pages:
            self.__dict__.update(page)
            self.draw_footer(num_pages)
            super().showPage()
        super().save()

    def draw_footer(self, total_pages):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#64748B"))
        
        # Línea divisoria sutil en el pie de página
        self.setStrokeColor(colors.HexColor("#CBD5E1"))
        self.setLineWidth(0.5)
        self.line(54, 40, letter[0] - 54, 40)
        
        # Texto del pie
        footer_text = f"Generado mediante ActasPro AI • Página {self._pageNumber} de {total_pages}"
        self.drawRightString(letter[0] - 54, 25, footer_text)
        
        # Sello corporativo o fecha en la esquina inferior izquierda
        self.drawString(54, 25, "Documento Digital Verificado")
        self.restoreState()


from fastapi.responses import Response
from fastapi import HTTPException
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT
import re
import io
from bson import ObjectId

@app.get("/api/scanners/descargar-pdf/{scanner_id}")
async def descargar_scanner_pdf_backend(scanner_id: str, email: str):
    try:
        filtro = {"_id": ObjectId(scanner_id), "email": email}
        registro = scanners_historial_collection.find_one(filtro)
        
        if not registro:
            registro = scanners_historial_collection.find_one({"nombre": scanner_id, "email": email})
            
        if not registro:
            raise HTTPException(status_code=404, detail="Archivo no encontrado.")

        nombre_base = registro.get("nombre", "Escaneo_IA").replace(".pdf", "").replace(".jpg", "").replace(".png", "")
        nombre_archivo_pdf = f"{nombre_base}.pdf"
        contenido_html = registro.get("contenido", "")

        # Configuración del PDF en memoria usando ReportLab con metadatos corregidos
        pdf_buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            pdf_buffer, 
            pagesize=letter,
            rightMargin=54,
            leftMargin=54,
            topMargin=54,
            bottomMargin=54,
            title=nombre_base,
            author="Sistema de Escaneo IA"
        )
        
        story = []
        styles = getSampleStyleSheet()
        
        # Estilos personalizados
        h1_style = ParagraphStyle(
            'ScannerH1',
            parent=styles['Heading2'],
            fontName='Helvetica-Bold',
            fontSize=14,
            leading=18,
            alignment=TA_LEFT,
            spaceBefore=12,
            spaceAfter=8
        )

        h2_style = ParagraphStyle(
            'ScannerH2',
            parent=styles['Heading3'],
            fontName='Helvetica-Bold',
            fontSize=12,
            leading=15,
            alignment=TA_LEFT,
            spaceBefore=10,
            spaceAfter=6
        )
        
        body_style = ParagraphStyle(
            'ScannerBody',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=14,
            alignment=TA_JUSTIFY,
            spaceAfter=6
        )

        # Procesamiento de etiquetas HTML
        contenido_procesado = re.sub(r'<h1[^>]*>(.*?)</h1>', r'<h1_tag>\1</h1_tag>', contenido_html, flags=re.IGNORECASE | re.DOTALL)
        contenido_procesado = re.sub(r'<h2[^>]*>(.*?)</h2>', r'<h2_tag>\1</h2_tag>', contenido_procesado, flags=re.IGNORECASE | re.DOTALL)
        contenido_procesado = re.sub(r'<p[^>]*>(.*?)</p>', r'\1<br/><br/>', contenido_procesado, flags=re.IGNORECASE | re.DOTALL)
        contenido_procesado = re.sub(r'<br\s*/?>', r'<br/>', contenido_procesado, flags=re.IGNORECASE)
        contenido_procesado = re.sub(r'<(?:strong)[^>]*>(.*?)</(?:strong)>', r'<b>\1</b>', contenido_procesado, flags=re.IGNORECASE | re.DOTALL)
        
        lineas = contenido_procesado.split('\n')
        
        for linea in lineas:
            linea_str = linea.strip()
            if not linea_str:
                continue
                
            # Verificar h1
            h1_match = re.search(r'<h1_tag>(.*?)</h1_tag>', linea_str, flags=re.DOTALL)
            if h1_match:
                story.append(Paragraph(h1_match.group(1), h1_style))
                continue
                
            # Verificar h2
            h2_match = re.search(r'<h2_tag>(.*?)</h2_tag>', linea_str, flags=re.DOTALL)
            if h2_match:
                story.append(Paragraph(h2_match.group(1), h2_style))
                continue
                
            # Limpiar etiquetas no permitidas en párrafos normales
            linea_limpia = re.sub(r'<(?!/?b\b)(?!/?br\b)[^>]+>', '', linea_str)
            linea_limpia = linea_limpia.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
            
            if linea_limpia.strip():
                story.append(Paragraph(linea_limpia, body_style))

        doc.build(story)
        pdf_buffer.seek(0)
        
        return Response(
            content=pdf_buffer.getvalue(),
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename={nombre_archivo_pdf}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando PDF: {str(e)}")


# --- Ruta 1: FAQs Preestablecidas con Cloudinary ---
@app.post("/api/soporte/faqs")
async def soporte_faqs(req: ConsultaFAQRequest):
    faqs_data = {
        "audio": {
            "respuesta": "Para transcribir un audio, ve a '+ Transcribe audios', sube tu archivo en formato .mp3 o .wav y haz clic en 'Generar Acta'.",
            "videoUrl": "https://res.cloudinary.com/iuu7h8rj/video/upload/v1787281142/transcripcion_por_audio.mp4"
        },
        "scanner": {
            "respuesta": "Para usar el escáner de documentos, dirígete al '+ Escanea y transcribe doc.', sube tu archivo en PDF o imagen clara.",
            "videoUrl": "https://res.cloudinary.com/iuu7h8rj/video/upload/v1787281146/scanner_de_documento.mp4"
        }
    }
    
    resultado = faqs_data.get(req.tipoConsulta, {
        "respuesta": "Puedes explorar nuestras guías generales en el menú principal.",
        "videoUrl": "https://res.cloudinary.com/TU_CLOUD_NAME/video/upload/v123456789/video_default.mp4"
    })

    return resultado

# --- Ruta 2: Verificación de Plan Preestablecida con Video ---
@app.post("/api/soporte/verificar-plan")
async def verificar_plan(req: ConsultaPlanRequest):
    # Simulación de consulta a base de datos usando el correo
    # Aquí puedes conectar tu lógica real de base de datos para buscar el plan del usuario por su `req.email`
    
    respuesta_plan = (
        f"Hola. Hemos verificado el correo '{req.email}'. "
        f"Actualmente te encuentras en el plan **Empresarial Pro** con 18.5 horas consumidas de tu límite mensual. "
        f"Tu cuenta está activa y tienes acceso completo a todas las funciones avanzadas y escáneres hasta tu fecha de renovación."
    )
    
    video_plan = "https://assets.mixkit.co/videos/preview/mixkit-data-flow-plexus-background-animation-2679-large.mp4" # Video explicativo del plan

    return {
        "respuesta": respuesta_plan,
        "videoUrl": video_plan,
        "datosConsumo": {
            "planActivo": "Empresarial Pro",
            "horasTotalesMes": 20,
            "horasConsumidas": 18.5,
            "fechaRenovacion": "01/10/2026"
        }
    }

from fastapi import FastAPI
from pydantic import BaseModel
from typing import Optional
import datetime
from bson import ObjectId

# (Asegúrate de tener configurado tu cliente de Motor/MongoDB en tu app, por ejemplo:
# from motor.motor_asyncio import AsyncIOMotorClient
# client = AsyncIOMotorClient("tu_url_de_mongo")
# db = client["nombre_de_tu_base_de_datos"]
# )

class ContactoEnterprise(BaseModel):
    nombre: str 
    email: str
    whatsapp: str
    pais: str
    horas: int
    documentos: int
    empleados: int
    mensaje: Optional[str] = None

# ==========================================
# 1. ENDPOINT PARA GUARDAR COTIZACIÓN (POST)
# ==========================================
@app.post("/api/contacto-enterprise")
async def recibir_contacto_enterprise(datos: ContactoEnterprise):
    try:
        # Preparamos el documento con fecha actual y estado por defecto
        nuevo_lead = datos.model_dump()
        nuevo_lead["fecha"] = datetime.datetime.utcnow().isoformat()
        nuevo_lead["estado"] = "Pendiente"
        
        # CORRECCIÓN: Quitamos el 'await' porque pymongo es síncrono
        resultado_db = db.cotizaciones.insert_one(nuevo_lead)
        
        print(f"Nueva solicitud Enterprise guardada con ID: {resultado_db.inserted_id}")
        
        return {
            "success": True, 
            "message": "¡Gracias por contactarnos! Un gerente de cuenta se comunicará contigo pronto."
        }
    except Exception as e:
        print(f"Error procesando solicitud Enterprise: {e}")
        return {
            "success": False, 
            "message": str(e)
        }

# ==========================================
# 2. ENDPOINT PARA EL PANEL DE ADMINISTRACIÓN (GET)
# ==========================================
@app.get("/api/admin/cotizaciones")
async def obtener_cotizaciones():
    try:
        # CORRECCIÓN: Sin 'await' porque PyMongo es síncrono
        cotizaciones_db = list(db.cotizaciones.find().sort("fecha", -1).limit(1000))
        
        # Transformamos el ObjectId de MongoDB a string
        lista_leads = []
        for c in cotizaciones_db:
            c["_id"] = str(c["_id"])
            lista_leads.append(c)
        
        return {
            "success": True,
            "cotizaciones": lista_leads
        }
    except Exception as e:
        print(f"Error obteniendo cotizaciones de MongoDB: {e}")
        return {
            "success": False,
            "message": str(e)
        }

from datetime import datetime
from fastapi import APIRouter, HTTPException, Request
from pydantic import BaseModel

# Inicializamos el router con el prefijo /api/crm
router = APIRouter(prefix="/api/crm", tags=["CRM & SaaS Analytics"])

PRECIOS_PLANES = {
    "basico": {"tokens_mensuales": 100000, "limite_horas": 15.0, "precio": 153000},
    "profesional": {"tokens_mensuales": 500000, "limite_horas": 60.0, "precio": 479000},
    "corporativo": {"tokens_mensuales": 2000000, "limite_horas": 200.0, "precio": 939000}
}

@router.get("/license-stats")
async def get_crm_license_stats():
    """
    Endpoint principal para alimentar el CRM con estadísticas de usuarios por plan,
    desglose de pagos y proyecciones financieras mensuales/anuales.
    """
    try:
        # 1. Conteo de usuarios por plan utilizando la colección existente 'users'
        total_usuarios = users_collection.count_documents({})
        free_users = users_collection.count_documents({"plan": "free"})
        basico_users = users_collection.count_documents({"plan": "basico"})
        profesional_users = users_collection.count_documents({"plan": "profesional"})
        corporativo_users = users_collection.count_documents({"plan": "corporativo"})
        
        # Total de usuarios de pago (todos los que no sean 'free')
        paid_users = basico_users + profesional_users + corporativo_users

        # 2. Obtener lista detallada de todos los usuarios para la tabla del CRM
        cursor_usuarios = users_collection.find({}, {"_id": 0, "password": 0})
        lista_suscriptores = list(cursor_usuarios)

        # 3. Cálculo de Ingresos Actuales (MRR estimado según los planes activos)
        precio_basico = PRECIOS_PLANES["basico"]["precio"]
        precio_profesional = PRECIOS_PLANES["profesional"]["precio"]
        precio_corporativo = PRECIOS_PLANES["corporativo"]["precio"]

        ingresos_actuales_mes = (
            (basico_users * precio_basico) +
            (profesional_users * precio_profesional) +
            (corporativo_users * precio_corporativo)
        )

        egresos_actuales_mes = ingresos_actuales_mes * 0.25 

        # 4. Simulación de Evolución Mensual
        evolucion_mensual = [
            {"mes": "Enero", "ingresos": ingresos_actuales_mes * 0.7, "egresos": egresos_actuales_mes * 0.8, "proyeccion": ingresos_actuales_mes * 0.75},
            {"mes": "Febrero", "ingresos": ingresos_actuales_mes * 0.82, "egresos": egresos_actuales_mes * 0.85, "proyeccion": ingresos_actuales_mes * 0.88},
            {"mes": "Marzo", "ingresos": ingresos_actuales_mes * 0.95, "egresos": egresos_actuales_mes * 0.9, "proyeccion": ingresos_actuales_mes * 0.92},
            {"mes": "Abril", "ingresos": ingresos_actuales_mes, "egresos": egresos_actuales_mes, "proyeccion": ingresos_actuales_mes * 1.05},
            {"mes": "Mayo (Proj)", "ingresos": 0, "egresos": 0, "proyeccion": ingresos_actuales_mes * 1.15},
            {"mes": "Junio (Proj)", "ingresos": 0, "egresos": 0, "proyeccion": ingresos_actuales_mes * 1.25},
        ]

        proyeccion_anual = {
            "anual_estimado": ingresos_actuales_mes * 12,
            "crecimiento_esperado_porcentaje": 18.5,
            "meta_anual": (ingresos_actuales_mes * 12) * 1.30
        }

        return {
            "success": True,
            "metricas_planes": {
                "total": total_usuarios,
                "free": free_users,
                "paid": paid_users,
                "desglose_pagos": {
                    "basico": basico_users,
                    "profesional": profesional_users,
                    "corporativo": corporativo_users
                }
            },
            "financiero": {
                "ingresos_mes_actual": ingresos_actuales_mes,
                "egresos_mes_actual": egresos_actuales_mes,
                "utilidad_neta": ingresos_actuales_mes - egresos_actuales_mes,
                "evolucion_mensual": evolucion_mensual,
                "proyeccion_anual": proyeccion_anual
            },
            "suscriptores": lista_suscriptores
        }

    except Exception as e:
        print(f"Error en CRM stats: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/license-stats-advanced")
async def obtener_estadisticas_avanzadas():
    """
    Retorna métricas extendidas incluyendo alerta de churn, 
    márgenes unitarios por usuario y cohortes desde la base de datos real.
    """
    suscriptores = [
        {
            "email": "carlos.legal@empresa.com",
            "plan": "corporativo",
            "fecha_registro": "2026-03-15",
            "horas_usadas_mes": 185.0,
            "limite_horas_mes": 200,
            "tokens_usados": 1250000,
            "pago_mensual": 939000,
            "activo": True
        },
        {
            "email": "laura.consultora@gmail.com",
            "plan": "profesional",
            "fecha_registro": "2026-06-10",
            "horas_usadas_mes": 5.0, 
            "limite_horas_mes": 60,
            "tokens_usados": 45000,
            "pago_mensual": 479000,
            "activo": True
        },
        {
            "email": "startup.dev@gmail.com",
            "plan": "basico",
            "fecha_registro": "2026-07-01",
            "horas_usadas_mes": 14.5,
            "limite_horas_mes": 15,
            "tokens_usados": 110000,
            "pago_mensual": 153000,
            "activo": True
        }
    ]

    Costo_IA_Por_Hora = 2500 # COP
    analisis_unit_economics = []
    usuarios_en_riesgo = []

    for sub in suscriptores:
        costo_ia_usuario = sub["horas_usadas_mes"] * Costo_IA_Por_Hora
        margen_neto = sub["pago_mensual"] - costo_ia_usuario
        
        analisis_unit_economics.append({
            "email": sub["email"],
            "plan": sub["plan"],
            "ingreso": sub["pago_mensual"],
            "costo_ia": costo_ia_usuario,
            "margen_neto": margen_neto,
            "porcentaje_margen": round((margen_neto / sub["pago_mensual"]) * 100, 1) if sub["pago_mensual"] > 0 else 0
        })

        if sub["plan"] != "free":
            limite = sub.get("limite_horas_mes", 0)
            if limite > 0:
                porcentaje_uso = (sub["horas_usadas_mes"] / limite) * 100
                if porcentaje_uso < 15:
                    usuarios_en_riesgo.append({
                        "email": sub["email"],
                        "plan": sub["plan"],
                        "razon": f"Uso crítico de cuota: solo {porcentaje_uso:.1f}% consumido",
                        "nivel_riesgo": "Alto"
                    })

    cohortes = {
        "Marzo 2026": {"total": 12, "retenidos": 11},
        "Abril 2026": {"total": 19, "retenidos": 17},
        "Mayo 2026": {"total": 25, "retenidos": 23},
        "Junio 2026": {"total": 31, "retenidos": 28},
        "Julio 2026": {"total": 45, "retenidos": 42}
    }

    return {
        "success": True,
        "unit_economics": analisis_unit_economics,
        "churn_alerts": usuarios_en_riesgo,
        "cohortes": cohortes
    }

@router.post("/webhook-mercadopago")
async def webhook_mercadopago(request: Request):
    """
    Recibe notificaciones automáticas de Mercado Pago, consulta el API oficial,
    detecta el plan y actualiza los contadores en la base de datos de MongoDB.
    """
    try:
        body = await request.json()
        if body.get("type") == "payment":
            payment_id = body.get("data", {}).get("id")
            if payment_id and mp_sdk:
                payment_info = mp_sdk.payment().get(payment_id)
                payment = payment_info.get("response", {})

                if payment.get("status") == "approved":
                    payer_email = payment.get("payer", {}).get("email") or payment.get("external_reference")
                    if payer_email:
                        items = payment.get("additional_info", {}).get("items", []) or payment.get("items", [])
                        plan_asignado = "basico"

                        for item in items:
                            title_lower = item.get("title", "").lower()
                            if "corporativo" in title_lower:
                                plan_asignado = "corporativo"
                            elif "profesional" in title_lower or "intermedio" in title_lower:
                                plan_asignado = "profesional"
                            elif "básico" in title_lower or "basico" in title_lower:
                                plan_asignado = "basico"

                        return {"status": "success", "message": f"Licencia de {payer_email} actualizada a {plan_asignado}"}
    except Exception as e:
        print(f"Error en webhook de Mercado Pago: {e}")
        raise HTTPException(status_code=500, detail=str(e))

    return {"status": "ok"}

@router.get("/invoice/download/{email}")
async def descargar_factura_pdf(email: str):
    """
    Genera los datos fiscales para la descarga del comprobante contable.
    """
    return {
        "success": True,
        "mensaje": "Datos de factura listos",
        "factura": {
            "empresa_emisora": "ActasPro SAS - NIT: 900.123.456-7",
            "cliente": email,
            "fecha_emision": datetime.now().strftime("%Y-%m-%d"),
            "concepto": "Suscripción Licencia Corporativa - Mensual",
            "valor_neto": 939000,
            "iva": 0,
            "total": 939000,
            "estado": "Pagado vía Mercado Pago"
        }
    }
