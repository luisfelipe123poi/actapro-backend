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
import io
from bson import ObjectId

load_dotenv()

app = FastAPI(title="ActaBot PH con MongoDB y Mercado Pago", version="1.9.5")

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
PROMPT_SISTEMA_ACTAS = """
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

# Diccionario seguro de precios y planes controlados estrictamente en el backend
PRECIOS_PLANES = {
    "basico": {"nombre": "Plan Básico", "precio": 49000.0, "tokens": 50},
    "profesional": {
        "nombre": "Plan Profesional",
        "precio": 149000.0,
        "tokens": 200,
    },
    "corporativo": {
        "nombre": "Plan Corporativo",
        "precio": 299000.0,
        "tokens": 9999,
    },
}


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


@app.post("/api/registro")
def registrar_usuario(data: AuthModel):
  existing_user = users_collection.find_one({"email": data.email})
  if existing_user:
    raise HTTPException(status_code=400, detail="El correo ya está registrado.")

  tokens_iniciales = 5 if data.plan == "free" else 9999

  nuevo_usuario = {
      "email": data.email,
      "password": data.password,
      "plan": data.plan,
      "tokens": tokens_iniciales,
  }

  users_collection.insert_one(nuevo_usuario)
  return {
      "message": "Usuario registrado con éxito en MongoDB",
      "email": data.email,
      "plan": data.plan,
      "tokens": tokens_iniciales,
  }


@app.post("/api/login")
def login_usuario(data: AuthModel):
  user = users_collection.find_one({"email": data.email})
  if not user or user["password"] != data.password:
    raise HTTPException(status_code=401, detail="Credenciales inválidas.")

  return {
      "message": "Login exitoso",
      "email": user["email"],
      "plan": user["plan"],
      "tokens": user["tokens"],
  }


@app.get("/api/user/status")
def get_user_status(email: str):
  usuario = users_collection.find_one({"email": email})
  if not usuario:
    raise HTTPException(
        status_code=404, detail="Usuario no encontrado en la base de datos"
    )

  return {
      "email": usuario.get("email"),
      "plan": usuario.get("plan", "free"),
      "tokens": usuario.get("tokens", 0),
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


@app.get("/api/actas/descargar-pdf/{acta_id}")
async def descargar_acta_pdf(acta_id: str, email: str):
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
    nombre_archivo = acta.get("nombre_acta", "acta").replace(".docx", "").replace(".pdf", "") + ".pdf"
    
    # Configuración del PDF en memoria usando ReportLab Platypus
    pdf_buffer = io.BytesIO()
    
    # Márgenes de 0.5 pulgadas (36 puntos) o 54 puntos (3/4 de pulgada)
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
    
    # Añadir título principal al documento
    story.append(Paragraph("ACTA DE ASAMBLEA GENERAL DE COPROPIETARIOS", title_style))
    story.append(Spacer(1, 10))
    
    # Procesar el contenido línea por línea o párrafo por párrafo
    lineas = contenido_texto.split('\n')
    for linea in lineas:
        # Limpieza básica de caracteres Markdown para el reporte PDF
        linea_limpia = linea.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        
        # Reemplazar negritas de markdown (**texto**) por etiquetas HTML soportadas por ReportLab (<b>texto</b>)
        import re
        linea_limpia = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', linea_limpia)
        linea_limpia = linea_limpia.replace('*', '').strip()
        
        if linea_limpia:
            story.append(Paragraph(linea_limpia, body_style))
        else:
            # Si hay una línea vacía, agregamos un pequeño espacio vertical
            story.append(Spacer(1, 6))

    # Construir el PDF
    doc.build(story)
    pdf_buffer.seek(0)
    
    return Response(
        content=pdf_buffer.getvalue(),
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={nombre_archivo}"}
    )


@app.post("/api/crear-preferencia-pago")
def crear_preferencia_pago(data: PaymentPreferenceModel):
  if not mp_sdk:
    raise HTTPException(
        status_code=500, detail="Mercado Pago no está configurado en el servidor."
    )

  plan_recibido = data.plan_name.lower().strip()

  mapeo = {
      "basico": "basico",
      "básico": "basico",
      "plan basico": "basico",
      "plan básico": "basico",
      "profesional": "profesional",
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

  user = users_collection.find_one({"email": data.email})
  if not user:
    users_collection.insert_one({
        "email": data.email,
        "password": "temp_password_temporal",
        "plan": "free",
        "tokens": 5,
    })

  preference_data = {
      "items": [{
          "title": f"ActaBot PH - {info_plan['nombre']}",
          "quantity": 1,
          "currency_id": "COP",
          "unit_price": info_plan["precio"],
      }],
      "payer": {"email": data.email},
      "back_urls": {
          "success": "https://actapro-backend.onrender.com/pago-exitoso",
          "failure": "https://actapro-backend.onrender.com/pago-fallido",
          "pending": "https://actapro-backend.onrender.com/pago-pendiente",
      },
      "auto_return": "approved",
      "notification_url": (
          "https://actapro-backend.onrender.com/api/webhook-mercadopago"
      ),
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
          detail=(
              "Mercado Pago rechazó la preferencia o devolvió credenciales"
              " inválidas."
          ),
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
          payer_email = payment.get("payer", {}).get(
              "email"
          ) or payment.get("external_reference")
          if payer_email:
            items = payment.get("additional_info", {}).get("items", []) or (
                payment.get("items", [])
            )
            plan_asignado = "profesional"
            tokens_otorgados = 200

            for item in items:
              title_lower = item.get("title", "").lower()
              if "básico" in title_lower or "basico" in title_lower:
                plan_asignado = "basico"
                tokens_otorgados = 50
              elif "corporativo" in title_lower:
                plan_asignado = "corporativo"
                tokens_otorgados = 9999

            users_collection.update_one(
                {"email": payer_email},
                {
                    "$set": {
                        "plan": plan_asignado,
                        "tokens": tokens_otorgados,
                    }
                },
            )
  except Exception as e:
    print(f"Error en webhook: {e}")

  return {"status": "ok"}


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

    if user["plan"] == "free" and user["tokens"] <= 0:
        raise HTTPException(
            status_code=403,
            detail=(
                "Has agotado tus 5 tokens gratuitos. Actualiza a un plan de pago."
            ),
        )

    os.makedirs("temp_uploads", exist_ok=True)
    os.makedirs("temp_outputs", exist_ok=True)

    session_id = str(uuid.uuid4())
    temp_audio_path = f"temp_uploads/{session_id}_{file.filename}"

    try:
        content_bytes = await file.read()
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
            with open(temp_audio_path, "wb") as buffer:
                buffer.write(content_bytes)

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
            # Asegurar que no tenga extensiones duplicadas
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

        if user["plan"] == "free":
            nuevos_tokens = user["tokens"] - 1
            users_collection.update_one(
                {"email": email}, {"$set": {"tokens": nuevos_tokens}}
            )

        return FileResponse(
            path=output_docx_path,
            filename=nombre_archivo_acta,
            media_type=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.Document"
            ),
        )

    except Exception as e:
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
