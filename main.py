import os
import uuid
import shutil
from typing import Optional
import assemblyai as aai
from openai import OpenAI
from docx import Document
from dotenv import load_dotenv
from pydantic import BaseModel
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pymongo import MongoClient
import mercadopago

load_dotenv()

app = FastAPI(title="ActaBot PH con MongoDB y Mercado Pago", version="1.5")

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
    raise ValueError("❌ Error: Claves de API de AssemblyAI u OpenAI no configuradas en .env")

# Inicializar cliente de MongoDB especificando la base de datos aislada 'actabot_db'
mongo_client = MongoClient(MONGO_URI)
db = mongo_client["actabot_db"]
users_collection = db["users"]

# Inicializar SDK de Mercado Pago
mp_sdk = mercadopago.SDK(MP_ACCESS_TOKEN) if MP_ACCESS_TOKEN else None

aai.settings.api_key = AAI_API_KEY
openai_client = OpenAI(api_key=OPENAI_API_KEY)

# Diccionario seguro de precios y planes en el backend para evitar fraudes desde el navegador
PRECIOS_PLANES = {
    "basico": {
        "nombre": "Plan Básico",
        "precio": 49000.0,
        "tokens": 50
    },
    "profesional": {
        "nombre": "Plan Profesional",
        "precio": 149000.0,
        "tokens": 200
    },
    "corporativo": {
        "nombre": "Plan Corporativo",
        "precio": 299000.0,
        "tokens": 9999
    }
}

class AuthModel(BaseModel):
    email: str
    password: str
    plan: Optional[str] = "free"

class PaymentPreferenceModel(BaseModel):
    email: str
    plan_name: str  # Recibe la clave segura: 'basico', 'profesional', 'corporativo'

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
        "tokens": tokens_iniciales
    }
    
    users_collection.insert_one(nuevo_usuario)
    return {"message": "Usuario registrado con éxito en MongoDB", "email": data.email, "plan": data.plan, "tokens": tokens_iniciales}

@app.post("/api/login")
def login_usuario(data: AuthModel):
    user = users_collection.find_one({"email": data.email})
    if not user or user["password"] != data.password:
        raise HTTPException(status_code=401, detail="Credenciales inválidas.")
    
    return {
        "message": "Login exitoso",
        "email": user["email"],
        "plan": user["plan"],
        "tokens": user["tokens"]
    }

@app.get("/api/user/status")
def get_user_status(email: str):
    usuario = users_collection.find_one({"email": email})
    
    if not usuario:
        raise HTTPException(status_code=404, detail="Usuario no encontrado en la base de datos")
    
    return {
        "email": usuario.get("email"),
        "plan": usuario.get("plan", "free"),
        "tokens": usuario.get("tokens", 0)
    }    

@app.post("/api/crear-preferencia-pago")
def crear_preferencia_pago(data: PaymentPreferenceModel):
    if not mp_sdk:
        raise HTTPException(status_code=500, detail="Mercado Pago no está configurado en el servidor.")
        
    # Validar que el plan exista en el diccionario seguro del backend
    plan_id = data.plan_name.lower()
    if plan_id not in PRECIOS_PLANES:
        raise HTTPException(status_code=400, detail="Plan no válido.")
    
    info_plan = PRECIOS_PLANES[plan_id]

    # Verificar si el usuario existe en MongoDB; si no existe, crearlo como invitado provisional
    user = users_collection.find_one({"email": data.email})
    if not user:
        users_collection.insert_one({
            "email": data.email,
            "password": "temp_password_temporal", 
            "plan": "free",
            "tokens": 5
        })

    preference_data = {
        "items": [
            {
                "title": f"ActaBot PH - {info_plan['nombre']}",
                "quantity": 1,
                "currency_id": "COP",
                "unit_price": info_plan["precio"]  # Precio seguro definido estrictamente en el backend
            }
        ],
        "payer": {
            "email": data.email
        },
        "back_urls": {
            "success": "http://127.0.0.1:8000/pago-exitoso",
            "failure": "http://127.0.0.1:8000/pago-fallido",
            "pending": "http://127.0.0.1:8000/pago-pendiente"
        },
        "auto_return": "approved",
        "notification_url": "https://tu-dominio.com/api/webhook-mercadopago" 
    }

    try:
        preference_response = mp_sdk.preference().create(preference_data)
        preference = preference_response["response"]
        
        return {
            "init_point": preference.get("init_point"),
            "sandbox_init_point": preference.get("sandbox_init_point")
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
                    payer_email = payment.get("payer", {}).get("email")
                    if payer_email:
                        # Identificar qué plan se pagó a través del título o el monto para actualizar acorde
                        items = payment.get("additional_info", {}).get("items", []) or payment.get("items", [])
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
                            {"$set": {"plan": plan_asignado, "tokens": tokens_otorgados}}
                        )
    except Exception as e:
        print(f"Error en webhook: {e}")
        
    return {"status": "ok"}

@app.post("/procesar")
async def procesar_asamblea(
    file: UploadFile = File(...),
    instrucciones: Optional[str] = Form(""),
    email: str = Form(...)
):
    user = users_collection.find_one({"email": email})
    if not user:
        raise HTTPException(status_code=401, detail="Usuario no autenticado.")
    
    if user["plan"] == "free" and user["tokens"] <= 0:
        raise HTTPException(status_code=403, detail="Has agotado tus 5 tokens gratuitos. Actualiza a un plan de pago.")

    os.makedirs("temp_uploads", exist_ok=True)
    os.makedirs("temp_outputs", exist_ok=True)

    session_id = str(uuid.uuid4())
    temp_audio_path = f"temp_uploads/{session_id}_{file.filename}"
    output_docx_path = f"temp_outputs/Acta_{session_id}.docx"

    try:
        with open(temp_audio_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        config = aai.TranscriptionConfig(speaker_labels=True, language_code="es")
        transcriber = aai.Transcriber()
        transcript = transcriber.transcribe(temp_audio_path, config=config)

        if transcript.status == aai.TranscriptStatus.error:
            raise HTTPException(status_code=500, detail=f"Error en AssemblyAI: {transcript.error}")

        texto_transcrito = ""
        if transcript.utterances:
            for utterance in transcript.utterances:
                texto_transcrito += f"[Persona {utterance.speaker}]: {utterance.text}\n"
        else:
            texto_transcrito = transcript.text

        prompt_sistema = "Eres un Secretario Jurídico experto en Propiedad Horizontal en Colombia (Ley 675 de 2001). Redacta un ACTA DE ASAMBLEA FORMAL."
        if instrucciones:
            prompt_sistema += f"\n\nREGLAS:\n{instrucciones}"

        response = openai_client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": prompt_sistema},
                {"role": "user", "content": f"Transcripción:\n\n{texto_transcrito}"}
            ],
            temperature=0.3
        )

        acta_final = response.choices[0].message.content

        doc = Document()
        doc.add_heading('ACTA DE ASAMBLEA DE COPROPIETARIOS', 0)
        for linea in acta_final.split('\n'):
            if linea.strip():
                doc.add_paragraph(linea.strip())
        doc.save(output_docx_path)

        if user["plan"] == "free":
            nuevos_tokens = user["tokens"] - 1
            users_collection.update_one(
                {"email": email},
                {"$set": {"tokens": nuevos_tokens}}
            )

        return FileResponse(
            path=output_docx_path,
            filename=f"Acta_Asamblea_{session_id[:8]}.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if os.path.exists(temp_audio_path):
            os.remove(temp_audio_path)