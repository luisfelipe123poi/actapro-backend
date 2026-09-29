import os
import io
import uuid
import base64
import hashlib
from datetime import datetime, timezone
from io import BytesIO

import boto3
from botocore.config import Config
import requests
import pymupdf as fitz
import pdfplumber
from docx import Document
from pymongo import MongoClient
import assemblyai as aai
import openai

from celery_app import celery_app

# ==========================================
# 1. CONFIGURACIÓN Y CREDENCIALES (ENV)
# ==========================================
R2_ENDPOINT_URL = os.getenv("R2_ENDPOINT_URL")
R2_ACCESS_KEY_ID = os.getenv("R2_ACCESS_KEY_ID")
R2_SECRET_ACCESS_KEY = os.getenv("R2_SECRET_ACCESS_KEY")
R2_BUCKET_NAME = os.getenv("R2_BUCKET_NAME", "archivos-temporales-actaprocore")
R2_PUBLIC_URL = os.getenv("R2_PUBLIC_URL", "https://cdn.actaprocore.com")

AAI_API_KEY = os.getenv("ASSEMBLYAI_API_KEY")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
MONGO_URI = os.getenv("MONGO_URI", "mongodb://localhost:27017/")

aai.settings.api_key = AAI_API_KEY
openai_client = openai.OpenAI(api_key=OPENAI_API_KEY)
mongo_client = MongoClient(MONGO_URI)
db = mongo_client["actabot_db"]

# Colecciones de MongoDB
users_collection = db["users"]
actas_collection = db["actas_historial"]
transripciones_collection = db["transripciones_cache"]
scanners_historial_collection = db["scanners_historial"]

# ==========================================
# PROMPTS ESPECIALIZADOS POR NICHO / MOTOR
# ==========================================

PROMPT_SISTEMA_ACTAS = """
Eres un Secretario Jurídico experto en Propiedad Horizontal en Colombia (Ley 675 de 2001). 
Tu objetivo es redactar un acta de asamblea formal, íntegra y detallada a partir de la transcripción provista (con diarización de voces), siguiendo estrictamente la estructura estándar de un documento oficial corporativo listo para imprimir y firmar.

ESTRUCTURA OBLIGATORIA QUE DEBES GENERAR (SIN MODIFICAR EL ORDEN):

ACTA DE ASAMBLEA GENERAL DE COPROPIETARIOS
ACTA DE ASAMBLEA DE COPROPIETARIOS DEL CONJUNTO RESIDENCIAL [Nombre del Conjunto o Copropiedad extraído del audio]
En la ciudad de [Ciudad], a [Fecha], siendo las [Hora], se reunieron los copropietarios... [Desarrollar la introducción formal con los datos reales capturados del audio].

ASISTENTES
Se registraron los siguientes asistentes:
1. [Nombre completo] - Unidad [Número] - Coeficiente [Valor]
(Lista detallada de todos los copropietarios, unidades y coeficientes que se mencionen en la transcripción).
Se constató la existencia de quórum suficiente para deliberar y decidir sobre los puntos del orden del día.

ORDEN DEL DÍA
1. [Primer punto tratado]
2. [Segundo punto tratado]
(Y así sucesivamente según los puntos reales de la reunión).

DESARROLLO DE LA ASAMBLEA
PUNTO PRIMERO: [TÍTULO DEL PUNTO]
[Narrativa detallada y circunstanciada del debate, montos, cifras, saldos, nombres de quienes intervinieron y explicaciones dadas, sin omitir información clave].
DECISIONES: [Detalle preciso de lo approved, votado o resuelto].
PENDIENTES: [Tareas, responsables o acciones abiertas, o "Ninguno"].

(Repetir la misma estructura de DECISIONES y PENDIENTES para cada uno de los puntos del orden del día).

RESUMEN GENERAL DE DECISIONES Y PENDIENTES DE LA REUNIÓN
1. [Resumen consolidado de la decisión 1]
2. [Resumen consolidado de la decisión 2]
(Lista limpia y numerada con todos los acuerdos y tareas pendientes).

Sin más asuntos que tratar, se da por concluida la asamblea a las [Hora de Cierre], firmando al pie los asistentes.

FIRMAS
_____________________________
[Nombre real extraído de la transcripción como Presidente]
Presidente de la Asamblea

_____________________________
[Nombre real extraído de la transcripción como Secretario(a)]
Secretario(a) de la Asamblea

_____________________________
[Nombre real o miembro de la Comisión Verificadora]
Comisión Verificadora del Acta

_____________________________
[Nombre real o miembro de la Comisión Verificadora]
Comisión Verificadora del Acta


REGLAS ESTRICTAS DE REDACCIÓN:
1. EXTRACCIÓN REAL: Utiliza los nombres de las personas, cargos, montos, saldos, unidades y datos específicos que escuches en la transcripción para rellenar automáticamente los campos. Solo usa corchetes si un dato es imposible de determinar.
2. CERO RESUMENES VACÍOS: Mantén una redacción profesional, jurídica, detallada y extensa donde se refleje todo lo que se discutió.
3. FORMATO LIMPIO: NO utilices asteriscos (*), símbolos de almohadilla (#) ni markdown crudo. Usa exclusivamente texto plano con títulos en MAYÚSCULAS SOSTENIDAS tal como se indicó en la estructura.
"""

PROMPT_SISTEMA_CORPORATIVO = """
Eres un Secretario Corporativo y Consultor de Negocios experto en Juntas Directivas y Reuniones Gerenciales.
Tu objetivo es redactar un acta corporativa formal, íntegra y detallada a partir de la transcripción provista (con diarización de voces), siguiendo estrictamente la estructura estándar de gobierno corporativo lista para imprimir y firmar.

ESTRUCTURA OBLIGATORIA QUE DEBES GENERAR (SIN MODIFICAR EL ORDEN):

ACTA DE JUNTA DIRECTIVA / REUNIÓN CORPORATIVA
ACTA DE LA SESIÓN DE LA JUNTA DIRECTIVA DE [Nombre de la Empresa u Organización extraído del audio]
En la ciudad de [Ciudad], a [Fecha], siendo las [Hora], se reunieron de manera virtual o presencial los miembros de la Junta Directiva... [Desarrollar la introducción formal con los datos reales capturados del audio].

MIEMBROS Y ASISTENTES PRESENTES
Se registraron los siguientes asistentes y cargos:
1. [Nombre completo] - [Cargo o Rol, ej. Presidente de Junta / Director Ejecutivo]
(Lista detallada de directivos, invitados y asistentes mencionados en la transcripción).
Se verificó el quórum reglamentario para dar inicio formal a la sesión.

ORDEN DEL DÍA
1. [Primer punto tratado]
2. [Segundo punto tratado]
(Y así sucesivamente según los puntos reales de la sesión).

DESARROLLO DE LA SESIÓN
PUNTO PRIMERO: [TÍTULO DEL PUNTO]
[Narrativa detallada y gerencial del análisis financiero, operativo o estratégico, indicadores clave (KPIs), discusiones, cifras expuestas y observaciones de los directivos].
ACUERDOS Y APROBACIONES: [Detalle preciso de las votaciones, resoluciones o directrices aprobadas].
TAREAS Y PLANES DE ACCIÓN: [Tareas asignadas, responsables y plazos de entrega, o "Ninguno"].

(Repetir la misma estructura de ACUERDOS y TAREAS para cada punto del orden del día).

RESUMEN EJECUTIVO DE ACUERDOS Y COMPROMISOS
1. [Resumen consolidado del acuerdo 1]
2. [Resumen consolidado del acuerdo 2]
(Lista limpia y numerada con todos los compromisos y responsables).

Sin otro particular que tratar, se levanta la sesión a las [Hora de Cierre], en constancia firman los inescritos.

FIRMAS
_____________________________
[Nombre real extraído de la transcripción]
Presidente de la Junta Directiva

_____________________________
[Nombre real extraído de la transcripción]
Secretario(a) de la Junta

REGLAS ESTRICTAS DE REDACCIÓN:
1. EXTRACCIÓN REAL: Utiliza los nombres, cargos, metas, presupuestos y plazos específicos mencionados en el audio. Usa corchetes solo si falta información indispensable.
2. CERO RESUMENES VACÍOS: Mantén un lenguaje ejecutivo, financiero y estratégico detallado.
3. FORMATO LIMPIO: NO utilices asteriscos (*), símbolos de almohadilla (#) ni markdown crudo. Usa texto plano con títulos en MAYÚSCULAS SOSTENIDAS.
"""

PROMPT_SISTEMA_LEGAL = """
Eres un Abogado Litigante y Asesor Jurídico Senior experto en derecho procesal y redacción de dictámenes y audiencias.
Tu objetivo es redactar un dictamen, acta de audiencia o concepto legal formal, íntegro y riguroso a partir de la transcripción provista (con diarización de voces), siguiendo strictly la estructura jurídica oficial.

ESTRUCTURA OBLIGATORIA QUE DEBES GENERAR (SIN MODIFICAR EL ORDEN):

DICTAMEN Y ACTA DE AUDIENCIA / ASESORÍA LEGAL
EXPEDIENTE / RADICADO: [Número o referencia extraída, o "N/A"]
En la ciudad de [Ciudad], a [Fecha], siendo las [Hora], se lleva a cabo la diligencia jurídica con la participación de las partes... [Introducción detallada del contexto legal].

INTERVINIENTES Y PARTES
1. [Nombre completo] - [Calidad jurídica, ej. Abogado Demandante / Asesor / Perito]
(Lista detallada de los sujetos procesales o participantes identificados en el audio).

CONSIDERACIONES Y ANÁLISIS JURÍDICO
PUNTO PRIMERO: [ASUNTO O HECHO JURÍDICO TRATADO]
[Fundamentación jurídica detallada, análisis probatorio, exposición de normas, jurisprudencia o argumentos debatidos durante la sesión].
DECISIONES Y RESOLUCIONES: [Resoluciones adoptadas, medidas cautelares o conclusiones jurídicas específicas].
DIRECTRICES OBLIGATORIAS: [Plazos procesales, recursos interpuestos o mandatos fijados, o "Ninguno"].

(Repetir la estructura para cada punto o etapa procesal tratada).

RESUMEN DE DISPOSICIONES Y RESOLUCIONES FINALES
1. [Resumen de resolución 1]
2. [Resumen de resolución 2]

Finalizada la diligencia, se firma conforme a derecho a las [Hora de Cierre].

FIRMAS
_____________________________
[Nombre real del Juez, Árbitro o Director de la Audiencia]
Director / Funcionario Competente

_____________________________
[Nombre real de la contraparte o interviniente principal]
Interviniente / Apoderado

REGLAS ESTRICTAS DE REDACCIÓN:
1. EXTRACCIÓN REAL: Utiliza nombres de autoridades, códigos normativos, artículos, montos de condenas o acuerdos citados en el audio.
2. RIGOR JURÍDICO: Mantén una redacción técnica, formal, precisa y exhaustiva propia del derecho.
3. FORMATO LIMPIO: NO utilices asteriscos (*), símbolos de almohadilla (#) ni markdown crudo. Usa texto plano con títulos en MAYÚSCULAS SOSTENIDAS.
"""

PROMPT_SISTEMA_MEDICO = """
Eres un Médico Especialista y Auditor de Calidad en Salud experto en comités médicos, juntas de especialistas y gestión clínica.
Tu objetivo es redactar un resumen clínico y acta de comité médico formal, íntegro y detallado a partir de la transcripción provista (con diarización de voces), asegurando el rigor ético y científico.

ESTRUCTURA OBLIGATORIA QUE DEBES GENERAR (SIN MODIFICAR EL ORDEN):

ACTA DE COMITÉ MÉDICO Y GESTIÓN CLÍNICA
COMITÉ MÉDICO INSTITUCIONAL - [Nombre de la Institución o Clínica extraída del audio]
En la ciudad de [Ciudad], a [Fecha], siendo las [Hora], se instala el comité médico con la participación del equipo asistencial... [Introducción clínica formal].

PROFESIONALES ASISTENTES
1. [Nombre completo] - [Especialidad o Cargo, ej. Médico Internista / Cirujano / Auditor]
(Lista detallada de los profesionales de la salud presentes en la transcripción).

CASOS Y PUNTOS CLÍNICOS TRATADOS
PUNTO PRIMERO: [CASO CLÍNICO O ASUNTO ADMINISTRATIVO]
[Descripción detallada de la condición del paciente, antecedentes, hallazgos de diagnósticos, imágenes o laboratorios discutidos en la sesión].
CONCLUSIONES CLÍNICAS: [Criterio médico colegiado, diagnóstico definitivo o ajustes en el tratamiento].
PLAN DE MANEJO Y SEGUIMIENTO: [Órdenes médicas, interconsultas, procedimientos programados o "Ninguno"].

(Repetir la estructura de CONCLUSIONES y PLAN DE MANEJO para cada caso o punto abordado).

RESUMEN DE DIRECTRICES Y PLANES DE TRATAMIENTO
1. [Resumen de directriz clínica 1]
2. [Resumen de directriz clínica 2]

Concluida la sesión clínica a las [Hora de Cierre], firman los profesionales intervinientes.

FIRMAS
_____________________________
[Nombre real del Médico Coordinador o Jefe de Comité]
Coordinador del Comité Médico

_____________________________
[Nombre real del Médico o Especialista participante]
Especialista / Asistente

REGLAS ESTRICTAS DE REDACCIÓN:
1. EXTRACCIÓN REAL: Utiliza nombres de pacientes (o códigos anonimizados si se usaron), patologías, medicamentos, dosis y procedimientos mencionados en el audio.
2. RIGOR CIENTÍFICO: Mantén un lenguaje clínico, técnico, objetivo y estricto.
3. FORMATO LIMPIO: NO utilices asteriscos (*), símbolos de almohadilla (#) ni markdown crudo. Usa texto plano con títulos en MAYÚSCULAS SOSTENIDAS.
"""

PROMPT_SISTEMA_EDUCATIVO = """
Eres un Académico y Secretario de Consejo Educativo experto en gestión institucional y normatividad educativa.
Tu objetivo es redactar un acta de consejo académico o reunión pedagógica formal, íntegra y detallada a partir de la transcripción provista (con diarización de voces), siguiendo la estructura oficial de instituciones educativas.

ESTRUCTURA OBLIGATORIA QUE DEBES GENERAR (SIN MODIFICAR EL ORDEN):

ACTA DE CONSEJO ACADÉMICO Y REUNIÓN EDUCATIVA
CONSEJO ACADÉMICO DE LA INSTITUCIÓN [Nombre del Colegio o Universidad extraído del audio]
En la ciudad de [Ciudad], a [Fecha], siendo las [Hora], se reúnen los miembros del Consejo Académico... [Introducción formal de la reunión institucional].

MIEMBROS Y ASISTENTES
1. [Nombre completo] - [Cargo o Rol, ej. Rector / Coordinador Académico / Docente]
(Lista detallada de directivos, docentes o representantes asistentes).
Se constata el quórum reglamentario para deliberar.

ORDEN DEL DÍA
1. [Primer punto tratado]
2. [Segundo punto tratado]
(Y así sucesivamente según los puntos reales de la agenda).

DESARROLLO DE LA REUNIÓN
PUNTO PRIMERO: [TÍTULO DEL PUNTO]
[Narrativa detallada del análisis pedagógico, rendimiento estudiantil, curricular o administrativo debatido, mencionando intervenciones específicas].
ACUERDOS Y APROBACIONES: [Decisiones tomadas por votación o consenso sobre el punto].
COMPROMISOS ACADÉMICOS: [Tareas, responsables y plazos asignados, o "Ninguno"].

(Repetir la estructura de ACUERDOS y COMPROMISOS para cada punto tratado).

RESUMEN GENERAL DE ACUERDOS Y TAREAS PENDIENTES
1. [Resumen consolidado de acuerdo 1]
2. [Resumen consolidado de acuerdo 2]

Agotado el orden del día, se da por terminada la sesión a las [Hora de Cierre], firmando los asistentes en señal de conformidad.

FIRMAS
_____________________________
[Nombre real del Rector o Presidente del Consejo]
Rector / Presidente del Consejo Académico

_____________________________
[Nombre real del Secretario(a)]
Secretario(a) Académico(a)

REGLAS ESTRICTAS DE REDACCIÓN:
1. EXTRACCIÓN REAL: Utiliza nombres de profesores, programas, asignaturas, porcentajes o normativas institucionales discutidas en el audio.
2. RIGOR INSTITUCIONAL: Mantén una redacción formal, pedagógica, clara y detallada.
3. FORMATO LIMPIO: NO utilices asteriscos (*), símbolos de almohadilla (#) ni markdown crudo. Usa texto plano con títulos en MAYÚSCULAS SOSTENIDAS.
"""


def get_r2_client():
    return boto3.client(
        's3',
        endpoint_url=R2_ENDPOINT_URL,
        aws_access_key_id=R2_ACCESS_KEY_ID,
        aws_secret_access_key=R2_SECRET_ACCESS_KEY,
        config=Config(signature_version='s3v4'),
        region_name='auto'
    )


# ==========================================
# 2. TAREA: PROCESAR ASAMBLEA (AUDIO -> DOCX)
# ==========================================
@celery_app.task(
    bind=True,
    max_retries=3,
    default_retry_delay=15,
    autoretry_for=(requests.RequestException, openai.APIError)
)
def task_procesar_asamblea(self, temp_audio_path: str, email: str, instrucciones: str, nombre_personalizado: str, original_filename: str, motor: str = "asamblea_ph"):
    try:
        celery_task_id = self.request.id
        self.update_state(state="PROCESSING", meta={"status": "Procesando audio e identificando oradores desde la nube..."})

        # Descarga con stream=True para calcular el hash por bloques sin saturar la RAM
        response_audio = requests.get(temp_audio_path, stream=True, timeout=60)
        if response_audio.status_code != 200:
            raise Exception(f"No se pudo descargar el archivo de audio desde la nube: {temp_audio_path}")
            
        # Cálculo eficiente del hash SHA-256 mediante streaming (Bloques de 8KB)
        sha256_hash = hashlib.sha256()
        for chunk in response_audio.iter_content(chunk_size=8192):
            if chunk:
                sha256_hash.update(chunk)
                
        file_hash = sha256_hash.hexdigest()

        # PROTECCIÓN DE IDEMPOTENCIA: Verificar si la tarea ya fue completada exitosamente
        acta_existente = actas_collection.find_one({"celery_task_id": celery_task_id})
        if acta_existente and acta_existente.get("estado") == "COMPLETED":
            return {
                "status": "COMPLETED",
                "acta_id": str(acta_existente["_id"]),
                "nombre_acta": acta_existente["nombre_acta"],
                "file_url": acta_existente["file_url"]
            }

        cached = transripciones_collection.find_one({"file_hash": file_hash})
        duracion_segundos = 0

        if cached:
            texto_transcrito = cached["texto_transcrito"]
            duracion_segundos = cached.get("duracion_segundos", 300)
        else:
            config = aai.TranscriptionConfig(speaker_labels=True, language_code="es")
            transcriber = aai.Transcriber()
            
            transcript = transcriber.transcribe(temp_audio_path, config=config)

            if transcript.status == aai.TranscriptStatus.error:
                raise Exception(f"Error en AssemblyAI: {transcript.error}")

            # Calcular duración en segundos
            duracion_segundos = getattr(transcript, 'audio_duration', 300) or 300

            texto_transcrito = ""
            if transcript.utterances:
                for utterance in transcript.utterances:
                    texto_transcrito += f"[Persona {utterance.speaker}]: {utterance.text}\n"
            else:
                texto_transcrito = transcript.text

            transripciones_collection.insert_one({
                "file_hash": file_hash,
                "filename": original_filename,
                "texto_transcrito": texto_transcrito,
                "duracion_segundos": duracion_segundos,
                "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "createdAt": datetime.now(timezone.utc),
            })

        self.update_state(state="PROCESSING", meta={"status": "Generando el Documento Especializado con IA..."})

        session_id = str(uuid.uuid4())
        
        # Seleccionar el prompt, el sufijo del archivo y el título dinámico según el motor recibido
        if motor == "corporativo":
            prompt_base = PROMPT_SISTEMA_CORPORATIVO
            sufijo_nombre = "Junta_Directiva"
            titulo_documento = "ACTA DE JUNTA DIRECTIVA Y REUNIÓN CORPORATIVA"
        elif motor == "legal_abogados":
            prompt_base = PROMPT_SISTEMA_LEGAL
            sufijo_nombre = "Dictamen_Legal"
            titulo_documento = "DICTAMEN Y ACTA DE AUDIENCIA / ASESORÍA LEGAL"
        elif motor == "medico_salud":
            prompt_base = PROMPT_SISTEMA_MEDICO
            sufijo_nombre = "Comite_Medico"
            titulo_documento = "ACTA DE COMITÉ MÉDICO Y GESTIÓN CLÍNICA"
        elif motor == "educativo_academico":
            prompt_base = PROMPT_SISTEMA_EDUCATIVO
            sufijo_nombre = "Consejo_Academico"
            titulo_documento = "ACTA DE CONSEJO ACADÉMICO Y REUNIÓN EDUCATIVA"
        else:
            prompt_base = PROMPT_SISTEMA_ACTAS
            sufijo_nombre = "Acta_Asamblea"
            titulo_documento = "ACTA DE ASAMBLEA GENERAL DE COPROPIETARIOS"

        nombre_archivo_acta = f"{sufijo_nombre}_{session_id[:8]}.docx" if not nombre_personalizado else f"{nombre_personalizado.strip().replace(' ', '_')}.docx"

        prompt_final = prompt_base
        if instrucciones:
            prompt_final += f"\n\nINSTRUCCIONES ADICIONALES DEL USUARIO:\n{instrucciones}"

        response = openai_client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": prompt_final},
                {"role": "user", "content": f"Transcripción de la reunión:\n\n{texto_transcrito}"},
            ],
            temperature=0.3,
        )
        acta_final = response.choices[0].message.content

        # Crear documento Word en memoria usando el título dinámico del nicho seleccionado
        doc = Document()
        doc.add_heading(titulo_documento, level=0).alignment = 1
        for linea in acta_final.split("\n"):
            if linea.strip():
                doc.add_paragraph(linea.strip())
        
        doc_io = BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        docx_bytes = doc_io.read()

        # Subir .docx a Cloudflare R2
        r2_docx_key = f"actas_generadas/{session_id}_{nombre_archivo_acta}"
        s3 = get_r2_client()
        s3.put_object(
            Bucket=R2_BUCKET_NAME,
            Key=r2_docx_key,
            Body=docx_bytes,
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        docx_url = f"{R2_PUBLIC_URL.rstrip('/')}/{r2_docx_key}"
        peso_archivo = f"{round(len(docx_bytes) / 1024, 1)} KB"

        # Cálculo de consumo de tiempo (Horas)
        duracion_horas = round(duracion_segundos / 3600.0, 2)
        if duracion_horas <= 0:
            duracion_horas = 0.01

        # Actualizar cuota de usuario en MongoDB
        if email:
            users_collection.update_one(
                {"email": email},
                {
                    "$inc": {
                        "horas_usadas_mes": duracion_horas,
                        "horas_restantes": -duracion_horas
                    }
                }
            )

        # ACTUALIZACIÓN DEL REGISTRO PRE-CREADO EN MONGODB (NO USAR insert_one AQUÍ)
        update_data = {
            "estado": "COMPLETED",
            "file_hash": file_hash,
            "nombre_acta": nombre_archivo_acta,
            "peso": peso_archivo,
            "contenido": acta_final,
            "duracion_horas": duracion_horas,
            "file_url": docx_url,
            "motor": motor,
            "updatedAt": datetime.now(timezone.utc)
        }
        
        actas_collection.update_one(
            {"celery_task_id": celery_task_id},
            {"$set": update_data}
        )

        doc_actualizado = actas_collection.find_one({"celery_task_id": celery_task_id})
        acta_db_id = str(doc_actualizado["_id"]) if doc_actualizado else None

        return {
            "status": "COMPLETED",
            "acta_id": acta_db_id,
            "nombre_acta": nombre_archivo_acta,
            "file_url": docx_url
        }

    except Exception as exc:
        # Marcar estado como FAILED en MongoDB si ocurre un error inesperado
        actas_collection.update_one(
            {"celery_task_id": self.request.id},
            {"$set": {"estado": "FAILED", "error": str(exc)}}
        )
        raise Exception(f"Fallo en la tarea de procesamiento: {str(exc)}")


# ==========================================
# 3. TAREA: ESCANEAR DOCUMENTO (OCR / HTML)
# ==========================================
@celery_app.task(
    bind=True,
    max_retries=3,
    default_retry_delay=10,
    autoretry_for=(openai.APIError, requests.RequestException)
)
def task_escanear_documento(self, file_bytes_b64: str, filename: str, email: str = None):
    try:
        celery_task_id = self.request.id
        self.update_state(state="PROCESSING", meta={"status": "Validando permisos y leyendo archivo..."})

        # PROTECCIÓN DE IDEMPOTENCIA: Verificar si la tarea ya fue completada
        scanner_existente = scanners_historial_collection.find_one({"celery_task_id": celery_task_id})
        if scanner_existente and scanner_existente.get("estado") == "COMPLETED":
            return {
                "status": "COMPLETED",
                "transcripcion": scanner_existente["contenido"],
                "id": str(scanner_existente["_id"])
            }

        # Validar límite de tokens del usuario
        if email:
            usuario = users_collection.find_one({"email": email})
            if usuario:
                tokens_usados = usuario.get("tokens_usados", 0)
                limite_tokens = usuario.get("limite_tokens_mes", 0)
                if limite_tokens > 0 and tokens_usados >= limite_tokens:
                    raise Exception("Has alcanzado el límite de tokens mensuales de tu plan. Actualiza tu suscripción para continuar.")

        file_bytes = base64.b64decode(file_bytes_b64)
        filename_lower = filename.lower()
        
        texto_extraido = ""
        es_imagen = filename_lower.endswith((".png", ".jpg", ".jpeg", ".webp"))
        
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
            if filename_lower.endswith(".pdf"):
                doc = fitz.open(stream=file_bytes, filetype="pdf")
                try:
                    for page_num in range(len(doc)):
                        page = doc[page_num]
                        texto_pagina = page.get_text()
                        
                        if texto_pagina.strip():
                            texto_extraido += f"\n--- Página {page_num + 1} ---\n" + texto_pagina
                finally:
                    doc.close()
                
                # Respaldo con pdfplumber si la lectura simple extrae poco texto
                if len(texto_extraido.strip()) < 50:
                    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                        for i, page in enumerate(pdf.pages):
                            t = page.extract_text()
                            if t:
                                texto_extraido += f"\n--- Página (Tablas) {i + 1} ---\n" + t

            elif filename_lower.endswith((".txt", ".doc", ".docx")):
                texto_extraido = file_bytes.decode("utf-8", errors="ignore")
            else:
                raise Exception("Formato de archivo no soportado. Sube un PDF, imagen o documento de texto.")

            if not texto_extraido.strip():
                raise Exception("El documento está vacío o no se pudo extraer texto legible.")

            contenido_usuario = f"""Analyze the following text extracted from the document. Your output must be EXCLUSIVELY corporate semantic HTML ready to render directly in a browser or web container.
- Replicate the original visual and section structure.
- Use <h1>, <h2> for main and section titles.
- Use <p> for paragraphs with Tailwind classes (e.g., text-slate-900, text-xs, leading-relaxed).
- Use complete table tags (<table>, <thead>, <tbody>, <tr>, <th>, <td>) with borders and corporate classes if there is structured data.
- If you detect references to charts, schemes, or diagrams, create them as a visual block with a dotted border.
- FORBIDDEN to use Markdown, asterisks (*), markdown list hyphens (#), or wrap the result in markdown code quotes.

Extracted text:
{texto_extraido[:15000]}"""

        self.update_state(state="PROCESSING", meta={"status": "Procesando HTML estructurado con GPT-4o..."})

        response_openai = openai_client.chat.completions.create(
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

        # Descontar tokens utilizados
        tokens_consumidos = 0
        if email and hasattr(response_openai, "usage") and response_openai.usage:
            tokens_consumidos = response_openai.usage.total_tokens
            users_collection.update_one(
                {"email": email},
                {"$inc": {"tokens_usados": tokens_consumidos}}
            )

        # Limpieza defensiva de tags de Markdown
        if resultado_html.startswith("```html"):
            resultado_html = resultado_html[7:]
        if resultado_html.startswith("```"):
            resultado_html = resultado_html[3:]
        if resultado_html.endswith("```"):
            resultado_html = resultado_html[:-3]
        resultado_html = resultado_html.strip()

        # ACTUALIZACIÓN DEL REGISTRO PRE-CREADO EN HISTORIAL DE SCANNERS (NO USAR insert_one AQUÍ)
        scanner_id = None
        if email:
            scanners_historial_collection.update_one(
                {"celery_task_id": celery_task_id},
                {
                    "$set": {
                        "estado": "COMPLETED",
                        "tokens": tokens_consumidos,
                        "contenido": resultado_html,
                        "updatedAt": datetime.now(timezone.utc)
                    }
                }
            )
            scanner_doc = scanners_historial_collection.find_one({"celery_task_id": celery_task_id})
            if scanner_doc:
                scanner_id = str(scanner_doc["_id"])

        return {
            "status": "COMPLETED",
            "transcripcion": resultado_html,
            "id": scanner_id
        }

    except Exception as e:
        # Marcar estado como FAILED en MongoDB si ocurre un error inesperado
        if email:
            scanners_historial_collection.update_one(
                {"celery_task_id": self.request.id},
                {"$set": {"estado": "FAILED", "error": str(e)}}
            )
        raise Exception(f"Error procesando el archivo: {str(e)}")
