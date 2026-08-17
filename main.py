from __future__ import annotations

import base64
import html
import json
import os
import re
import traceback
import uuid
import time
from datetime import datetime
from email.message import EmailMessage
from functools import wraps
from typing import Any
from pathlib import PurePosixPath
from urllib.parse import parse_qs, quote, unquote, urlparse
from zoneinfo import ZoneInfo

import requests
from flask import Flask, g, jsonify, request
from google.auth.transport.requests import Request as GoogleAuthRequest
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaInMemoryUpload


app = Flask(__name__)


def registrar_tiempo(**campos: Any) -> None:
    """Escribe una medición estructurada en stdout para Cloud Logging.

    Se usa print(..., flush=True) para no depender del nivel INFO configurado
    por Flask o Gunicorn. Cloud Run interpreta el JSON como jsonPayload.
    """
    payload: dict[str, Any] = {
        "severity": "INFO",
        "message": "TIEMPO",
        "marca": "TIEMPO",
        "fecha_chile": ahora_iso() if "ahora_iso" in globals() else "",
        **campos,
    }

    try:
        payload.setdefault("solicitud_id", getattr(g, "solicitud_id", ""))
        payload.setdefault("endpoint", request.path)
        payload.setdefault("metodo", request.method)

        trace_header = request.headers.get("X-Cloud-Trace-Context", "")
        trace_id = trace_header.split("/", 1)[0].strip()
        project_id = (
            os.environ.get("GOOGLE_CLOUD_PROJECT")
            or os.environ.get("GCP_PROJECT")
            or ""
        ).strip()
        if trace_id and project_id:
            payload["logging.googleapis.com/trace"] = (
                f"projects/{project_id}/traces/{trace_id}"
            )
    except RuntimeError:
        # Permite reutilizar el helper fuera de un contexto Flask.
        pass

    print(json.dumps(payload, ensure_ascii=False, default=str), flush=True)


def medir_operacion(nombre: str):
    """Registra en Cloud Run la duración de una operación externa o costosa."""

    def decorador(func):
        @wraps(func)
        def envoltura(*args, **kwargs):
            inicio = time.perf_counter()
            estado = "ok"
            try:
                return func(*args, **kwargs)
            except Exception:
                estado = "error"
                raise
            finally:
                registrar_tiempo(
                    categoria="operacion",
                    operacion=nombre,
                    estado=estado,
                    duracion_ms=round(
                        (time.perf_counter() - inicio) * 1000,
                        1,
                    ),
                )

        return envoltura

    return decorador


@app.before_request
def iniciar_medicion_solicitud() -> None:
    g.inicio_solicitud = time.perf_counter()
    g.solicitud_id = uuid.uuid4().hex[:12]


@app.after_request
def finalizar_medicion_solicitud(response):
    inicio = getattr(g, "inicio_solicitud", None)
    if inicio is None:
        return response

    duracion_ms = (time.perf_counter() - inicio) * 1000
    registrar_tiempo(
        categoria="endpoint",
        endpoint=request.path,
        metodo=request.method,
        status=response.status_code,
        duracion_ms=round(duracion_ms, 1),
    )
    response.headers["Server-Timing"] = f"total;dur={duracion_ms:.1f}"
    response.headers["X-Solicitud-ID"] = getattr(g, "solicitud_id", "")
    return response


# -----------------------------------------------------------------------------
# Configuración
# -----------------------------------------------------------------------------

APPSHEET_APP_ID = os.environ.get("APPSHEET_APP_ID")
APPSHEET_ACCESS_KEY = os.environ.get("APPSHEET_ACCESS_KEY")
APPSHEET_DOMAIN = os.environ.get("APPSHEET_DOMAIN", "www.appsheet.com")

TABLA_PLANTILLAS = os.environ.get(
    "TABLA_PLANTILLAS",
    "Plantillas_Documentos",
)

TABLA_APROBADORES = os.environ.get(
    "TABLA_APROBADORES",
    "Documentos_Aprobadores",
)

TABLA_APROBADORES_ACTUAL = os.environ.get(
    "TABLA_APROBADORES_ACTUAL",
    "Documentos_Aprobadores_Actual",
)

TABLA_DOCUMENTOS = os.environ.get(
    "TABLA_DOCUMENTOS",
    "Documentos",
)

TABLA_VERSIONES = os.environ.get(
    "TABLA_VERSIONES",
    "Documento_Versiones",
)

TABLA_EVENTOS = os.environ.get(
    "TABLA_EVENTOS",
    "Documento_Eventos",
)

TABLA_NOTIFICACIONES = os.environ.get(
    "TABLA_NOTIFICACIONES",
    "Documento_Notificaciones",
)

APPSHEET_DOCUMENT_VIEW_URL = os.environ.get(
    "APPSHEET_DOCUMENT_VIEW_URL",
    "",
).strip()

NOMBRE_APLICACION = (
    os.environ.get("NOMBRE_APLICACION")
    or os.environ.get("Nombre_Aplicacion")
    or "Gestión documental"
).strip()

WEBHOOK_TOKEN = os.environ.get("WEBHOOK_TOKEN")

GOOGLE_OAUTH_CLIENT_ID = os.environ.get("GOOGLE_OAUTH_CLIENT_ID")
GOOGLE_OAUTH_CLIENT_SECRET = os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET")
GOOGLE_OAUTH_REFRESH_TOKEN = os.environ.get("GOOGLE_OAUTH_REFRESH_TOKEN")
GOOGLE_OAUTH_TOKEN_URI = os.environ.get(
    "GOOGLE_OAUTH_TOKEN_URI",
    "https://oauth2.googleapis.com/token",
)

DRIVE_SEND_NOTIFICATION_EMAIL = os.environ.get(
    "DRIVE_SEND_NOTIFICATION_EMAIL",
    "false",
).strip().lower() in {"1", "true", "yes", "si", "sí"}

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/gmail.send",
]

DOCX_MIME_TYPE = (
    "application/"
    "vnd.openxmlformats-officedocument.wordprocessingml.document"
)

GMAIL_SENDER_EMAIL = os.environ.get(
    "GMAIL_SENDER_EMAIL",
    "gestion.enfoke@gmail.com",
).strip()

CHILE_TZ = ZoneInfo("America/Santiago")


# -----------------------------------------------------------------------------
# Utilidades generales
# -----------------------------------------------------------------------------


def ahora_iso() -> str:
    """
    Devuelve un DateTime compatible con la API de AppSheet cuando
    Properties.Locale está configurado como es-CL.
    """
    return datetime.now(CHILE_TZ).strftime("%d/%m/%Y %H:%M:%S")


def nuevo_id() -> str:
    return str(uuid.uuid4())


def texto(valor: Any) -> str:
    """
    Convierte valores de AppSheet a texto simple.

    AppSheet puede devolver columnas URL como un objeto:
    {"Url": "...", "LinkText": "..."}.
    Al volver a enviar ese objeto a una columna de tipo Url, la API lo
    rechaza. Esta función extrae únicamente la URL real.
    """
    if valor is None:
        return ""

    if isinstance(valor, dict):
        for clave in ("Url", "URL", "url"):
            contenido = valor.get(clave)
            if contenido is not None:
                return str(contenido).strip()

        # Otros valores enriquecidos de AppSheet pueden traer Value.
        for clave in ("Value", "value"):
            contenido = valor.get(clave)
            if contenido is not None:
                return str(contenido).strip()

        return ""

    return str(valor).strip()


def normalizar_url_appsheet(valor: Any) -> str:
    """
    Devuelve una URL plana aunque AppSheet la entregue como:

    - dict: {"Url": "...", "LinkText": "..."}
    - string JSON: '{"Url":"...","LinkText":"..."}'
    - URL normal: "https://..."

    Si recibe cualquier otro texto, intenta extraer la primera URL http/https.
    """
    if valor is None:
        return ""

    if isinstance(valor, dict):
        for clave in ("Url", "URL", "url", "Value", "value"):
            if clave in valor:
                return normalizar_url_appsheet(valor.get(clave))
        return ""

    if isinstance(valor, (list, tuple)):
        for item in valor:
            url = normalizar_url_appsheet(item)
            if url:
                return url
        return ""

    valor_texto = str(valor).strip()
    if not valor_texto:
        return ""

    # AppSheet puede devolver el objeto URL serializado como texto JSON.
    if valor_texto.startswith("{") and valor_texto.endswith("}"):
        try:
            objeto = json.loads(valor_texto)
        except json.JSONDecodeError:
            objeto = None

        if objeto is not None:
            url = normalizar_url_appsheet(objeto)
            if url:
                return url

    if valor_texto.lower().startswith(("https://", "http://")):
        return valor_texto

    coincidencia = re.search(r'https?://[^"\'\s}<]+', valor_texto)
    if coincidencia:
        return coincidencia.group(0)

    return ""


def entero(valor: Any, nombre_campo: str) -> int:
    try:
        return int(float(str(valor).strip()))
    except (TypeError, ValueError) as exc:
        raise ValueError(
            f"{nombre_campo} debe contener un número entero. Valor recibido: {valor!r}"
        ) from exc


def es_verdadero(valor: Any) -> bool:
    if isinstance(valor, bool):
        return valor
    return texto(valor).lower() in {
        "true",
        "verdadero",
        "yes",
        "si",
        "sí",
        "1",
        "y",
    }


def literal_appsheet(valor: str) -> str:
    """Devuelve un literal de texto seguro para expresiones Selector de AppSheet."""
    return json.dumps(valor, ensure_ascii=False)


def limpiar_nombre_archivo(nombre: str) -> str:
    nombre = re.sub(r"[\r\n\t]+", " ", nombre).strip()
    nombre = nombre.replace("/", "-").replace("\\", "-")
    nombre = re.sub(r"\s+", " ", nombre)
    return nombre[:180] or "Documento"


def validar_configuracion() -> None:
    faltantes: list[str] = []

    variables_obligatorias = {
        "APPSHEET_APP_ID": APPSHEET_APP_ID,
        "APPSHEET_ACCESS_KEY": APPSHEET_ACCESS_KEY,
        "WEBHOOK_TOKEN": WEBHOOK_TOKEN,
        "GOOGLE_OAUTH_CLIENT_ID": GOOGLE_OAUTH_CLIENT_ID,
        "GOOGLE_OAUTH_CLIENT_SECRET": GOOGLE_OAUTH_CLIENT_SECRET,
        "GOOGLE_OAUTH_REFRESH_TOKEN": GOOGLE_OAUTH_REFRESH_TOKEN,
        "APPSHEET_DOCUMENT_VIEW_URL": APPSHEET_DOCUMENT_VIEW_URL,
    }

    for nombre, valor in variables_obligatorias.items():
        if not valor:
            faltantes.append(nombre)

    if faltantes:
        raise RuntimeError(
            "Faltan variables de entorno: " + ", ".join(faltantes)
        )


def validar_token() -> None:
    token_recibido = request.headers.get("X-Webhook-Token", "")
    if token_recibido != WEBHOOK_TOKEN:
        raise PermissionError("Token de webhook inválido")


# -----------------------------------------------------------------------------
# Google Drive
# -----------------------------------------------------------------------------


@medir_operacion("google.oauth.refresh")
def obtener_google_credentials() -> Credentials:
    credentials = Credentials(
        token=None,
        refresh_token=GOOGLE_OAUTH_REFRESH_TOKEN,
        token_uri=GOOGLE_OAUTH_TOKEN_URI,
        client_id=GOOGLE_OAUTH_CLIENT_ID,
        client_secret=GOOGLE_OAUTH_CLIENT_SECRET,
        scopes=GOOGLE_SCOPES,
    )

    credentials.refresh(GoogleAuthRequest())
    return credentials


def obtener_drive_service():
    return build(
        "drive",
        "v3",
        credentials=obtener_google_credentials(),
        cache_discovery=False,
    )


def obtener_gmail_service():
    return build(
        "gmail",
        "v1",
        credentials=obtener_google_credentials(),
        cache_discovery=False,
    )


@medir_operacion("drive.copiar_plantilla")
def copiar_plantilla(
    drive_service: Any,
    template_id: str,
    folder_id: str,
    nombre_documento: str,
) -> dict[str, str]:
    metadata = {
        "name": nombre_documento,
        "parents": [folder_id],
    }

    copia = (
        drive_service.files()
        .copy(
            fileId=template_id,
            body=metadata,
            fields="id,name,webViewLink",
            supportsAllDrives=True,
        )
        .execute()
    )

    google_doc_id = copia["id"]
    google_doc_url = copia.get(
        "webViewLink",
        f"https://docs.google.com/document/d/{google_doc_id}/edit",
    )

    return {
        "id": google_doc_id,
        "url": google_doc_url,
        "name": copia.get("name", nombre_documento),
    }


@medir_operacion("drive.asegurar_permiso")
def asegurar_permiso_rol(
    drive_service: Any,
    file_id: str,
    email: str,
    role: str,
) -> str:
    """Crea o actualiza el permiso directo de un usuario sobre un archivo."""
    roles_validos = {"reader", "commenter", "writer"}
    if role not in roles_validos:
        raise ValueError(f"Rol de Drive no válido: {role!r}")

    if not email or "@" not in email:
        raise ValueError(f"Correo de responsable no válido: {email!r}")

    permisos = (
        drive_service.permissions()
        .list(
            fileId=file_id,
            fields="permissions(id,emailAddress,role,type)",
            supportsAllDrives=True,
        )
        .execute()
        .get("permissions", [])
    )

    for permiso in permisos:
        if texto(permiso.get("emailAddress")).lower() != email.lower():
            continue

        permission_id = texto(permiso.get("id"))
        rol_actual = texto(permiso.get("role")).lower()

        # No se puede degradar al propietario mediante esta operación.
        if rol_actual in {"owner", "organizer", "fileorganizer"}:
            return permission_id

        if rol_actual == role:
            return permission_id

        permiso_actualizado = (
            drive_service.permissions()
            .update(
                fileId=file_id,
                permissionId=permission_id,
                body={"role": role},
                fields="id",
                supportsAllDrives=True,
            )
            .execute()
        )
        return texto(permiso_actualizado.get("id"))

    permiso_creado = (
        drive_service.permissions()
        .create(
            fileId=file_id,
            body={
                "type": "user",
                "role": role,
                "emailAddress": email,
            },
            fields="id",
            sendNotificationEmail=DRIVE_SEND_NOTIFICATION_EMAIL,
            supportsAllDrives=True,
        )
        .execute()
    )

    return texto(permiso_creado.get("id"))


def asegurar_permiso_editor(
    drive_service: Any,
    file_id: str,
    email: str,
) -> str:
    return asegurar_permiso_rol(
        drive_service=drive_service,
        file_id=file_id,
        email=email,
        role="writer",
    )


def escapar_consulta_drive(valor: str) -> str:
    return valor.replace("\\", "\\\\").replace("'", "\\'")


@medir_operacion("drive.buscar_archivo")
def buscar_archivo_en_carpeta(
    drive_service: Any,
    folder_id: str,
    nombre_archivo: str,
) -> dict[str, str] | None:
    """Busca un archivo exacto para reutilizarlo tras un reintento."""
    nombre_q = escapar_consulta_drive(nombre_archivo)
    folder_q = escapar_consulta_drive(folder_id)
    consulta = (
        f"name = '{nombre_q}' and "
        f"'{folder_q}' in parents and trashed = false"
    )

    archivos = (
        drive_service.files()
        .list(
            q=consulta,
            fields="files(id,name,webViewLink)",
            spaces="drive",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            pageSize=10,
        )
        .execute()
        .get("files", [])
    )

    if len(archivos) > 1:
        raise RuntimeError(
            f"Existen varios archivos con el nombre {nombre_archivo!r} "
            "en la carpeta destino. Elimina los duplicados antes de continuar."
        )

    if not archivos:
        return None

    archivo = archivos[0]
    file_id = texto(archivo.get("id"))
    return {
        "id": file_id,
        "name": texto(archivo.get("name")) or nombre_archivo,
        "url": texto(archivo.get("webViewLink"))
        or f"https://docs.google.com/document/d/{file_id}/edit",
    }


def copiar_archivo_o_reutilizar(
    drive_service: Any,
    source_file_id: str,
    folder_id: str,
    nombre_archivo: str,
) -> dict[str, str]:
    existente = buscar_archivo_en_carpeta(
        drive_service=drive_service,
        folder_id=folder_id,
        nombre_archivo=nombre_archivo,
    )
    if existente:
        return existente

    return copiar_plantilla(
        drive_service=drive_service,
        template_id=source_file_id,
        folder_id=folder_id,
        nombre_documento=nombre_archivo,
    )


# -----------------------------------------------------------------------------
# AppSheet API
# -----------------------------------------------------------------------------


def normalizar_respuesta_appsheet(data: Any) -> list[dict[str, Any]]:
    if data is None:
        return []

    if isinstance(data, list):
        return [fila for fila in data if isinstance(fila, dict)]

    if isinstance(data, dict):
        filas = data.get("Rows")
        if isinstance(filas, list):
            return [fila for fila in filas if isinstance(fila, dict)]
        return [data]

    raise RuntimeError(
        f"Respuesta inesperada de AppSheet: {type(data).__name__}"
    )


def appsheet_action(
    table_name: str,
    action: str,
    rows: list[dict[str, Any]] | None = None,
    selector: str | None = None,
) -> list[dict[str, Any]]:
    table_encoded = quote(table_name, safe="")
    url = (
        f"https://{APPSHEET_DOMAIN}/api/v2/apps/"
        f"{APPSHEET_APP_ID}/tables/{table_encoded}/Action"
    )

    headers = {
        "ApplicationAccessKey": APPSHEET_ACCESS_KEY,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

    properties: dict[str, Any] = {
        "Locale": "es-CL",
        "Location": "-33.4489,-70.6693",
        "Timezone": "America/Santiago",
    }

    if selector:
        properties["Selector"] = selector

    payload = {
        "Action": action,
        "Properties": properties,
        "Rows": rows or [],
    }

    inicio_appsheet = time.perf_counter()
    try:
        response = requests.post(
            url,
            headers=headers,
            json=payload,
            timeout=90,
        )
    except Exception:
        registrar_tiempo(
            categoria="servicio",
            servicio="appsheet",
            tabla=table_name,
            accion=action,
            filas=len(rows or []),
            estado="error",
            duracion_ms=round(
                (time.perf_counter() - inicio_appsheet) * 1000,
                1,
            ),
        )
        raise

    registrar_tiempo(
        categoria="servicio",
        servicio="appsheet",
        tabla=table_name,
        accion=action,
        filas=len(rows or []),
        status=response.status_code,
        estado="ok" if response.status_code == 200 else "error",
        duracion_ms=round(
            (time.perf_counter() - inicio_appsheet) * 1000,
            1,
        ),
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Error AppSheet en {table_name} ({action}): "
            f"{response.status_code} - {response.text}"
        )

    if not response.text.strip():
        return []

    try:
        return normalizar_respuesta_appsheet(response.json())
    except ValueError as exc:
        raise RuntimeError(
            f"AppSheet respondió un contenido no JSON en {table_name}: "
            f"{response.text[:500]}"
        ) from exc


def appsheet_find(
    table_name: str,
    selector: str,
) -> list[dict[str, Any]]:
    return appsheet_action(
        table_name=table_name,
        action="Find",
        rows=[],
        selector=selector,
    )


def buscar_documento(id_documento: str) -> dict[str, Any]:
    selector = (
        f"FILTER({TABLA_DOCUMENTOS}, "
        f"[ID_DOCUMENTO] = {literal_appsheet(id_documento)})"
    )
    filas = appsheet_find(TABLA_DOCUMENTOS, selector)

    if not filas:
        raise LookupError(
            f"No se encontró ID_DOCUMENTO={id_documento} en {TABLA_DOCUMENTOS}"
        )

    return filas[0]


def buscar_plantilla(id_plantilla: str) -> dict[str, Any]:
    selector = (
        f"FILTER({TABLA_PLANTILLAS}, "
        f"[ID_PLANTILLA] = {literal_appsheet(id_plantilla)})"
    )
    filas = appsheet_find(TABLA_PLANTILLAS, selector)

    if not filas:
        raise LookupError(
            f"No se encontró ID_PLANTILLA={id_plantilla} en {TABLA_PLANTILLAS}"
        )

    plantilla = filas[0]
    if not es_verdadero(plantilla.get("ACTIVA")):
        raise ValueError("La plantilla seleccionada no está activa")

    return plantilla


def buscar_cadena_plantilla(id_plantilla: str) -> list[dict[str, Any]]:
    """
    Busca todas las filas asociadas a la plantilla y filtra VIGENTE en
    Python. Esto evita depender de cómo AppSheet serializa un Yes/No.
    """
    selector = (
        f"FILTER({TABLA_APROBADORES}, "
        f"[ID_PLANTILLA] = {literal_appsheet(id_plantilla)})"
    )

    filas_plantilla = appsheet_find(TABLA_APROBADORES, selector)

    if not filas_plantilla:
        raise ValueError(
            "No existen aprobadores asociados a la plantilla "
            f"{id_plantilla!r} en {TABLA_APROBADORES}"
        )

    filas_vigentes = [
        fila
        for fila in filas_plantilla
        if es_verdadero(fila.get("VIGENTE"))
    ]

    if not filas_vigentes:
        valores_vigente = sorted(
            {
                texto(fila.get("VIGENTE")) or "<vacío>"
                for fila in filas_plantilla
            }
        )
        raise ValueError(
            "La plantilla tiene aprobadores, pero ninguno está vigente. "
            f"ID_PLANTILLA={id_plantilla!r}; "
            f"valores encontrados en VIGENTE={valores_vigente}"
        )

    filas_ordenadas = sorted(
        filas_vigentes,
        key=lambda fila: entero(fila.get("ORDEN"), "ORDEN"),
    )

    ordenes = [
        entero(fila.get("ORDEN"), "ORDEN")
        for fila in filas_ordenadas
    ]

    if len(ordenes) != len(set(ordenes)):
        raise ValueError(
            "La cadena de aprobación tiene dos o más responsables "
            "vigentes con el mismo ORDEN"
        )

    return filas_ordenadas


def buscar_versiones_documento(id_documento: str) -> list[dict[str, Any]]:
    selector = (
        f"FILTER({TABLA_VERSIONES}, "
        f"[ID_DOCUMENTO] = {literal_appsheet(id_documento)})"
    )
    return appsheet_find(TABLA_VERSIONES, selector)


def marcar_documento_error(id_documento: str, mensaje: str) -> None:
    try:
        appsheet_action(
            TABLA_DOCUMENTOS,
            "Edit",
            [
                {
                    "ID_DOCUMENTO": id_documento,
                    "ESTADO": "Error",
                    "OBSERVACION_ACTUAL": mensaje[:1000],
                    "FECHA_ULTIMA_ACTUALIZACION": ahora_iso(),
                }
            ],
        )
    except Exception:
        traceback.print_exc()



def registrar_error_transicion(id_documento: str, mensaje: str) -> None:
    """Registra el error sin cambiar a ciegas el estado vigente."""
    try:
        appsheet_action(
            TABLA_DOCUMENTOS,
            "Edit",
            [
                {
                    "ID_DOCUMENTO": id_documento,
                    "OBSERVACION_ACTUAL": mensaje[:1000],
                    "FECHA_ULTIMA_ACTUALIZACION": ahora_iso(),
                }
            ],
        )
    except Exception:
        traceback.print_exc()


def buscar_aprobacion_actual(
    id_aprobacion_actual: str,
) -> dict[str, Any]:
    selector = (
        f"FILTER({TABLA_APROBADORES_ACTUAL}, "
        f"[ID_APROBACION_ACTUAL] = "
        f"{literal_appsheet(id_aprobacion_actual)})"
    )
    filas = appsheet_find(TABLA_APROBADORES_ACTUAL, selector)
    if not filas:
        raise LookupError(
            "No se encontró ID_APROBACION_ACTUAL="
            f"{id_aprobacion_actual}"
        )
    return filas[0]


def buscar_cadena_actual_documento(
    id_documento: str,
    numero_version: int,
) -> list[dict[str, Any]]:
    selector = (
        f"FILTER({TABLA_APROBADORES_ACTUAL}, "
        f"[ID_DOCUMENTO] = {literal_appsheet(id_documento)})"
    )
    filas = appsheet_find(TABLA_APROBADORES_ACTUAL, selector)
    activas = [
        fila
        for fila in filas
        if es_verdadero(fila.get("CADENA_ACTIVA"))
        and entero(fila.get("NUMERO_VERSION"), "NUMERO_VERSION")
        == numero_version
    ]
    return sorted(
        activas,
        key=lambda fila: entero(fila.get("ORDEN"), "ORDEN"),
    )


def buscar_version_por_id(id_version: str) -> dict[str, Any]:
    selector = (
        f"FILTER({TABLA_VERSIONES}, "
        f"[ID_VERSION] = {literal_appsheet(id_version)})"
    )
    filas = appsheet_find(TABLA_VERSIONES, selector)
    if not filas:
        raise LookupError(f"No se encontró ID_VERSION={id_version}")
    return filas[0]


def buscar_version_numero_revision(
    id_documento: str,
    numero_version: int,
    numero_revision: int,
) -> dict[str, Any] | None:
    versiones = buscar_versiones_documento(id_documento)
    coincidentes: list[dict[str, Any]] = []

    for fila in versiones:
        try:
            version_fila = entero(
                fila.get("NUMERO_VERSION"),
                "NUMERO_VERSION",
            )
            revision_fila = entero(
                fila.get("NUMERO_REVISION"),
                "NUMERO_REVISION",
            )
        except ValueError:
            continue

        if (
            version_fila == numero_version
            and revision_fila == numero_revision
        ):
            coincidentes.append(fila)

    if len(coincidentes) > 1:
        raise RuntimeError(
            "Existen varias filas de Documento_Versiones para "
            f"V{numero_version:02d} REV{numero_revision:02d}"
        )

    return coincidentes[0] if coincidentes else None


# -----------------------------------------------------------------------------
# Flujo: creación inicial
# -----------------------------------------------------------------------------


def construir_cadena_actual(
    cadena_plantilla: list[dict[str, Any]],
    id_documento: str,
    id_plantilla: str,
    id_version: str,
    permission_id_drive: str,
    fecha_inicio: str,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    filas_actuales: list[dict[str, Any]] = []
    primer_actual: dict[str, Any] | None = None

    for indice, fila_base in enumerate(cadena_plantilla):
        id_aprobacion_actual = nuevo_id()
        orden = entero(fila_base.get("ORDEN"), "ORDEN")
        aprobador = texto(fila_base.get("APROBADOR"))
        nombre = texto(fila_base.get("NOMBRE"))
        rol_flujo = texto(fila_base.get("ROL_FLUJO"))

        if not aprobador:
            raise ValueError(
                f"El responsable de orden {orden} no tiene APROBADOR"
            )

        fila_actual: dict[str, Any] = {
            "ID_APROBACION_ACTUAL": id_aprobacion_actual,
            "ID_APROBACION_PLANTILLA": texto(
                fila_base.get("ID_APROBACION_PLANTILLA")
            ),
            "ID_PLANTILLA": id_plantilla,
            "ID_DOCUMENTO": id_documento,
            "NUMERO_VERSION": 1,
            "ORDEN": orden,
            "APROBADOR": aprobador,
            "NOMBRE": nombre,
            "ROL_FLUJO": rol_flujo,
            "ESTADO": "En elaboración" if indice == 0 else "Pendiente",
            "CADENA_ACTIVA": True,
        }

        if indice == 0:
            fila_actual.update(
                {
                    "FECHA_INICIO": fecha_inicio,
                    "PERMISSION_ID_DRIVE": permission_id_drive,
                }
            )
            primer_actual = fila_actual

        filas_actuales.append(fila_actual)

    if primer_actual is None:
        raise ValueError("No fue posible identificar al primer encargado")

    return filas_actuales, primer_actual


def crear_registro_version_inicial(
    id_version: str,
    id_documento: str,
    id_aprobacion_responsable: str,
    orden_responsable: int,
    nombre_archivo: str,
    google_doc_id: str,
    google_doc_url: str,
    creado_por: str,
    fecha_creacion: str,
) -> None:
    fila_version = {
        "ID_VERSION": id_version,
        "ID_DOCUMENTO": id_documento,
        "NUMERO_VERSION": 1,
        "NUMERO_REVISION": 0,
        "ETAPA": "Borrador",
        "ESTADO_VERSION": "Activa",
        "NOMBRE_ARCHIVO": nombre_archivo,
        "GOOGLE_DOC_ID": google_doc_id,
        "GOOGLE_DOC_URL": google_doc_url,
        "ID_APROBACION_RESPONSABLE": id_aprobacion_responsable,
        "ORDEN_RESPONSABLE": orden_responsable,
        "MOTIVO_CREACION": "Creación inicial",
        "CREADO_POR": creado_por,
        "FECHA_CREACION": fecha_creacion,
    }

    appsheet_action(TABLA_VERSIONES, "Add", [fila_version])


def actualizar_documento_inicial(
    id_documento: str,
    id_version: str,
    google_doc_id: str,
    google_doc_url: str,
    primer_actual: dict[str, Any],
    fecha_actualizacion: str,
) -> None:
    fila_documento = {
        "ID_DOCUMENTO": id_documento,
        "ESTADO": "Borrador",
        "VERSION_ACTUAL": 1,
        "REVISION_ACTUAL": 0,
        "ID_VERSION_ACTUAL": id_version,
        "GOOGLE_DOC_ID": google_doc_id,
        "GOOGLE_DOC_URL": google_doc_url,
        "ORDEN_ACTUAL": primer_actual["ORDEN"],
        "ID_APROBACION_ACTUAL": primer_actual["ID_APROBACION_ACTUAL"],
        "ENCARGADO_ACTUAL_NOMBRE": primer_actual.get("NOMBRE", ""),
        "ENCARGADO_ACTUAL_EMAIL": primer_actual.get("APROBADOR", ""),
        "ESTADO_FIRMA": "No iniciado",
        "FECHA_ULTIMA_ACTUALIZACION": fecha_actualizacion,
    }

    appsheet_action(TABLA_DOCUMENTOS, "Edit", [fila_documento])


def crear_eventos_iniciales(
    id_documento: str,
    id_version: str,
    id_aprobacion_actual: str,
    usuario: str,
    fecha_evento: str,
    nombre_archivo: str,
    cantidad_responsables: int,
) -> None:
    filas_eventos = [
        {
            "ID_EVENTO": nuevo_id(),
            "ID_DOCUMENTO": id_documento,
            "ID_VERSION": id_version,
            "ID_APROBACION_ACTUAL": id_aprobacion_actual,
            "TIPO_EVENTO": "Documento creado",
            "ESTADO_ANTERIOR": "Creando",
            "ESTADO_NUEVO": "Borrador",
            "USUARIO": usuario,
            "FECHA_EVENTO": fecha_evento,
            "COMENTARIO": "Se inició el flujo documental.",
        },
        {
            "ID_EVENTO": nuevo_id(),
            "ID_DOCUMENTO": id_documento,
            "ID_VERSION": id_version,
            "ID_APROBACION_ACTUAL": id_aprobacion_actual,
            "TIPO_EVENTO": "Cadena creada",
            "ESTADO_ANTERIOR": "",
            "ESTADO_NUEVO": "Activa",
            "USUARIO": usuario,
            "FECHA_EVENTO": fecha_evento,
            "COMENTARIO": (
                f"Se copiaron {cantidad_responsables} responsables desde la plantilla."
            ),
        },
        {
            "ID_EVENTO": nuevo_id(),
            "ID_DOCUMENTO": id_documento,
            "ID_VERSION": id_version,
            "ID_APROBACION_ACTUAL": id_aprobacion_actual,
            "TIPO_EVENTO": "Borrador creado",
            "ESTADO_ANTERIOR": "",
            "ESTADO_NUEVO": "Activa",
            "USUARIO": usuario,
            "FECHA_EVENTO": fecha_evento,
            "COMENTARIO": f"Se creó el archivo {nombre_archivo}.",
        },
    ]

    appsheet_action(TABLA_EVENTOS, "Add", filas_eventos)


# -----------------------------------------------------------------------------
# Endpoints
# -----------------------------------------------------------------------------


@app.route("/")
def home():
    return "API Documentos funcionando"


@app.route("/health")
def health():
    return {
        "status": "ok",
        "api": "documentos",
        "server_time": ahora_iso(),
    }


# -----------------------------------------------------------------------------
# Flujo: envío a revisión
# -----------------------------------------------------------------------------


def crear_registro_version_revision(
    id_version: str,
    id_documento: str,
    id_version_origen: str,
    numero_version: int,
    numero_revision: int,
    nombre_archivo: str,
    google_doc_id: str,
    google_doc_url: str,
    id_aprobacion_responsable: str,
    orden_responsable: int,
    creado_por: str,
    fecha_creacion: str,
    comentario: str,
) -> None:
    fila = {
        "ID_VERSION": id_version,
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION_ORIGEN": id_version_origen,
        "NUMERO_VERSION": numero_version,
        "NUMERO_REVISION": numero_revision,
        "ETAPA": "Revisión",
        "ESTADO_VERSION": "Activa",
        "NOMBRE_ARCHIVO": nombre_archivo,
        "GOOGLE_DOC_ID": google_doc_id,
        "GOOGLE_DOC_URL": google_doc_url,
        "ID_APROBACION_RESPONSABLE": id_aprobacion_responsable,
        "ORDEN_RESPONSABLE": orden_responsable,
        "MOTIVO_CREACION": "Envío a revisión",
        "COMENTARIO_CAMBIO": comentario,
        "CREADO_POR": creado_por,
        "FECHA_CREACION": fecha_creacion,
    }
    appsheet_action(TABLA_VERSIONES, "Add", [fila])


def cerrar_version(id_version: str, fecha_cierre: str) -> None:
    appsheet_action(
        TABLA_VERSIONES,
        "Edit",
        [
            {
                "ID_VERSION": id_version,
                "ESTADO_VERSION": "Cerrada",
                "FECHA_CIERRE": fecha_cierre,
            }
        ],
    )


def actualizar_aprobadores_envio_revision(
    actual: dict[str, Any],
    siguiente: dict[str, Any],
    id_version_nueva: str,
    permission_id_siguiente: str,
    comentario: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_APROBADORES_ACTUAL,
        "Edit",
        [
            {
                "ID_APROBACION_ACTUAL": actual["ID_APROBACION_ACTUAL"],
                "ESTADO": "Cerrado",
                "RESULTADO": "Enviado",
                "COMENTARIO": comentario,
                "FECHA_RESPUESTA": fecha,
            },
            {
                "ID_APROBACION_ACTUAL": siguiente[
                    "ID_APROBACION_ACTUAL"
                ],
                "ESTADO": "En revisión",
                "ID_VERSION_TRABAJADA": id_version_nueva,
                "FECHA_INICIO": fecha,
                "PERMISSION_ID_DRIVE": permission_id_siguiente,
            },
        ],
    )


def actualizar_documento_envio_revision(
    id_documento: str,
    numero_version: int,
    numero_revision: int,
    id_version: str,
    copia: dict[str, str],
    siguiente: dict[str, Any],
    usuario: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "ESTADO": "En revisión",
                "VERSION_ACTUAL": numero_version,
                "REVISION_ACTUAL": numero_revision,
                "ID_VERSION_ACTUAL": id_version,
                "GOOGLE_DOC_ID": copia["id"],
                "GOOGLE_DOC_URL": copia["url"],
                "ORDEN_ACTUAL": siguiente["ORDEN"],
                "ID_APROBACION_ACTUAL": siguiente[
                    "ID_APROBACION_ACTUAL"
                ],
                "ENCARGADO_ACTUAL_NOMBRE": siguiente.get("NOMBRE", ""),
                "ENCARGADO_ACTUAL_EMAIL": siguiente.get("APROBADOR", ""),
                "ULTIMO_ENVIADO_POR": usuario,
                "FECHA_ULTIMO_ENVIO": fecha,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "OBSERVACION_ACTUAL": "",
            }
        ],
    )


def crear_evento_envio_revision(
    id_documento: str,
    id_version: str,
    id_aprobacion_actual: str,
    usuario: str,
    fecha: str,
    comentario: str,
    nombre_archivo: str,
) -> dict[str, Any]:
    """
    Registra el evento que identifica de forma inequívoca el envío a revisión.

    La fila se devuelve para que las notificaciones utilicen exactamente el
    mismo ID_EVENTO como clave de idempotencia.
    """
    detalle = f"Se creó {nombre_archivo} y se entregó al siguiente responsable."
    if comentario:
        detalle += f" Comentario: {comentario}"

    evento = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION": id_version,
        "ID_APROBACION_ACTUAL": id_aprobacion_actual,
        "TIPO_EVENTO": "Enviado a revisión",
        "ESTADO_ANTERIOR": "Borrador",
        "ESTADO_NUEVO": "En revisión",
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": detalle,
    }

    appsheet_action(
        TABLA_EVENTOS,
        "Add",
        [evento],
    )
    return evento


@app.route("/crear-documento", methods=["POST"])
def crear_documento():
    id_documento = ""

    try:
        validar_configuracion()
        validar_token()

        data = request.get_json(silent=True) or {}
        id_documento = texto(data.get("id_documento"))

        if not id_documento:
            return {"error": "Falta id_documento"}, 400

        documento = buscar_documento(id_documento)

        # Evita duplicados por reintentos del Bot o doble ejecución.
        id_version_existente = texto(documento.get("ID_VERSION_ACTUAL"))
        google_doc_existente = texto(documento.get("GOOGLE_DOC_ID"))
        versiones_existentes = buscar_versiones_documento(id_documento)

        if id_version_existente or google_doc_existente or versiones_existentes:
            return jsonify(
                {
                    "ok": True,
                    "ya_existia": True,
                    "id_documento": id_documento,
                    "id_version": id_version_existente,
                    "google_doc_id": google_doc_existente,
                    "google_doc_url": texto(documento.get("GOOGLE_DOC_URL")),
                    "mensaje": "El documento ya tiene una versión creada.",
                }
            )

        id_plantilla = texto(documento.get("ID_PLANTILLA"))
        titulo = texto(documento.get("TITULO")) or f"Documento_{id_documento}"
        creado_por = (
            texto(documento.get("CREADO_POR"))
            or texto(data.get("creado_por"))
        )

        if not id_plantilla:
            raise ValueError("El documento no tiene ID_PLANTILLA")

        plantilla = buscar_plantilla(id_plantilla)
        cadena_plantilla = buscar_cadena_plantilla(id_plantilla)

        template_id = texto(plantilla.get("GOOGLE_DOC_TEMPLATE_ID"))
        folder_id = texto(plantilla.get("CARPETA_DESTINO_ID"))

        if not template_id:
            raise ValueError("La plantilla no tiene GOOGLE_DOC_TEMPLATE_ID")

        if not folder_id:
            raise ValueError("La plantilla no tiene CARPETA_DESTINO_ID")

        primer_base = cadena_plantilla[0]
        primer_email = texto(primer_base.get("APROBADOR"))
        primer_orden = entero(primer_base.get("ORDEN"), "ORDEN")

        fecha_creacion = ahora_iso()
        id_version = nuevo_id()
        nombre_archivo = limpiar_nombre_archivo(
            f"{titulo}_V01_BORRADOR"
        )

        drive_service = obtener_drive_service()
        copia = copiar_plantilla(
            drive_service=drive_service,
            template_id=template_id,
            folder_id=folder_id,
            nombre_documento=nombre_archivo,
        )

        permission_id_drive = asegurar_permiso_editor(
            drive_service=drive_service,
            file_id=copia["id"],
            email=primer_email,
        )

        filas_cadena_actual, primer_actual = construir_cadena_actual(
            cadena_plantilla=cadena_plantilla,
            id_documento=id_documento,
            id_plantilla=id_plantilla,
            id_version=id_version,
            permission_id_drive=permission_id_drive,
            fecha_inicio=fecha_creacion,
        )

        # 1. Copia la cadena de aprobación particular del documento.
        appsheet_action(
            TABLA_APROBADORES_ACTUAL,
            "Add",
            filas_cadena_actual,
        )

        # 2. Registra la versión/archivo inicial.
        crear_registro_version_inicial(
            id_version=id_version,
            id_documento=id_documento,
            id_aprobacion_responsable=primer_actual[
                "ID_APROBACION_ACTUAL"
            ],
            orden_responsable=primer_orden,
            nombre_archivo=copia["name"],
            google_doc_id=copia["id"],
            google_doc_url=copia["url"],
            creado_por=creado_por,
            fecha_creacion=fecha_creacion,
        )

        # Vincula al primer responsable con la versión que trabajará.
        # Se realiza después de crear Documento_Versiones para que la Ref ya exista.
        appsheet_action(
            TABLA_APROBADORES_ACTUAL,
            "Edit",
            [
                {
                    "ID_APROBACION_ACTUAL": primer_actual[
                        "ID_APROBACION_ACTUAL"
                    ],
                    "ID_VERSION_TRABAJADA": id_version,
                }
            ],
        )

        # 3. Actualiza la cabecera operativa del documento.
        actualizar_documento_inicial(
            id_documento=id_documento,
            id_version=id_version,
            google_doc_id=copia["id"],
            google_doc_url=copia["url"],
            primer_actual=primer_actual,
            fecha_actualizacion=fecha_creacion,
        )

        # 4. Registra la bitácora inicial.
        crear_eventos_iniciales(
            id_documento=id_documento,
            id_version=id_version,
            id_aprobacion_actual=primer_actual[
                "ID_APROBACION_ACTUAL"
            ],
            usuario=creado_por,
            fecha_evento=fecha_creacion,
            nombre_archivo=copia["name"],
            cantidad_responsables=len(filas_cadena_actual),
        )

        return jsonify(
            {
                "ok": True,
                "ya_existia": False,
                "id_documento": id_documento,
                "id_version": id_version,
                "numero_version": 1,
                "numero_revision": 0,
                "estado": "Borrador",
                "id_aprobacion_actual": primer_actual[
                    "ID_APROBACION_ACTUAL"
                ],
                "orden_actual": primer_actual["ORDEN"],
                "encargado_actual": primer_actual.get("NOMBRE", ""),
                "encargado_email": primer_actual.get("APROBADOR", ""),
                "google_doc_id": copia["id"],
                "google_doc_url": copia["url"],
                "nombre_archivo": copia["name"],
                "cantidad_responsables": len(filas_cadena_actual),
            }
        )

    except PermissionError as exc:
        return {"error": str(exc)}, 403

    except (ValueError, LookupError) as exc:
        if id_documento and not documento_cerrado:
            marcar_documento_error(id_documento, str(exc))
        return {"error": str(exc)}, 400

    except Exception as exc:
        traceback.print_exc()
        if id_documento and not documento_cerrado:
            marcar_documento_error(id_documento, str(exc))
        return {"error": str(exc)}, 500


@app.route("/enviar-revision", methods=["POST"])
def enviar_revision():
    id_documento = ""

    try:
        validar_configuracion()
        validar_token()

        data = request.get_json(silent=True) or {}
        id_documento = texto(data.get("id_documento"))
        id_aprobacion_solicitud = texto(
            data.get("id_aprobacion_actual")
        )
        usuario = texto(data.get("usuario"))
        comentario = texto(data.get("comentario"))

        if not id_documento:
            return {"error": "Falta id_documento"}, 400

        documento = buscar_documento(id_documento)
        estado_documento = texto(documento.get("ESTADO"))

        # Respuesta idempotente ante reintentos del Bot. Además, intenta
        # completar notificaciones pendientes o con error sin repetir las ya
        # enviadas, gracias a CLAVE_IDEMPOTENCIA.
        if estado_documento == "En revisión":
            notificaciones_reintento: list[dict[str, Any]] = []
            advertencias_reintento: list[str] = []

            try:
                (
                    notificaciones_reintento,
                    advertencias_reintento,
                ) = reanudar_notificaciones_envio_revision(
                    documento=documento,
                    datos_solicitud=data,
                )
            except Exception as exc_notificacion:
                traceback.print_exc()
                advertencias_reintento.append(
                    "El documento ya estaba en revisión, pero no fue posible "
                    "reanudar sus notificaciones: "
                    f"{exc_notificacion}"
                )

            return jsonify(
                {
                    "ok": True,
                    "ya_procesado": True,
                    "id_documento": id_documento,
                    "id_version": texto(documento.get("ID_VERSION_ACTUAL")),
                    "google_doc_id": texto(documento.get("GOOGLE_DOC_ID")),
                    "google_doc_url": texto(documento.get("GOOGLE_DOC_URL")),
                    "estado": estado_documento,
                    "mensaje": "El documento ya fue enviado a revisión.",
                    "notificaciones": notificaciones_reintento,
                    "advertencias": advertencias_reintento,
                }
            )

        if estado_documento != "Borrador":
            raise ValueError(
                "Solo se puede enviar a revisión un documento en estado "
                f"Borrador. Estado actual: {estado_documento!r}"
            )

        numero_version = entero(
            documento.get("VERSION_ACTUAL") or 1,
            "VERSION_ACTUAL",
        )
        revision_actual = entero(
            documento.get("REVISION_ACTUAL") or 0,
            "REVISION_ACTUAL",
        )
        numero_revision_nueva = revision_actual + 1

        id_version_actual = texto(documento.get("ID_VERSION_ACTUAL"))
        id_aprobacion_actual = texto(
            documento.get("ID_APROBACION_ACTUAL")
        )
        google_doc_id_actual = texto(documento.get("GOOGLE_DOC_ID"))

        if not id_version_actual:
            raise ValueError("Documentos no tiene ID_VERSION_ACTUAL")
        if not id_aprobacion_actual:
            raise ValueError("Documentos no tiene ID_APROBACION_ACTUAL")
        if not google_doc_id_actual:
            raise ValueError("Documentos no tiene GOOGLE_DOC_ID")

        if (
            id_aprobacion_solicitud
            and id_aprobacion_solicitud != id_aprobacion_actual
        ):
            raise ValueError(
                "El encargado enviado por AppSheet ya no coincide con el "
                "encargado actual del documento"
            )

        aprobacion_actual = buscar_aprobacion_actual(
            id_aprobacion_actual
        )

        if not es_verdadero(aprobacion_actual.get("CADENA_ACTIVA")):
            raise ValueError("La cadena de aprobación actual no está activa")

        estado_aprobador = texto(aprobacion_actual.get("ESTADO"))
        if estado_aprobador not in {"En elaboración", "Cerrado"}:
            raise ValueError(
                "El encargado actual no está en elaboración. "
                f"Estado encontrado: {estado_aprobador!r}"
            )

        email_actual = texto(aprobacion_actual.get("APROBADOR"))
        if usuario and email_actual.lower() != usuario.lower():
            raise PermissionError(
                "Solo el encargado actual puede enviar el documento "
                "a revisión"
            )
        usuario = usuario or email_actual

        cadena_actual = buscar_cadena_actual_documento(
            id_documento=id_documento,
            numero_version=numero_version,
        )
        orden_actual = entero(
            aprobacion_actual.get("ORDEN"),
            "ORDEN",
        )

        siguientes = [
            fila
            for fila in cadena_actual
            if entero(fila.get("ORDEN"), "ORDEN") > orden_actual
            and texto(fila.get("ESTADO")) in {"Pendiente", "En revisión"}
        ]

        if not siguientes:
            raise ValueError(
                "No existe un siguiente responsable en la cadena. "
                "Cuando sea el último encargado se debe usar el flujo "
                "Listo para firma."
            )

        siguiente = siguientes[0]
        siguiente_email = texto(siguiente.get("APROBADOR"))
        siguiente_orden = entero(siguiente.get("ORDEN"), "ORDEN")

        if not siguiente_email:
            raise ValueError(
                f"El responsable de orden {siguiente_orden} no tiene correo"
            )

        version_actual = buscar_version_por_id(id_version_actual)
        etapa_actual = texto(version_actual.get("ETAPA"))
        if etapa_actual != "Borrador":
            raise ValueError(
                "La versión vigente no corresponde a un borrador. "
                f"ETAPA encontrada: {etapa_actual!r}"
            )

        id_plantilla = texto(documento.get("ID_PLANTILLA"))
        plantilla = buscar_plantilla(id_plantilla)
        folder_id = texto(plantilla.get("CARPETA_DESTINO_ID"))
        if not folder_id:
            raise ValueError("La plantilla no tiene CARPETA_DESTINO_ID")

        titulo = texto(documento.get("TITULO")) or f"Documento_{id_documento}"
        nombre_archivo = limpiar_nombre_archivo(
            f"{titulo}_V{numero_version:02d}_REV{numero_revision_nueva:02d}"
        )
        fecha = ahora_iso()

        version_existente = buscar_version_numero_revision(
            id_documento=id_documento,
            numero_version=numero_version,
            numero_revision=numero_revision_nueva,
        )

        drive_service = obtener_drive_service()

        if version_existente:
            id_version_nueva = texto(version_existente.get("ID_VERSION"))
            copia = {
                "id": texto(version_existente.get("GOOGLE_DOC_ID")),
                "url": texto(version_existente.get("GOOGLE_DOC_URL")),
                "name": texto(version_existente.get("NOMBRE_ARCHIVO"))
                or nombre_archivo,
            }
            if not copia["id"]:
                raise RuntimeError(
                    "La versión de revisión existente no tiene GOOGLE_DOC_ID"
                )
        else:
            id_version_nueva = nuevo_id()
            copia = copiar_archivo_o_reutilizar(
                drive_service=drive_service,
                source_file_id=google_doc_id_actual,
                folder_id=folder_id,
                nombre_archivo=nombre_archivo,
            )

        # El archivo de borrador queda congelado para su elaborador.
        asegurar_permiso_rol(
            drive_service=drive_service,
            file_id=google_doc_id_actual,
            email=email_actual,
            role="reader",
        )

        # En la nueva revisión, el elaborador comenta y el siguiente revisa.
        asegurar_permiso_rol(
            drive_service=drive_service,
            file_id=copia["id"],
            email=email_actual,
            role="commenter",
        )
        permission_id_siguiente = asegurar_permiso_rol(
            drive_service=drive_service,
            file_id=copia["id"],
            email=siguiente_email,
            role="writer",
        )

        if not version_existente:
            crear_registro_version_revision(
                id_version=id_version_nueva,
                id_documento=id_documento,
                id_version_origen=id_version_actual,
                numero_version=numero_version,
                numero_revision=numero_revision_nueva,
                nombre_archivo=copia["name"],
                google_doc_id=copia["id"],
                google_doc_url=copia["url"],
                id_aprobacion_responsable=siguiente[
                    "ID_APROBACION_ACTUAL"
                ],
                orden_responsable=siguiente_orden,
                creado_por=usuario,
                fecha_creacion=fecha,
                comentario=comentario,
            )

        cerrar_version(id_version_actual, fecha)

        actualizar_aprobadores_envio_revision(
            actual=aprobacion_actual,
            siguiente=siguiente,
            id_version_nueva=id_version_nueva,
            permission_id_siguiente=permission_id_siguiente,
            comentario=comentario,
            fecha=fecha,
        )

        actualizar_documento_envio_revision(
            id_documento=id_documento,
            numero_version=numero_version,
            numero_revision=numero_revision_nueva,
            id_version=id_version_nueva,
            copia=copia,
            siguiente=siguiente,
            usuario=usuario,
            fecha=fecha,
        )

        advertencias: list[str] = []
        notificaciones: list[dict[str, Any]] = []
        evento_revision: dict[str, Any] | None = None

        try:
            evento_revision = crear_evento_envio_revision(
                id_documento=id_documento,
                id_version=id_version_nueva,
                id_aprobacion_actual=siguiente[
                    "ID_APROBACION_ACTUAL"
                ],
                usuario=usuario,
                fecha=fecha,
                comentario=comentario,
                nombre_archivo=copia["name"],
            )
        except Exception as exc_evento:
            traceback.print_exc()
            advertencias.append(
                "La transición terminó, pero no se pudo crear el evento: "
                f"{exc_evento}"
            )

        # La transición documental ya terminó. Desde este punto, cualquier
        # error de notificación se registra como advertencia y nunca revierte
        # el cambio de estado ni los permisos de Drive.
        if evento_revision is not None:
            try:
                documento_actualizado = buscar_documento(id_documento)
                notificaciones = ejecutar_notificaciones_envio_revision(
                    documento=documento_actualizado,
                    evento=evento_revision,
                    cadena=cadena_actual,
                    aprobador_anterior=aprobacion_actual,
                    aprobador_actual=siguiente,
                )

                fallidas = [
                    resultado
                    for resultado in notificaciones
                    if not resultado.get("ok")
                ]
                if fallidas:
                    advertencias.append(
                        "La transición terminó, pero "
                        f"{len(fallidas)} notificación(es) quedaron "
                        "omitidas o con error. Revisa "
                        "Documento_Notificaciones."
                    )
            except Exception as exc_notificacion:
                traceback.print_exc()
                advertencias.append(
                    "La transición terminó, pero no se pudieron procesar "
                    f"las notificaciones internas: {exc_notificacion}"
                )

        return jsonify(
            {
                "ok": True,
                "ya_procesado": False,
                "id_documento": id_documento,
                "estado": "En revisión",
                "numero_version": numero_version,
                "numero_revision": numero_revision_nueva,
                "id_version": id_version_nueva,
                "google_doc_id": copia["id"],
                "google_doc_url": copia["url"],
                "nombre_archivo": copia["name"],
                "orden_actual": siguiente_orden,
                "id_aprobacion_actual": siguiente[
                    "ID_APROBACION_ACTUAL"
                ],
                "encargado_actual": siguiente.get("NOMBRE", ""),
                "encargado_email": siguiente_email,
                "notificaciones": notificaciones,
                "advertencias": advertencias,
            }
        )

    except PermissionError as exc:
        if id_documento:
            registrar_error_transicion(id_documento, str(exc))
        return {"error": str(exc)}, 403

    except (ValueError, LookupError) as exc:
        if id_documento:
            registrar_error_transicion(id_documento, str(exc))
        return {"error": str(exc)}, 400

    except Exception as exc:
        traceback.print_exc()
        if id_documento:
            registrar_error_transicion(id_documento, str(exc))
        return {"error": str(exc)}, 500


# -----------------------------------------------------------------------------
# Flujo: aprobar revisión
# -----------------------------------------------------------------------------


@medir_operacion("drive.exportar_pdf_y_guardar")
def exportar_pdf_o_reutilizar(
    drive_service: Any,
    google_doc_id: str,
    folder_id: str,
    nombre_pdf: str,
) -> dict[str, str]:
    """
    Exporta un Google Docs a PDF y guarda el archivo en la carpeta destino.
    Si ya existe un PDF con el mismo nombre, lo reutiliza para tolerar
    reintentos del Bot.
    """
    existente = buscar_archivo_en_carpeta(
        drive_service=drive_service,
        folder_id=folder_id,
        nombre_archivo=nombre_pdf,
    )
    if existente:
        return existente

    contenido_pdf = (
        drive_service.files()
        .export(
            fileId=google_doc_id,
            mimeType="application/pdf",
        )
        .execute()
    )

    media = MediaInMemoryUpload(
        contenido_pdf,
        mimetype="application/pdf",
        resumable=False,
    )

    archivo = (
        drive_service.files()
        .create(
            body={
                "name": nombre_pdf,
                "parents": [folder_id],
                "mimeType": "application/pdf",
            },
            media_body=media,
            fields="id,name,webViewLink,webContentLink",
            supportsAllDrives=True,
        )
        .execute()
    )

    file_id = texto(archivo.get("id"))
    return {
        "id": file_id,
        "name": texto(archivo.get("name")) or nombre_pdf,
        "url": (
            texto(archivo.get("webViewLink"))
            or texto(archivo.get("webContentLink"))
            or f"https://drive.google.com/file/d/{file_id}/view"
        ),
    }


def actualizar_estado_version(
    id_version: str,
    estado_version: str,
    fecha_cierre: str,
) -> None:
    appsheet_action(
        TABLA_VERSIONES,
        "Edit",
        [
            {
                "ID_VERSION": id_version,
                "ESTADO_VERSION": estado_version,
                "FECHA_CIERRE": fecha_cierre,
            }
        ],
    )


def crear_registro_version_por_aprobacion(
    *,
    id_version: str,
    id_documento: str,
    id_version_origen: str,
    numero_version: int,
    numero_revision: int,
    etapa: str,
    nombre_archivo: str,
    google_doc_id: str,
    google_doc_url: str,
    id_aprobacion_responsable: str,
    orden_responsable: int,
    motivo_creacion: str,
    comentario: str,
    creado_por: str,
    fecha_creacion: str,
    pdf_version_id: str = "",
    pdf_version_url: str = "",
) -> None:
    fila = {
        "ID_VERSION": id_version,
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION_ORIGEN": id_version_origen,
        "NUMERO_VERSION": numero_version,
        "NUMERO_REVISION": numero_revision,
        "ETAPA": etapa,
        "ESTADO_VERSION": "Activa",
        "NOMBRE_ARCHIVO": nombre_archivo,
        "GOOGLE_DOC_ID": google_doc_id,
        "GOOGLE_DOC_URL": google_doc_url,
        "PDF_VERSION_ID": pdf_version_id,
        "PDF_VERSION_URL": pdf_version_url,
        "ID_APROBACION_RESPONSABLE": id_aprobacion_responsable,
        "ORDEN_RESPONSABLE": orden_responsable,
        "MOTIVO_CREACION": motivo_creacion,
        "COMENTARIO_CAMBIO": comentario,
        "CREADO_POR": creado_por,
        "FECHA_CREACION": fecha_creacion,
    }
    appsheet_action(TABLA_VERSIONES, "Add", [fila])


def actualizar_aprobadores_aprobacion_intermedia(
    actual: dict[str, Any],
    siguiente: dict[str, Any],
    id_version_nueva: str,
    permission_id_siguiente: str,
    comentario: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_APROBADORES_ACTUAL,
        "Edit",
        [
            {
                "ID_APROBACION_ACTUAL": actual[
                    "ID_APROBACION_ACTUAL"
                ],
                "ESTADO": "Cerrado",
                "RESULTADO": "Aprobado",
                "COMENTARIO": comentario,
                "FECHA_RESPUESTA": fecha,
            },
            {
                "ID_APROBACION_ACTUAL": siguiente[
                    "ID_APROBACION_ACTUAL"
                ],
                "ESTADO": "En revisión",
                "ID_VERSION_TRABAJADA": id_version_nueva,
                "FECHA_INICIO": fecha,
                "PERMISSION_ID_DRIVE": permission_id_siguiente,
            },
        ],
    )


def cerrar_cadena_para_firma(
    cadena_actual: list[dict[str, Any]],
    aprobacion_actual: dict[str, Any],
    comentario: str,
    fecha: str,
) -> None:
    filas: list[dict[str, Any]] = []
    id_actual = texto(
        aprobacion_actual.get("ID_APROBACION_ACTUAL")
    )

    for fila in cadena_actual:
        id_fila = texto(fila.get("ID_APROBACION_ACTUAL"))
        cambios: dict[str, Any] = {
            "ID_APROBACION_ACTUAL": id_fila,
            "CADENA_ACTIVA": False,
        }

        if id_fila == id_actual:
            cambios.update(
                {
                    "ESTADO": "Cerrado",
                    "RESULTADO": "Aprobado",
                    "COMENTARIO": comentario,
                    "FECHA_RESPUESTA": fecha,
                }
            )

        filas.append(cambios)

    appsheet_action(
        TABLA_APROBADORES_ACTUAL,
        "Edit",
        filas,
    )


def actualizar_documento_aprobacion_intermedia(
    *,
    id_documento: str,
    numero_version: int,
    numero_revision: int,
    id_version: str,
    copia: dict[str, str],
    siguiente: dict[str, Any],
    usuario: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "ESTADO": "En revisión",
                "VERSION_ACTUAL": numero_version,
                "REVISION_ACTUAL": numero_revision,
                "ID_VERSION_ACTUAL": id_version,
                "GOOGLE_DOC_ID": copia["id"],
                "GOOGLE_DOC_URL": copia["url"],
                "ORDEN_ACTUAL": siguiente["ORDEN"],
                "ID_APROBACION_ACTUAL": siguiente[
                    "ID_APROBACION_ACTUAL"
                ],
                "ENCARGADO_ACTUAL_NOMBRE": siguiente.get("NOMBRE", ""),
                "ENCARGADO_ACTUAL_EMAIL": siguiente.get("APROBADOR", ""),
                "ULTIMO_ENVIADO_POR": usuario,
                "FECHA_ULTIMO_ENVIO": fecha,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "OBSERVACION_ACTUAL": "",
                "ACCION_SOLICITADA": "",
            }
        ],
    )


def actualizar_documento_listo_para_firma(
    *,
    id_documento: str,
    numero_version: int,
    numero_revision: int,
    id_version: str,
    copia: dict[str, str],
    pdf: dict[str, str],
    usuario: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "ESTADO": "Listo para firma",
                "VERSION_ACTUAL": numero_version,
                "REVISION_ACTUAL": numero_revision,
                "ID_VERSION_ACTUAL": id_version,
                "GOOGLE_DOC_ID": copia["id"],
                "GOOGLE_DOC_URL": copia["url"],
                "ORDEN_ACTUAL": "",
                "ID_APROBACION_ACTUAL": "",
                "ENCARGADO_ACTUAL_NOMBRE": "",
                "ENCARGADO_ACTUAL_EMAIL": "",
                "PDF_PARA_FIRMA_ID": pdf["id"],
                "PDF_PARA_FIRMA_URL": pdf["url"],
                "ESTADO_FIRMA": "No iniciado",
                "ULTIMO_ENVIADO_POR": usuario,
                "FECHA_ULTIMO_ENVIO": fecha,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "OBSERVACION_ACTUAL": "",
                "ACCION_SOLICITADA": "",
            }
        ],
    )


def crear_evento_revision_aprobada(
    *,
    id_documento: str,
    id_version: str,
    id_aprobacion_actual: str,
    usuario: str,
    fecha: str,
    comentario: str,
    orden_actual: int,
    orden_siguiente: int | None,
) -> dict[str, Any]:
    if orden_siguiente is None:
        estado_nuevo = "Listo para firma"
        detalle = (
            f"El responsable de orden {orden_actual} aprobó la revisión final."
        )
    else:
        estado_nuevo = f"En revisión - Orden {orden_siguiente}"
        detalle = (
            f"El responsable de orden {orden_actual} aprobó la revisión "
            f"y el documento avanzó al orden {orden_siguiente}."
        )

    if comentario:
        detalle += f" Comentario: {comentario}"

    evento = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION": id_version,
        "ID_APROBACION_ACTUAL": id_aprobacion_actual,
        "TIPO_EVENTO": "Revisión aprobada",
        "ESTADO_ANTERIOR": f"En revisión - Orden {orden_actual}",
        "ESTADO_NUEVO": estado_nuevo,
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": detalle,
    }

    appsheet_action(TABLA_EVENTOS, "Add", [evento])
    return evento


def crear_evento_preparado_firma(
    *,
    id_documento: str,
    id_version: str,
    id_aprobacion_actual: str,
    usuario: str,
    fecha: str,
    nombre_archivo: str,
    nombre_pdf: str,
) -> None:
    appsheet_action(
        TABLA_EVENTOS,
        "Add",
        [
            {
                "ID_EVENTO": nuevo_id(),
                "ID_DOCUMENTO": id_documento,
                "ID_VERSION": id_version,
                "ID_APROBACION_ACTUAL": id_aprobacion_actual,
                "TIPO_EVENTO": "Preparado para firma",
                "ESTADO_ANTERIOR": "En revisión",
                "ESTADO_NUEVO": "Listo para firma",
                "USUARIO": usuario,
                "FECHA_EVENTO": fecha,
                "COMENTARIO": (
                    f"Se creó {nombre_archivo} y se exportó {nombre_pdf}."
                ),
            }
        ],
    )


@app.route("/aprobar-revision", methods=["POST"])
def aprobar_revision():
    id_documento = ""

    try:
        validar_configuracion()
        validar_token()

        data = request.get_json(silent=True) or {}
        id_documento = texto(data.get("id_documento"))
        id_aprobacion_solicitud = texto(
            data.get("id_aprobacion_actual")
        )
        usuario = texto(data.get("usuario"))
        comentario = texto(data.get("comentario"))

        if not id_documento:
            return {"error": "Falta id_documento"}, 400

        documento = buscar_documento(id_documento)
        estado_documento = texto(documento.get("ESTADO"))

        if estado_documento == "Listo para firma":
            notificaciones_reintento: list[dict[str, Any]] = []
            advertencias_reintento: list[str] = []
            try:
                (
                    notificaciones_reintento,
                    advertencias_reintento,
                ) = reanudar_notificaciones_aprobacion_revision(
                    documento=documento,
                    id_aprobacion_aprueba=id_aprobacion_solicitud,
                    datos_solicitud=data,
                )
            except Exception as exc_notificacion:
                traceback.print_exc()
                advertencias_reintento.append(
                    "El documento ya estaba Listo para firma, pero no se "
                    "pudieron reanudar sus notificaciones: "
                    f"{exc_notificacion}"
                )

            return jsonify(
                {
                    "ok": True,
                    "ya_procesado": True,
                    "id_documento": id_documento,
                    "estado": estado_documento,
                    "id_version": texto(
                        documento.get("ID_VERSION_ACTUAL")
                    ),
                    "google_doc_id": texto(
                        documento.get("GOOGLE_DOC_ID")
                    ),
                    "google_doc_url": texto(
                        documento.get("GOOGLE_DOC_URL")
                    ),
                    "pdf_id": texto(
                        documento.get("PDF_PARA_FIRMA_ID")
                    ),
                    "pdf_url": texto(
                        documento.get("PDF_PARA_FIRMA_URL")
                    ),
                    "notificaciones": notificaciones_reintento,
                    "advertencias": advertencias_reintento,
                }
            )

        # Si el Bot repite la petición después de una aprobación intermedia,
        # reconoce la fila cerrada y responde sin crear otra copia.
        if id_aprobacion_solicitud:
            aprobacion_solicitada = buscar_aprobacion_actual(
                id_aprobacion_solicitud
            )
            if (
                texto(aprobacion_solicitada.get("ESTADO")) == "Cerrado"
                and texto(aprobacion_solicitada.get("RESULTADO"))
                == "Aprobado"
            ):
                notificaciones_reintento: list[dict[str, Any]] = []
                advertencias_reintento: list[str] = []
                try:
                    (
                        notificaciones_reintento,
                        advertencias_reintento,
                    ) = reanudar_notificaciones_aprobacion_revision(
                        documento=documento,
                        id_aprobacion_aprueba=(
                            id_aprobacion_solicitud
                        ),
                        datos_solicitud=data,
                    )
                except Exception as exc_notificacion:
                    traceback.print_exc()
                    advertencias_reintento.append(
                        "La aprobación ya estaba procesada, pero no se "
                        "pudieron reanudar sus notificaciones: "
                        f"{exc_notificacion}"
                    )

                return jsonify(
                    {
                        "ok": True,
                        "ya_procesado": True,
                        "id_documento": id_documento,
                        "estado": estado_documento,
                        "id_version": texto(
                            documento.get("ID_VERSION_ACTUAL")
                        ),
                        "google_doc_id": texto(
                            documento.get("GOOGLE_DOC_ID")
                        ),
                        "google_doc_url": texto(
                            documento.get("GOOGLE_DOC_URL")
                        ),
                        "notificaciones": notificaciones_reintento,
                        "advertencias": advertencias_reintento,
                    }
                )

        if estado_documento != "En revisión":
            raise ValueError(
                "Solo se puede aprobar un documento en estado En revisión. "
                f"Estado actual: {estado_documento!r}"
            )

        numero_version = entero(
            documento.get("VERSION_ACTUAL"),
            "VERSION_ACTUAL",
        )
        revision_actual = entero(
            documento.get("REVISION_ACTUAL"),
            "REVISION_ACTUAL",
        )
        numero_revision_nueva = revision_actual + 1

        id_version_actual = texto(documento.get("ID_VERSION_ACTUAL"))
        id_aprobacion_actual = texto(
            documento.get("ID_APROBACION_ACTUAL")
        )
        google_doc_id_actual = texto(documento.get("GOOGLE_DOC_ID"))

        if not id_version_actual:
            raise ValueError("Documentos no tiene ID_VERSION_ACTUAL")
        if not id_aprobacion_actual:
            raise ValueError("Documentos no tiene ID_APROBACION_ACTUAL")
        if not google_doc_id_actual:
            raise ValueError("Documentos no tiene GOOGLE_DOC_ID")

        if (
            id_aprobacion_solicitud
            and id_aprobacion_solicitud != id_aprobacion_actual
        ):
            raise ValueError(
                "El aprobador enviado por AppSheet ya no coincide con el "
                "encargado actual del documento"
            )

        aprobacion_actual = buscar_aprobacion_actual(
            id_aprobacion_actual
        )

        if not es_verdadero(aprobacion_actual.get("CADENA_ACTIVA")):
            raise ValueError("La cadena de aprobación actual no está activa")

        if texto(aprobacion_actual.get("ESTADO")) != "En revisión":
            raise ValueError(
                "El encargado actual no se encuentra En revisión. "
                f"Estado encontrado: "
                f"{texto(aprobacion_actual.get('ESTADO'))!r}"
            )

        email_actual = texto(aprobacion_actual.get("APROBADOR"))
        if usuario and email_actual.lower() != usuario.lower():
            raise PermissionError(
                "Solo el encargado actual puede aprobar la revisión"
            )
        usuario = usuario or email_actual

        version_actual = buscar_version_por_id(id_version_actual)
        if texto(version_actual.get("ETAPA")) != "Revisión":
            raise ValueError(
                "La versión vigente no corresponde a una revisión"
            )

        cadena_actual = buscar_cadena_actual_documento(
            id_documento=id_documento,
            numero_version=numero_version,
        )
        orden_actual = entero(
            aprobacion_actual.get("ORDEN"),
            "ORDEN",
        )

        siguientes = [
            fila
            for fila in cadena_actual
            if entero(fila.get("ORDEN"), "ORDEN") > orden_actual
            and texto(fila.get("ESTADO")) == "Pendiente"
        ]
        siguiente = siguientes[0] if siguientes else None

        id_plantilla = texto(documento.get("ID_PLANTILLA"))
        plantilla = buscar_plantilla(id_plantilla)
        folder_id = texto(plantilla.get("CARPETA_DESTINO_ID"))
        if not folder_id:
            raise ValueError("La plantilla no tiene CARPETA_DESTINO_ID")

        titulo = texto(documento.get("TITULO")) or f"Documento_{id_documento}"
        fecha = ahora_iso()
        drive_service = obtener_drive_service()

        if siguiente is not None:
            siguiente_orden = entero(siguiente.get("ORDEN"), "ORDEN")
            siguiente_email = texto(siguiente.get("APROBADOR"))
            if not siguiente_email:
                raise ValueError(
                    f"El responsable de orden {siguiente_orden} no tiene correo"
                )

            nombre_archivo = limpiar_nombre_archivo(
                f"{titulo}_V{numero_version:02d}"
                f"_REV{numero_revision_nueva:02d}"
            )

            version_existente = buscar_version_numero_revision(
                id_documento=id_documento,
                numero_version=numero_version,
                numero_revision=numero_revision_nueva,
            )

            if version_existente:
                id_version_nueva = texto(
                    version_existente.get("ID_VERSION")
                )
                copia = {
                    "id": texto(
                        version_existente.get("GOOGLE_DOC_ID")
                    ),
                    "url": texto(
                        version_existente.get("GOOGLE_DOC_URL")
                    ),
                    "name": (
                        texto(version_existente.get("NOMBRE_ARCHIVO"))
                        or nombre_archivo
                    ),
                }
            else:
                id_version_nueva = nuevo_id()
                copia = copiar_archivo_o_reutilizar(
                    drive_service=drive_service,
                    source_file_id=google_doc_id_actual,
                    folder_id=folder_id,
                    nombre_archivo=nombre_archivo,
                )

            asegurar_permiso_rol(
                drive_service=drive_service,
                file_id=google_doc_id_actual,
                email=email_actual,
                role="reader",
            )
            asegurar_permiso_rol(
                drive_service=drive_service,
                file_id=copia["id"],
                email=email_actual,
                role="commenter",
            )
            permission_id_siguiente = asegurar_permiso_rol(
                drive_service=drive_service,
                file_id=copia["id"],
                email=siguiente_email,
                role="writer",
            )

            if not version_existente:
                crear_registro_version_por_aprobacion(
                    id_version=id_version_nueva,
                    id_documento=id_documento,
                    id_version_origen=id_version_actual,
                    numero_version=numero_version,
                    numero_revision=numero_revision_nueva,
                    etapa="Revisión",
                    nombre_archivo=copia["name"],
                    google_doc_id=copia["id"],
                    google_doc_url=copia["url"],
                    id_aprobacion_responsable=siguiente[
                        "ID_APROBACION_ACTUAL"
                    ],
                    orden_responsable=siguiente_orden,
                    motivo_creacion="Aprobación de etapa",
                    comentario=comentario,
                    creado_por=usuario,
                    fecha_creacion=fecha,
                )

            actualizar_estado_version(
                id_version=id_version_actual,
                estado_version="Cerrada",
                fecha_cierre=fecha,
            )

            actualizar_aprobadores_aprobacion_intermedia(
                actual=aprobacion_actual,
                siguiente=siguiente,
                id_version_nueva=id_version_nueva,
                permission_id_siguiente=permission_id_siguiente,
                comentario=comentario,
                fecha=fecha,
            )

            actualizar_documento_aprobacion_intermedia(
                id_documento=id_documento,
                numero_version=numero_version,
                numero_revision=numero_revision_nueva,
                id_version=id_version_nueva,
                copia=copia,
                siguiente=siguiente,
                usuario=usuario,
                fecha=fecha,
            )

            advertencias: list[str] = []
            notificaciones: list[dict[str, Any]] = []
            evento_aprobacion: dict[str, Any] | None = None

            try:
                evento_aprobacion = crear_evento_revision_aprobada(
                    id_documento=id_documento,
                    id_version=id_version_nueva,
                    id_aprobacion_actual=siguiente[
                        "ID_APROBACION_ACTUAL"
                    ],
                    usuario=usuario,
                    fecha=fecha,
                    comentario=comentario,
                    orden_actual=orden_actual,
                    orden_siguiente=siguiente_orden,
                )
            except Exception as exc_evento:
                traceback.print_exc()
                advertencias.append(
                    "La aprobación terminó, pero no se pudo crear el "
                    f"evento: {exc_evento}"
                )

            if evento_aprobacion is not None:
                try:
                    documento_actualizado = buscar_documento(id_documento)
                    notificaciones = (
                        ejecutar_notificaciones_aprobacion_revision(
                            documento=documento_actualizado,
                            evento=evento_aprobacion,
                            cadena=cadena_actual,
                            aprobador_aprueba=aprobacion_actual,
                            aprobador_actual=siguiente,
                            ultimo_aprobador=False,
                        )
                    )
                    fallidas = [
                        resultado
                        for resultado in notificaciones
                        if not resultado.get("ok")
                    ]
                    if fallidas:
                        advertencias.append(
                            "La aprobación terminó, pero "
                            f"{len(fallidas)} notificación(es) quedaron "
                            "omitidas o con error. Revisa "
                            "Documento_Notificaciones."
                        )
                except Exception as exc_notificacion:
                    traceback.print_exc()
                    advertencias.append(
                        "La aprobación terminó, pero falló el proceso de "
                        f"notificaciones: {exc_notificacion}"
                    )

            return jsonify(
                {
                    "ok": True,
                    "ya_procesado": False,
                    "ultimo_aprobador": False,
                    "id_documento": id_documento,
                    "estado": "En revisión",
                    "numero_version": numero_version,
                    "numero_revision": numero_revision_nueva,
                    "id_version": id_version_nueva,
                    "google_doc_id": copia["id"],
                    "google_doc_url": copia["url"],
                    "nombre_archivo": copia["name"],
                    "orden_actual": siguiente_orden,
                    "id_aprobacion_actual": siguiente[
                        "ID_APROBACION_ACTUAL"
                    ],
                    "encargado_actual": siguiente.get("NOMBRE", ""),
                    "encargado_email": siguiente_email,
                    "notificaciones": notificaciones,
                    "advertencias": advertencias,
                }
            )

        # Último aprobador: prepara el Google Docs y el PDF para firma.
        nombre_archivo = limpiar_nombre_archivo(
            f"{titulo}_V{numero_version:02d}_PARA_FIRMA"
        )
        nombre_pdf = limpiar_nombre_archivo(
            f"{titulo}_V{numero_version:02d}_PARA_FIRMA.pdf"
        )

        version_existente = buscar_version_numero_revision(
            id_documento=id_documento,
            numero_version=numero_version,
            numero_revision=numero_revision_nueva,
        )

        if version_existente:
            id_version_nueva = texto(version_existente.get("ID_VERSION"))
            copia = {
                "id": texto(version_existente.get("GOOGLE_DOC_ID")),
                "url": texto(version_existente.get("GOOGLE_DOC_URL")),
                "name": (
                    texto(version_existente.get("NOMBRE_ARCHIVO"))
                    or nombre_archivo
                ),
            }
            pdf = {
                "id": texto(version_existente.get("PDF_VERSION_ID")),
                "url": texto(version_existente.get("PDF_VERSION_URL")),
                "name": nombre_pdf,
            }
        else:
            id_version_nueva = nuevo_id()
            copia = copiar_archivo_o_reutilizar(
                drive_service=drive_service,
                source_file_id=google_doc_id_actual,
                folder_id=folder_id,
                nombre_archivo=nombre_archivo,
            )
            pdf = exportar_pdf_o_reutilizar(
                drive_service=drive_service,
                google_doc_id=copia["id"],
                folder_id=folder_id,
                nombre_pdf=nombre_pdf,
            )

        # Todos los participantes conservan lectura sobre el archivo final.
        emails: set[str] = set()
        for fila in cadena_actual:
            email = texto(fila.get("APROBADOR")).lower()
            if email and "@" in email:
                emails.add(email)

        for email in emails:
            asegurar_permiso_rol(
                drive_service=drive_service,
                file_id=copia["id"],
                email=email,
                role="reader",
            )

        asegurar_permiso_rol(
            drive_service=drive_service,
            file_id=google_doc_id_actual,
            email=email_actual,
            role="reader",
        )

        if not version_existente:
            crear_registro_version_por_aprobacion(
                id_version=id_version_nueva,
                id_documento=id_documento,
                id_version_origen=id_version_actual,
                numero_version=numero_version,
                numero_revision=numero_revision_nueva,
                etapa="Para firma",
                nombre_archivo=copia["name"],
                google_doc_id=copia["id"],
                google_doc_url=copia["url"],
                pdf_version_id=pdf["id"],
                pdf_version_url=pdf["url"],
                id_aprobacion_responsable=id_aprobacion_actual,
                orden_responsable=orden_actual,
                motivo_creacion="Preparación para firma",
                comentario=comentario,
                creado_por=usuario,
                fecha_creacion=fecha,
            )

        actualizar_estado_version(
            id_version=id_version_actual,
            estado_version="Aprobada",
            fecha_cierre=fecha,
        )

        cerrar_cadena_para_firma(
            cadena_actual=cadena_actual,
            aprobacion_actual=aprobacion_actual,
            comentario=comentario,
            fecha=fecha,
        )

        actualizar_documento_listo_para_firma(
            id_documento=id_documento,
            numero_version=numero_version,
            numero_revision=numero_revision_nueva,
            id_version=id_version_nueva,
            copia=copia,
            pdf=pdf,
            usuario=usuario,
            fecha=fecha,
        )

        advertencias: list[str] = []
        notificaciones: list[dict[str, Any]] = []
        evento_aprobacion: dict[str, Any] | None = None

        try:
            evento_aprobacion = crear_evento_revision_aprobada(
                id_documento=id_documento,
                id_version=id_version_nueva,
                id_aprobacion_actual=id_aprobacion_actual,
                usuario=usuario,
                fecha=fecha,
                comentario=comentario,
                orden_actual=orden_actual,
                orden_siguiente=None,
            )
        except Exception as exc_evento:
            traceback.print_exc()
            advertencias.append(
                "La aprobación final terminó, pero no se pudo crear el "
                f"evento: {exc_evento}"
            )

        try:
            crear_evento_preparado_firma(
                id_documento=id_documento,
                id_version=id_version_nueva,
                id_aprobacion_actual=id_aprobacion_actual,
                usuario=usuario,
                fecha=fecha,
                nombre_archivo=copia["name"],
                nombre_pdf=pdf["name"],
            )
        except Exception as exc_evento_firma:
            traceback.print_exc()
            advertencias.append(
                "El documento quedó Listo para firma, pero no se pudo crear "
                "el evento técnico Preparado para firma: "
                f"{exc_evento_firma}"
            )

        if evento_aprobacion is not None:
            try:
                documento_actualizado = buscar_documento(id_documento)
                notificaciones = (
                    ejecutar_notificaciones_aprobacion_revision(
                        documento=documento_actualizado,
                        evento=evento_aprobacion,
                        cadena=cadena_actual,
                        aprobador_aprueba=aprobacion_actual,
                        aprobador_actual=None,
                        ultimo_aprobador=True,
                    )
                )
                fallidas = [
                    resultado
                    for resultado in notificaciones
                    if not resultado.get("ok")
                ]
                if fallidas:
                    advertencias.append(
                        "La aprobación final terminó, pero "
                        f"{len(fallidas)} notificación(es) quedaron "
                        "omitidas o con error. Revisa "
                        "Documento_Notificaciones."
                    )
            except Exception as exc_notificacion:
                traceback.print_exc()
                advertencias.append(
                    "La aprobación final terminó, pero falló el proceso de "
                    f"notificaciones: {exc_notificacion}"
                )

        return jsonify(
            {
                "ok": True,
                "ya_procesado": False,
                "ultimo_aprobador": True,
                "id_documento": id_documento,
                "estado": "Listo para firma",
                "numero_version": numero_version,
                "numero_revision": numero_revision_nueva,
                "id_version": id_version_nueva,
                "google_doc_id": copia["id"],
                "google_doc_url": copia["url"],
                "nombre_archivo": copia["name"],
                "pdf_id": pdf["id"],
                "pdf_url": pdf["url"],
                "pdf_nombre": pdf["name"],
                "notificaciones": notificaciones,
                "advertencias": advertencias,
            }
        )

    except PermissionError as exc:
        if id_documento:
            registrar_error_transicion(id_documento, str(exc))
        return {"error": str(exc)}, 403

    except (ValueError, LookupError) as exc:
        if id_documento:
            registrar_error_transicion(id_documento, str(exc))
        return {"error": str(exc)}, 400

    except Exception as exc:
        traceback.print_exc()
        if id_documento:
            registrar_error_transicion(id_documento, str(exc))
        return {"error": str(exc)}, 500


# -----------------------------------------------------------------------------
# Flujo: rechazar revisión
# -----------------------------------------------------------------------------


def buscar_cadena_documento_version(
    id_documento: str,
    numero_version: int,
) -> list[dict[str, Any]]:
    """Devuelve la cadena de una versión, esté activa o histórica."""
    selector = (
        f"FILTER({TABLA_APROBADORES_ACTUAL}, "
        f"[ID_DOCUMENTO] = {literal_appsheet(id_documento)})"
    )
    filas = appsheet_find(TABLA_APROBADORES_ACTUAL, selector)
    coincidentes: list[dict[str, Any]] = []

    for fila in filas:
        try:
            version_fila = entero(
                fila.get("NUMERO_VERSION"),
                "NUMERO_VERSION",
            )
        except ValueError:
            continue

        if version_fila == numero_version:
            coincidentes.append(fila)

    return sorted(
        coincidentes,
        key=lambda fila: entero(fila.get("ORDEN"), "ORDEN"),
    )


def construir_cadena_nueva_por_rechazo(
    *,
    cadena_anterior: list[dict[str, Any]],
    id_documento: str,
    id_plantilla: str,
    numero_version_anterior: int,
    numero_version_nueva: int,
    indice_destino: int,
    id_version_nueva: str,
    permission_id_destino: str,
    fecha: str,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """
    Crea una nueva instancia de la cadena sin alterar el historial anterior.

    - Responsables anteriores al destino: aprobación heredada.
    - Responsable destino: vuelve a elaboración o revisión.
    - Responsables posteriores: pendientes.
    """
    filas_nuevas: list[dict[str, Any]] = []
    destino_nuevo: dict[str, Any] | None = None

    for indice, fila_anterior in enumerate(cadena_anterior):
        orden = entero(fila_anterior.get("ORDEN"), "ORDEN")
        rol_flujo = texto(fila_anterior.get("ROL_FLUJO"))
        fila_nueva: dict[str, Any] = {
            "ID_APROBACION_ACTUAL": nuevo_id(),
            "ID_APROBACION_PLANTILLA": texto(
                fila_anterior.get("ID_APROBACION_PLANTILLA")
            ),
            "ID_PLANTILLA": id_plantilla,
            "ID_DOCUMENTO": id_documento,
            "NUMERO_VERSION": numero_version_nueva,
            "ORDEN": orden,
            "APROBADOR": texto(fila_anterior.get("APROBADOR")),
            "NOMBRE": texto(fila_anterior.get("NOMBRE")),
            "ROL_FLUJO": rol_flujo,
            "CADENA_ACTIVA": True,
        }

        if indice < indice_destino:
            fila_nueva.update(
                {
                    "ESTADO": "Heredado",
                    "RESULTADO": "Aprobación heredada",
                    "COMENTARIO": (
                        "Aprobación heredada desde la versión "
                        f"{numero_version_anterior}."
                    ),
                    "FECHA_RESPUESTA": fecha,
                }
            )
        elif indice == indice_destino:
            estado_destino = (
                "En elaboración"
                if indice_destino == 0
                or rol_flujo.lower() == "elaborador"
                else "En revisión"
            )
            fila_nueva.update(
                {
                    "ID_VERSION_TRABAJADA": id_version_nueva,
                    "ESTADO": estado_destino,
                    "FECHA_INICIO": fecha,
                    "PERMISSION_ID_DRIVE": permission_id_destino,
                }
            )
            destino_nuevo = fila_nueva
        else:
            fila_nueva["ESTADO"] = "Pendiente"

        filas_nuevas.append(fila_nueva)

    if destino_nuevo is None:
        raise RuntimeError(
            "No fue posible construir el responsable de retorno"
        )

    return filas_nuevas, destino_nuevo


def cerrar_cadena_anterior_por_rechazo(
    *,
    cadena_anterior: list[dict[str, Any]],
    id_aprobacion_rechaza: str,
    comentario: str,
    fecha: str,
) -> None:
    filas: list[dict[str, Any]] = []

    for fila in cadena_anterior:
        id_fila = texto(fila.get("ID_APROBACION_ACTUAL"))
        cambios: dict[str, Any] = {
            "ID_APROBACION_ACTUAL": id_fila,
            "CADENA_ACTIVA": False,
        }

        if id_fila == id_aprobacion_rechaza:
            cambios.update(
                {
                    "ESTADO": "Cerrado",
                    "RESULTADO": "Rechazado",
                    "COMENTARIO": comentario,
                    "FECHA_RESPUESTA": fecha,
                }
            )

        filas.append(cambios)

    appsheet_action(
        TABLA_APROBADORES_ACTUAL,
        "Edit",
        filas,
    )


def actualizar_destino_cadena_reutilizada(
    *,
    destino: dict[str, Any],
    id_version_nueva: str,
    permission_id_destino: str,
    estado_destino: str,
    fecha: str,
) -> None:
    """Completa una cadena ya creada por una ejecución parcial anterior."""
    appsheet_action(
        TABLA_APROBADORES_ACTUAL,
        "Edit",
        [
            {
                "ID_APROBACION_ACTUAL": destino[
                    "ID_APROBACION_ACTUAL"
                ],
                "ID_VERSION_TRABAJADA": id_version_nueva,
                "ESTADO": estado_destino,
                "RESULTADO": "",
                "COMENTARIO": "",
                "FECHA_INICIO": fecha,
                "FECHA_RESPUESTA": "",
                "PERMISSION_ID_DRIVE": permission_id_destino,
                "CADENA_ACTIVA": True,
            }
        ],
    )


def actualizar_documento_rechazo_revision(
    *,
    id_documento: str,
    numero_version: int,
    numero_revision: int,
    id_version: str,
    copia: dict[str, str],
    destino: dict[str, Any],
    estado_documento: str,
    usuario: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "ESTADO": estado_documento,
                "VERSION_ACTUAL": numero_version,
                "REVISION_ACTUAL": numero_revision,
                "ID_VERSION_ACTUAL": id_version,
                "GOOGLE_DOC_ID": copia["id"],
                "GOOGLE_DOC_URL": copia["url"],
                "ORDEN_ACTUAL": destino["ORDEN"],
                "ID_APROBACION_ACTUAL": destino[
                    "ID_APROBACION_ACTUAL"
                ],
                "ENCARGADO_ACTUAL_NOMBRE": destino.get("NOMBRE", ""),
                "ENCARGADO_ACTUAL_EMAIL": destino.get("APROBADOR", ""),
                "ESTADO_FIRMA": "No iniciado",
                "ULTIMO_ENVIADO_POR": usuario,
                "FECHA_ULTIMO_ENVIO": fecha,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "OBSERVACION_ACTUAL": "",
                "ACCION_SOLICITADA": "",
            }
        ],
    )


def crear_eventos_rechazo_revision(
    *,
    id_documento: str,
    id_version_rechazada: str,
    id_version_nueva: str,
    id_aprobacion_rechaza: str,
    id_aprobacion_destino: str,
    usuario: str,
    fecha: str,
    comentario: str,
    orden_rechaza: int,
    orden_destino: int,
    numero_version_nueva: int,
    nombre_archivo: str,
    estado_nuevo: str,
) -> tuple[list[str], dict[str, Any] | None]:
    """Crea la bitácora del rechazo y devuelve su evento principal."""
    advertencias: list[str] = []

    evento_rechazo: dict[str, Any] = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION": id_version_rechazada,
        "ID_APROBACION_ACTUAL": id_aprobacion_rechaza,
        "TIPO_EVENTO": "Revisión rechazada",
        "ESTADO_ANTERIOR": f"En revisión - Orden {orden_rechaza}",
        "ESTADO_NUEVO": f"{estado_nuevo} - Orden {orden_destino}",
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": comentario,
    }
    evento_version: dict[str, Any] = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION": id_version_nueva,
        "ID_APROBACION_ACTUAL": id_aprobacion_destino,
        "TIPO_EVENTO": "Nueva versión creada",
        "ESTADO_ANTERIOR": "Revisión rechazada",
        "ESTADO_NUEVO": estado_nuevo,
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": (
            f"Se creó la versión {numero_version_nueva}: "
            f"{nombre_archivo}. El flujo retrocedió desde el orden "
            f"{orden_rechaza} al orden {orden_destino}."
        ),
    }

    try:
        appsheet_action(
            TABLA_EVENTOS,
            "Add",
            [evento_rechazo, evento_version],
        )
        return advertencias, evento_rechazo
    except Exception as exc:
        traceback.print_exc()
        advertencias.append(
            "La transición terminó, pero no se pudieron crear los eventos: "
            f"{exc}"
        )
        return advertencias, None


@app.route("/rechazar-revision", methods=["POST"])
def rechazar_revision():
    id_documento = ""

    try:
        validar_configuracion()
        validar_token()

        data = request.get_json(silent=True) or {}
        id_documento = texto(data.get("id_documento"))
        id_aprobacion_solicitud = texto(
            data.get("id_aprobacion_actual")
        )
        usuario = texto(data.get("usuario"))
        comentario = texto(data.get("comentario"))

        if not id_documento:
            return {"error": "Falta id_documento"}, 400
        if not comentario:
            return {
                "error": "El comentario es obligatorio para rechazar una revisión"
            }, 400

        documento = buscar_documento(id_documento)

        # El ID enviado por AppSheet identifica de forma inequívoca la etapa
        # que tomó la decisión, incluso si una ejecución anterior quedó parcial.
        if not id_aprobacion_solicitud:
            id_aprobacion_solicitud = texto(
                documento.get("ID_APROBACION_ACTUAL")
            )
        if not id_aprobacion_solicitud:
            raise ValueError("Falta ID_APROBACION_ACTUAL")

        aprobacion_actual = buscar_aprobacion_actual(
            id_aprobacion_solicitud
        )
        numero_version_anterior = entero(
            aprobacion_actual.get("NUMERO_VERSION"),
            "NUMERO_VERSION",
        )
        estado_aprobacion = texto(aprobacion_actual.get("ESTADO"))
        resultado_aprobacion = texto(
            aprobacion_actual.get("RESULTADO")
        )
        ya_rechazada = (
            estado_aprobacion == "Cerrado"
            and resultado_aprobacion == "Rechazado"
        )

        version_documento = entero(
            documento.get("VERSION_ACTUAL"),
            "VERSION_ACTUAL",
        )

        # Reintento posterior a una transición ya completada. También intenta
        # completar notificaciones pendientes o con error.
        if ya_rechazada and version_documento > numero_version_anterior:
            notificaciones_reintento: list[dict[str, Any]] = []
            advertencias_reintento: list[str] = []
            try:
                documento = buscar_documento(id_documento)
                (
                    notificaciones_reintento,
                    advertencias_reintento,
                ) = reanudar_notificaciones_rechazo_revision(
                    documento=documento,
                    id_aprobacion_rechaza=id_aprobacion_solicitud,
                    datos_solicitud=data,
                )
            except Exception as exc_notificacion:
                traceback.print_exc()
                advertencias_reintento.append(
                    "El rechazo ya estaba procesado, pero no se pudieron "
                    "reanudar sus notificaciones: "
                    f"{exc_notificacion}"
                )

            return jsonify(
                {
                    "ok": True,
                    "ya_procesado": True,
                    "id_documento": id_documento,
                    "estado": texto(documento.get("ESTADO")),
                    "numero_version": version_documento,
                    "numero_revision": entero(
                        documento.get("REVISION_ACTUAL") or 0,
                        "REVISION_ACTUAL",
                    ),
                    "id_version": texto(
                        documento.get("ID_VERSION_ACTUAL")
                    ),
                    "google_doc_id": texto(
                        documento.get("GOOGLE_DOC_ID")
                    ),
                    "google_doc_url": normalizar_url_appsheet(
                        documento.get("GOOGLE_DOC_URL")
                    ),
                    "notificaciones": notificaciones_reintento,
                    "advertencias": advertencias_reintento,
                }
            )

        estado_documento_actual = texto(documento.get("ESTADO"))
        if estado_documento_actual != "En revisión":
            raise ValueError(
                "Solo se puede rechazar un documento en estado En revisión. "
                f"Estado actual: {estado_documento_actual!r}"
            )

        id_aprobacion_documento = texto(
            documento.get("ID_APROBACION_ACTUAL")
        )
        if (
            not ya_rechazada
            and id_aprobacion_documento != id_aprobacion_solicitud
        ):
            raise ValueError(
                "El aprobador enviado por AppSheet ya no coincide con el "
                "encargado actual del documento"
            )

        if not ya_rechazada:
            if not es_verdadero(aprobacion_actual.get("CADENA_ACTIVA")):
                raise ValueError(
                    "La cadena de aprobación actual no está activa"
                )
            if estado_aprobacion != "En revisión":
                raise ValueError(
                    "El encargado actual no se encuentra En revisión. "
                    f"Estado encontrado: {estado_aprobacion!r}"
                )

        email_rechaza = texto(aprobacion_actual.get("APROBADOR"))
        if usuario and email_rechaza.lower() != usuario.lower():
            raise PermissionError(
                "Solo el encargado actual puede rechazar la revisión"
            )
        usuario = usuario or email_rechaza

        id_version_rechazada = texto(
            aprobacion_actual.get("ID_VERSION_TRABAJADA")
        ) or texto(documento.get("ID_VERSION_ACTUAL"))
        if not id_version_rechazada:
            raise ValueError(
                "No se pudo identificar la versión que fue rechazada"
            )

        version_rechazada = buscar_version_por_id(
            id_version_rechazada
        )
        if texto(version_rechazada.get("ETAPA")) != "Revisión":
            raise ValueError(
                "La versión rechazada no corresponde a una revisión"
            )

        google_doc_id_rechazado = texto(
            version_rechazada.get("GOOGLE_DOC_ID")
        ) or texto(documento.get("GOOGLE_DOC_ID"))
        if not google_doc_id_rechazado:
            raise ValueError(
                "La versión rechazada no tiene GOOGLE_DOC_ID"
            )

        cadena_anterior = buscar_cadena_documento_version(
            id_documento=id_documento,
            numero_version=numero_version_anterior,
        )
        if not cadena_anterior:
            raise ValueError(
                "No se encontró la cadena asociada a la versión rechazada"
            )

        ids_cadena = [
            texto(fila.get("ID_APROBACION_ACTUAL"))
            for fila in cadena_anterior
        ]
        if id_aprobacion_solicitud not in ids_cadena:
            raise ValueError(
                "El responsable que rechaza no pertenece a la cadena indicada"
            )

        indice_rechaza = ids_cadena.index(id_aprobacion_solicitud)
        if indice_rechaza == 0:
            raise ValueError(
                "El primer responsable de la cadena no puede rechazar hacia atrás"
            )

        indice_destino = indice_rechaza - 1
        destino_anterior = cadena_anterior[indice_destino]
        orden_rechaza = entero(
            aprobacion_actual.get("ORDEN"),
            "ORDEN",
        )
        orden_destino = entero(
            destino_anterior.get("ORDEN"),
            "ORDEN",
        )
        email_destino = texto(destino_anterior.get("APROBADOR"))
        rol_destino = texto(destino_anterior.get("ROL_FLUJO"))
        if not email_destino:
            raise ValueError(
                f"El responsable de orden {orden_destino} no tiene correo"
            )

        numero_version_nueva = numero_version_anterior + 1
        es_borrador = (
            indice_destino == 0
            or rol_destino.lower() == "elaborador"
        )
        numero_revision_nueva = 0 if es_borrador else indice_destino
        etapa_nueva = "Borrador" if es_borrador else "Revisión"
        estado_documento_nuevo = (
            "Borrador" if es_borrador else "En revisión"
        )
        estado_destino = (
            "En elaboración" if es_borrador else "En revisión"
        )

        id_plantilla = texto(documento.get("ID_PLANTILLA"))
        plantilla = buscar_plantilla(id_plantilla)
        folder_id = texto(plantilla.get("CARPETA_DESTINO_ID"))
        if not folder_id:
            raise ValueError("La plantilla no tiene CARPETA_DESTINO_ID")

        titulo = texto(documento.get("TITULO")) or f"Documento_{id_documento}"
        if es_borrador:
            nombre_archivo = limpiar_nombre_archivo(
                f"{titulo}_V{numero_version_nueva:02d}_BORRADOR"
            )
        else:
            nombre_archivo = limpiar_nombre_archivo(
                f"{titulo}_V{numero_version_nueva:02d}"
                f"_REV{numero_revision_nueva:02d}"
            )

        fecha = ahora_iso()
        drive_service = obtener_drive_service()

        version_existente = buscar_version_numero_revision(
            id_documento=id_documento,
            numero_version=numero_version_nueva,
            numero_revision=numero_revision_nueva,
        )

        if version_existente:
            id_version_nueva = texto(
                version_existente.get("ID_VERSION")
            )
            copia = {
                "id": texto(version_existente.get("GOOGLE_DOC_ID")),
                "url": normalizar_url_appsheet(
                    version_existente.get("GOOGLE_DOC_URL")
                ),
                "name": (
                    texto(version_existente.get("NOMBRE_ARCHIVO"))
                    or nombre_archivo
                ),
            }
            if not copia["id"]:
                raise RuntimeError(
                    "La nueva versión existente no tiene GOOGLE_DOC_ID"
                )
        else:
            id_version_nueva = nuevo_id()
            copia = copiar_archivo_o_reutilizar(
                drive_service=drive_service,
                source_file_id=google_doc_id_rechazado,
                folder_id=folder_id,
                nombre_archivo=nombre_archivo,
            )

        # El archivo rechazado queda congelado para todos los participantes.
        emails_cadena = {
            texto(fila.get("APROBADOR")).lower()
            for fila in cadena_anterior
            if texto(fila.get("APROBADOR"))
            and "@" in texto(fila.get("APROBADOR"))
        }
        for email in emails_cadena:
            asegurar_permiso_rol(
                drive_service=drive_service,
                file_id=google_doc_id_rechazado,
                email=email,
                role="reader",
            )

        # En el nuevo archivo, el responsable anterior edita y quien rechazó
        # conserva permiso de comentario para responder observaciones.
        for fila in cadena_anterior[:indice_destino]:
            email_anterior = texto(fila.get("APROBADOR"))
            if email_anterior:
                asegurar_permiso_rol(
                    drive_service=drive_service,
                    file_id=copia["id"],
                    email=email_anterior,
                    role="reader",
                )

        permission_id_destino = asegurar_permiso_rol(
            drive_service=drive_service,
            file_id=copia["id"],
            email=email_destino,
            role="writer",
        )
        asegurar_permiso_rol(
            drive_service=drive_service,
            file_id=copia["id"],
            email=email_rechaza,
            role="commenter",
        )

        cadena_nueva_existente = buscar_cadena_documento_version(
            id_documento=id_documento,
            numero_version=numero_version_nueva,
        )

        if cadena_nueva_existente:
            destinos = [
                fila
                for fila in cadena_nueva_existente
                if entero(fila.get("ORDEN"), "ORDEN") == orden_destino
            ]
            if len(destinos) != 1:
                raise RuntimeError(
                    "La cadena nueva existente no tiene un único responsable "
                    f"de orden {orden_destino}"
                )
            destino_nuevo = destinos[0]
            actualizar_destino_cadena_reutilizada(
                destino=destino_nuevo,
                id_version_nueva=id_version_nueva,
                permission_id_destino=permission_id_destino,
                estado_destino=estado_destino,
                fecha=fecha,
            )
        else:
            filas_nuevas, destino_nuevo = (
                construir_cadena_nueva_por_rechazo(
                    cadena_anterior=cadena_anterior,
                    id_documento=id_documento,
                    id_plantilla=id_plantilla,
                    numero_version_anterior=numero_version_anterior,
                    numero_version_nueva=numero_version_nueva,
                    indice_destino=indice_destino,
                    id_version_nueva=id_version_nueva,
                    permission_id_destino=permission_id_destino,
                    fecha=fecha,
                )
            )
            appsheet_action(
                TABLA_APROBADORES_ACTUAL,
                "Add",
                filas_nuevas,
            )

        if not version_existente:
            crear_registro_version_por_aprobacion(
                id_version=id_version_nueva,
                id_documento=id_documento,
                id_version_origen=id_version_rechazada,
                numero_version=numero_version_nueva,
                numero_revision=numero_revision_nueva,
                etapa=etapa_nueva,
                nombre_archivo=copia["name"],
                google_doc_id=copia["id"],
                google_doc_url=copia["url"],
                id_aprobacion_responsable=destino_nuevo[
                    "ID_APROBACION_ACTUAL"
                ],
                orden_responsable=orden_destino,
                motivo_creacion="Rechazo",
                comentario=comentario,
                creado_por=usuario,
                fecha_creacion=fecha,
            )

        # Primero se cierra el historial rechazado. Si una operación posterior
        # falla, el endpoint puede reconstruir la transición usando la cadena
        # histórica y los recursos ya creados.
        actualizar_estado_version(
            id_version=id_version_rechazada,
            estado_version="Rechazada",
            fecha_cierre=fecha,
        )
        cerrar_cadena_anterior_por_rechazo(
            cadena_anterior=cadena_anterior,
            id_aprobacion_rechaza=id_aprobacion_solicitud,
            comentario=comentario,
            fecha=fecha,
        )

        actualizar_documento_rechazo_revision(
            id_documento=id_documento,
            numero_version=numero_version_nueva,
            numero_revision=numero_revision_nueva,
            id_version=id_version_nueva,
            copia=copia,
            destino=destino_nuevo,
            estado_documento=estado_documento_nuevo,
            usuario=usuario,
            fecha=fecha,
        )

        advertencias, evento_rechazo = crear_eventos_rechazo_revision(
            id_documento=id_documento,
            id_version_rechazada=id_version_rechazada,
            id_version_nueva=id_version_nueva,
            id_aprobacion_rechaza=id_aprobacion_solicitud,
            id_aprobacion_destino=destino_nuevo[
                "ID_APROBACION_ACTUAL"
            ],
            usuario=usuario,
            fecha=fecha,
            comentario=comentario,
            orden_rechaza=orden_rechaza,
            orden_destino=orden_destino,
            numero_version_nueva=numero_version_nueva,
            nombre_archivo=copia["name"],
            estado_nuevo=estado_documento_nuevo,
        )

        notificaciones: list[dict[str, Any]] = []
        if evento_rechazo is not None:
            try:
                documento_actualizado = buscar_documento(id_documento)
                cadena_nueva = buscar_cadena_documento_version(
                    id_documento=id_documento,
                    numero_version=numero_version_nueva,
                )
                aprobador_destino_actualizado = buscar_aprobacion_actual(
                    texto(documento_actualizado.get("ID_APROBACION_ACTUAL"))
                )
                notificaciones = ejecutar_notificaciones_rechazo_revision(
                    documento=documento_actualizado,
                    evento=evento_rechazo,
                    cadena=cadena_nueva,
                    aprobador_rechaza=aprobacion_actual,
                    aprobador_destino=aprobador_destino_actualizado,
                )
                fallidas = [
                    resultado
                    for resultado in notificaciones
                    if not resultado.get("ok")
                ]
                if fallidas:
                    advertencias.append(
                        f"{len(fallidas)} notificación(es) quedaron omitidas "
                        "o con error. Revisa Documento_Notificaciones."
                    )
            except Exception as exc_notificacion:
                traceback.print_exc()
                advertencias.append(
                    "El rechazo terminó correctamente, pero falló el proceso "
                    f"de notificaciones internas: {exc_notificacion}"
                )

        return jsonify(
            {
                "ok": True,
                "ya_procesado": False,
                "id_documento": id_documento,
                "estado": estado_documento_nuevo,
                "numero_version": numero_version_nueva,
                "numero_revision": numero_revision_nueva,
                "id_version": id_version_nueva,
                "google_doc_id": copia["id"],
                "google_doc_url": copia["url"],
                "nombre_archivo": copia["name"],
                "orden_actual": orden_destino,
                "id_aprobacion_actual": destino_nuevo[
                    "ID_APROBACION_ACTUAL"
                ],
                "encargado_actual": destino_nuevo.get("NOMBRE", ""),
                "encargado_email": email_destino,
                "notificaciones": notificaciones,
                "advertencias": advertencias,
            }
        )

    except PermissionError as exc:
        if id_documento:
            registrar_error_transicion(id_documento, str(exc))
        return {"error": str(exc)}, 403

    except (ValueError, LookupError) as exc:
        if id_documento:
            registrar_error_transicion(id_documento, str(exc))
        return {"error": str(exc)}, 400

    except Exception as exc:
        traceback.print_exc()
        if id_documento:
            registrar_error_transicion(id_documento, str(exc))
        return {"error": str(exc)}, 500


# -----------------------------------------------------------------------------
# Flujo: enviar a firma por correo
# -----------------------------------------------------------------------------


_EMAIL_RE = re.compile(
    r"^[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?"
    r"(?:\.[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?)+$",
    re.IGNORECASE,
)


def normalizar_destinatarios(valor: Any) -> list[str]:
    """Acepta EnumList de AppSheet, lista JSON o texto separado por comas."""
    elementos: list[Any]

    if isinstance(valor, (list, tuple, set)):
        elementos = list(valor)
    else:
        cadena = texto(valor)
        if not cadena:
            return []

        # Acepta una lista JSON cuando el webhook la envía como arreglo.
        if cadena.startswith("[") and cadena.endswith("]"):
            try:
                decodificado = json.loads(cadena)
            except json.JSONDecodeError:
                decodificado = None
            if isinstance(decodificado, list):
                elementos = decodificado
            else:
                elementos = re.split(r"[,;\n\r]+", cadena)
        else:
            elementos = re.split(r"[,;\n\r]+", cadena)

    resultado: list[str] = []
    vistos: set[str] = set()

    for elemento in elementos:
        email = texto(elemento).strip(' "\'<>')
        if not email:
            continue
        email_normalizado = email.lower()
        if not _EMAIL_RE.fullmatch(email_normalizado):
            raise ValueError(f"Correo destinatario no válido: {email!r}")
        if email_normalizado not in vistos:
            vistos.add(email_normalizado)
            resultado.append(email_normalizado)

    return resultado


def reemplazar_variables_email(
    contenido: str,
    variables: dict[str, Any],
) -> str:
    resultado = contenido
    for nombre, valor in variables.items():
        resultado = resultado.replace(
            "{{" + nombre + "}}",
            texto(valor),
        )
    return resultado


@medir_operacion("drive.descargar_pdf")
def descargar_pdf_drive(
    drive_service: Any,
    file_id: str,
) -> tuple[bytes, str]:
    metadata = (
        drive_service.files()
        .get(
            fileId=file_id,
            fields="id,name,mimeType,size",
            supportsAllDrives=True,
        )
        .execute()
    )

    mime_type = texto(metadata.get("mimeType"))
    if mime_type != "application/pdf":
        raise ValueError(
            "El archivo configurado en PDF_PARA_FIRMA_ID no es un PDF. "
            f"MIME encontrado: {mime_type!r}"
        )

    contenido = (
        drive_service.files()
        .get_media(
            fileId=file_id,
            supportsAllDrives=True,
        )
        .execute()
    )

    if not isinstance(contenido, bytes) or not contenido:
        raise RuntimeError("Google Drive devolvió un PDF vacío")

    nombre = texto(metadata.get("name")) or "documento_para_firma.pdf"
    if not nombre.lower().endswith(".pdf"):
        nombre += ".pdf"

    return contenido, nombre


@medir_operacion("drive.exportar_docx")
def exportar_docx_drive(
    drive_service: Any,
    google_doc_id: str,
    nombre_base: str,
) -> tuple[bytes, str]:
    """Exporta el Google Docs vigente a Microsoft Word sin guardarlo en Drive."""
    if not google_doc_id:
        raise ValueError("No se indicó GOOGLE_DOC_ID para exportar el DOCX")

    contenido = (
        drive_service.files()
        .export(
            fileId=google_doc_id,
            mimeType=DOCX_MIME_TYPE,
        )
        .execute()
    )

    if not isinstance(contenido, bytes) or not contenido:
        raise RuntimeError("Google Drive devolvió un DOCX vacío")

    nombre_sin_extension = re.sub(
        r"(?i)\.(pdf|docx)$",
        "",
        texto(nombre_base),
    )
    nombre = limpiar_nombre_archivo(
        nombre_sin_extension or "documento_para_firma"
    ) + ".docx"

    return contenido, nombre


def construir_email_firma_externo(
    *,
    documento: dict[str, Any],
    plantilla: dict[str, Any],
    usuario: str,
    mensaje_adicional: str,
    fecha_envio: str,
) -> tuple[str, str, str]:
    """Construye el asunto y los cuerpos texto/HTML del correo externo."""
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    titulo = texto(documento.get("TITULO")) or f"Documento {id_documento}"
    tipo_documento = texto(documento.get("TIPO_DOCUMENTO")) or "-"
    proyecto = (
        texto(documento.get("NOMBRE_PROYECTO"))
        or texto(documento.get("PROYECTO"))
        or texto(documento.get("ID_PROYECTO"))
    )

    mapa_nombres = construir_mapa_nombres_usuarios(
        id_documento=id_documento,
    )
    enviado_por_nombre = obtener_nombre_usuario_evento(
        usuario_evento=usuario,
        mapa_nombres=mapa_nombres,
    )
    if enviado_por_nombre == "Usuario no identificado":
        enviado_por_nombre = usuario

    variables = {
        "TITULO": titulo,
        "TIPO_DOCUMENTO": tipo_documento,
        "ID_PROYECTO": proyecto,
        "PROYECTO": proyecto,
        "ENVIADO_POR": usuario,
        "ENVIADO_POR_NOMBRE": enviado_por_nombre,
        "FECHA_ENVIO": fecha_envio,
    }

    asunto_base = texto(plantilla.get("ASUNTO_EMAIL_FIRMA"))
    if not asunto_base:
        asunto_base = "Solicitud de firma — {{TITULO}}"

    asunto = reemplazar_variables_email(
        asunto_base,
        variables,
    ).strip()

    cuerpo_base = texto(plantilla.get("CUERPO_EMAIL_FIRMA"))
    mensaje_plantilla = reemplazar_variables_email(
        cuerpo_base,
        variables,
    ).strip()

    # Las plantillas antiguas suelen repetir exactamente las instrucciones
    # estándar. En ese caso no se muestran nuevamente dentro del nuevo diseño.
    mensaje_normalizado = mensaje_plantilla.lower()
    es_mensaje_generico = (
        "adjuntamos el documento" in mensaje_normalizado
        and "una vez firmado" in mensaje_normalizado
        and "respondiendo a este correo" in mensaje_normalizado
    )
    if es_mensaje_generico:
        mensaje_plantilla = ""

    resumen_texto = [
        f"Documento: {titulo}",
        f"Tipo de documento: {tipo_documento}",
    ]
    if proyecto:
        resumen_texto.insert(1, f"Proyecto: {proyecto}")
    if enviado_por_nombre:
        resumen_texto.append(f"Enviado por: {enviado_por_nombre}")

    secciones_texto = [
        "SOLICITUD DE FIRMA",
        "",
        "Estimado/a:",
        "",
        f'Adjuntamos el documento "{titulo}" para su revisión y firma.',
        "",
        "RESUMEN DEL DOCUMENTO",
        *resumen_texto,
    ]

    if mensaje_plantilla:
        secciones_texto.extend(
            [
                "",
                "MENSAJE",
                mensaje_plantilla,
            ]
        )

    if mensaje_adicional:
        secciones_texto.extend(
            [
                "",
                "INDICACIONES ADICIONALES",
                mensaje_adicional,
            ]
        )

    secciones_texto.extend(
        [
            "",
            "¿QUÉ DEBE HACER?",
            "1. Revisar los archivos PDF y Word adjuntos.",
            "2. Firmar el documento utilizando el archivo PDF.",
            "3. Responder este mismo correo adjuntando el PDF firmado.",
            "",
            f"Este correo fue generado por {NOMBRE_APLICACION}.",
        ]
    )
    cuerpo_texto = "\n".join(secciones_texto)

    def escapar(valor: Any) -> str:
        return html.escape(texto(valor))

    def con_saltos(valor: Any) -> str:
        return escapar(valor).replace("\n", "<br>")

    filas_resumen = [
        ("Documento", titulo),
    ]
    if proyecto:
        filas_resumen.append(("Proyecto", proyecto))
    filas_resumen.extend(
        [
            ("Tipo de documento", tipo_documento),
            ("Enviado por", enviado_por_nombre),
        ]
    )

    filas_html = "".join(
        f"""
        <tr>
          <td style="
              padding:12px 16px;
              width:34%;
              border-top:1px solid #e5e7eb;
              color:#6b7280;
              font-size:14px;
              vertical-align:top;
          ">{escapar(etiqueta)}</td>
          <td style="
              padding:12px 16px;
              border-top:1px solid #e5e7eb;
              color:#111827;
              font-size:14px;
              font-weight:600;
              vertical-align:top;
          ">{escapar(valor)}</td>
        </tr>
        """
        for etiqueta, valor in filas_resumen
        if texto(valor)
    )

    bloque_mensaje = ""
    if mensaje_plantilla:
        bloque_mensaje = f"""
        <div style="
            margin:22px 0 0 0;
            padding:16px 18px;
            background:#f9fafb;
            border-left:4px solid #6b7280;
            border-radius:6px;
        ">
          <div style="
              margin-bottom:7px;
              color:#374151;
              font-size:13px;
              font-weight:700;
              text-transform:uppercase;
              letter-spacing:.3px;
          ">Mensaje</div>
          <div style="
              color:#374151;
              font-size:14px;
              line-height:1.65;
          ">{con_saltos(mensaje_plantilla)}</div>
        </div>
        """

    bloque_adicional = ""
    if mensaje_adicional:
        bloque_adicional = f"""
        <div style="
            margin:18px 0 0 0;
            padding:16px 18px;
            background:#fff7ed;
            border:1px solid #fed7aa;
            border-radius:8px;
        ">
          <div style="
              margin-bottom:7px;
              color:#9a3412;
              font-size:13px;
              font-weight:700;
              text-transform:uppercase;
              letter-spacing:.3px;
          ">Indicaciones adicionales</div>
          <div style="
              color:#7c2d12;
              font-size:14px;
              line-height:1.65;
          ">{con_saltos(mensaje_adicional)}</div>
        </div>
        """

    cuerpo_html = f"""
    <!doctype html>
    <html lang="es">
      <body style="
          margin:0;
          padding:0;
          background:#f3f4f6;
          font-family:Arial,Helvetica,sans-serif;
          color:#111827;
      ">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0"
               style="width:100%; background:#f3f4f6;">
          <tr>
            <td align="center" style="padding:24px 12px;">
              <table role="presentation" width="680" cellspacing="0" cellpadding="0"
                     style="
                         width:100%;
                         max-width:680px;
                         background:#ffffff;
                         border:1px solid #e5e7eb;
                         border-radius:10px;
                         overflow:hidden;
                     ">
                <tr>
                  <td style="padding:24px 28px; background:#111827; color:#ffffff;">
                    <div style="font-size:13px; color:#d1d5db;">
                      {escapar(NOMBRE_APLICACION)}
                    </div>
                    <div style="margin-top:5px; font-size:25px; font-weight:700;">
                      Solicitud de firma
                    </div>
                    <div style="margin-top:8px; font-size:14px; color:#d1d5db; line-height:1.5;">
                      Documento adjunto para revisión y firma
                    </div>
                  </td>
                </tr>

                <tr>
                  <td style="padding:28px;">
                    <p style="margin:0; font-size:15px; line-height:1.65;">
                      Estimado/a:
                    </p>
                    <p style="margin:16px 0 0 0; font-size:15px; line-height:1.65; color:#374151;">
                      Adjuntamos el documento
                      <strong>{escapar(titulo)}</strong>
                      para su revisión y firma.
                    </p>
                    <p style="margin:10px 0 0 0; font-size:15px; line-height:1.65; color:#374151;">
                      Se adjuntan el PDF para firma y una copia editable en formato Microsoft Word.
                    </p>

                    <div style="
                        margin-top:24px;
                        border:1px solid #e5e7eb;
                        border-radius:8px;
                        overflow:hidden;
                    ">
                      <div style="
                          padding:12px 16px;
                          background:#f9fafb;
                          color:#111827;
                          font-size:14px;
                          font-weight:700;
                      ">Resumen del documento</div>
                      <table role="presentation" width="100%" cellspacing="0" cellpadding="0"
                             style="width:100%; border-collapse:collapse;">
                        {filas_html}
                      </table>
                    </div>

                    {bloque_mensaje}
                    {bloque_adicional}

                    <div style="
                        margin-top:22px;
                        padding:17px 18px;
                        background:#eef2ff;
                        border:1px solid #c7d2fe;
                        border-radius:8px;
                    ">
                      <div style="
                          margin-bottom:10px;
                          color:#3730a3;
                          font-size:14px;
                          font-weight:700;
                      ">¿Qué debe hacer?</div>
                      <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
                        <tr>
                          <td style="padding:4px 8px 4px 0; width:24px; color:#4338ca; font-weight:700; vertical-align:top;">1.</td>
                          <td style="padding:4px 0; color:#374151; font-size:14px; line-height:1.5;">Revisar los archivos PDF y Word adjuntos.</td>
                        </tr>
                        <tr>
                          <td style="padding:4px 8px 4px 0; width:24px; color:#4338ca; font-weight:700; vertical-align:top;">2.</td>
                          <td style="padding:4px 0; color:#374151; font-size:14px; line-height:1.5;">Firmar el documento utilizando el archivo PDF.</td>
                        </tr>
                        <tr>
                          <td style="padding:4px 8px 4px 0; width:24px; color:#4338ca; font-weight:700; vertical-align:top;">3.</td>
                          <td style="padding:4px 0; color:#374151; font-size:14px; line-height:1.5;">Responder este mismo correo adjuntando el PDF firmado.</td>
                        </tr>
                      </table>
                    </div>

                    <p style="margin:26px 0 0 0; font-size:15px; color:#374151;">
                      Saludos.
                    </p>
                  </td>
                </tr>

                <tr>
                  <td style="
                      padding:16px 28px;
                      background:#f9fafb;
                      border-top:1px solid #e5e7eb;
                      color:#6b7280;
                      font-size:12px;
                      line-height:1.55;
                  ">
                    Este correo fue generado automáticamente por
                    <strong>{escapar(NOMBRE_APLICACION)}</strong>.<br>
                    Puede responder directamente a este mensaje con el documento firmado.
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </body>
    </html>
    """

    return asunto, cuerpo_texto, cuerpo_html


@medir_operacion("gmail.enviar_firma_externa")
def enviar_email_con_pdf(
    gmail_service: Any,
    destinatarios: list[str],
    asunto: str,
    cuerpo: str,
    pdf_bytes: bytes,
    pdf_nombre: str,
    docx_bytes: bytes,
    docx_nombre: str,
    reply_to: str = "",
    cuerpo_html: str = "",
) -> dict[str, str]:
    mensaje = EmailMessage()
    mensaje["To"] = ", ".join(destinatarios)
    mensaje["From"] = GMAIL_SENDER_EMAIL
    mensaje["Subject"] = asunto

    if reply_to and _EMAIL_RE.fullmatch(reply_to.lower()):
        mensaje["Reply-To"] = reply_to

    mensaje.set_content(cuerpo)
    if cuerpo_html:
        mensaje.add_alternative(cuerpo_html, subtype="html")

    mensaje.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_nombre,
    )
    mensaje.add_attachment(
        docx_bytes,
        maintype="application",
        subtype=(
            "vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
        filename=docx_nombre,
    )

    raw = base64.urlsafe_b64encode(mensaje.as_bytes()).decode("ascii")
    respuesta = (
        gmail_service.users()
        .messages()
        .send(
            userId="me",
            body={"raw": raw},
        )
        .execute()
    )

    message_id = texto(respuesta.get("id"))
    if not message_id:
        raise RuntimeError("Gmail no devolvió el ID del mensaje enviado")

    return {
        "message_id": message_id,
        "thread_id": texto(respuesta.get("threadId")),
    }

# -----------------------------------------------------------------------------
# Notificaciones internas por email - Fases 2 y 3
# -----------------------------------------------------------------------------

TIPOS_NOTIFICACION_VALIDOS = {
    "Acción requerida",
    "Informativa",
    "Confirmación",
    "Cierre",
}

ESTADOS_NOTIFICACION_REINTENTABLES = {
    "Pendiente",
    "Error",
}

# Eventos que aportan información relevante en los emails.
# Los demás eventos continúan guardados en Documento_Eventos,
# pero no se muestran en el correo.
EVENTOS_HISTORIAL_EMAIL = {
    "Enviado a revisión",
    "Revisión aprobada",
    "Revisión rechazada",
    "Proceso reiniciado",
    "Enviado a firma",
    "Proceso terminado",
}

# Cantidad máxima de eventos anteriores mostrados en cada correo.
MAX_EVENTOS_HISTORIAL_EMAIL = 8

def construir_link_appsheet(id_documento: str) -> str:
    """
    Construye un enlace web hacia la fila de Documentos.

    APPSHEET_DOCUMENT_VIEW_URL debe contener la URL completa de la vista
    Detail, por ejemplo:
    https://www.appsheet.com/start/...#view=Documentos_Notificacion_Detail
    """
    base = APPSHEET_DOCUMENT_VIEW_URL.strip()
    if not base:
        raise RuntimeError(
            "APPSHEET_DOCUMENT_VIEW_URL no está configurada"
        )

    id_codificado = quote(texto(id_documento), safe="")
    if not id_codificado:
        raise ValueError("No se puede construir el enlace sin ID_DOCUMENTO")

    # Reemplaza un parámetro row existente para evitar enlaces ambiguos.
    if re.search(r"([&#])row=[^&#]*", base, flags=re.IGNORECASE):
        return re.sub(
            r"([&#])row=[^&#]*",
            lambda coincidencia: (
                f"{coincidencia.group(1)}row={id_codificado}"
            ),
            base,
            count=1,
            flags=re.IGNORECASE,
        )

    if "#" in base:
        separador = "&"
    else:
        separador = "#"

    return f"{base}{separador}row={id_codificado}"


def obtener_cadena_notificacion(
    id_documento: str,
    numero_version: int | None = None,
) -> list[dict[str, Any]]:
    """
    Obtiene destinatarios exclusivamente desde
    Documentos_Aprobadores_Actual.
    """
    selector = (
        f"FILTER({TABLA_APROBADORES_ACTUAL}, "
        f"[ID_DOCUMENTO] = {literal_appsheet(id_documento)})"
    )
    filas = appsheet_find(TABLA_APROBADORES_ACTUAL, selector)

    resultado: list[dict[str, Any]] = []
    for fila in filas:
        if numero_version is not None:
            try:
                version_fila = entero(
                    fila.get("NUMERO_VERSION"),
                    "NUMERO_VERSION",
                )
            except ValueError:
                continue

            if version_fila != numero_version:
                continue

        resultado.append(fila)

    return sorted(
        resultado,
        key=lambda fila: entero(
            fila.get("ORDEN"),
            "ORDEN",
        ),
    )

def construir_mapa_nombres_usuarios(
    id_documento: str,
) -> dict[str, str]:
    """
    Construye un mapa para transformar los correos guardados en
    Documento_Eventos[USUARIO] en nombres de personas.

    Revisa tanto:
    - APROBADOR: correo operativo utilizado por el flujo y Drive.
    - Aprobador_v: correo utilizado para las notificaciones.
    """
    cadena = obtener_cadena_notificacion(
        id_documento=id_documento,
    )

    mapa: dict[str, str] = {}

    for fila in cadena:
        nombre = texto(fila.get("NOMBRE")).strip()

        if not nombre:
            continue

        for columna_email in ("APROBADOR", "Aprobador_v"):
            email = texto(
                fila.get(columna_email)
            ).strip().lower()

            if email:
                mapa[email] = nombre

    # Cuando el correo emisor ejecuta una acción y no pertenece a la cadena,
    # se muestra el nombre de la aplicación en vez del email.
    if GMAIL_SENDER_EMAIL:
        mapa.setdefault(
            GMAIL_SENDER_EMAIL.strip().lower(),
            NOMBRE_APLICACION,
        )

    return mapa

def obtener_nombre_usuario_evento(
    usuario_evento: Any,
    mapa_nombres: dict[str, str],
) -> str:
    """
    Devuelve el nombre asociado al usuario que ejecutó el evento.

    Si USUARIO ya contiene un nombre, lo conserva.
    Si contiene un email, busca su nombre en la cadena.
    """
    usuario = texto(usuario_evento).strip()

    if not usuario:
        return "Sistema"

    usuario_normalizado = usuario.lower()

    nombre = mapa_nombres.get(usuario_normalizado)
    if nombre:
        return nombre

    # Si el valor no es un email, probablemente ya corresponde a un nombre.
    if "@" not in usuario:
        return usuario

    # No mostramos el correo en el email cuando no se encuentra el nombre.
    return "Usuario no identificado"

def obtener_email_notificacion(
    aprobador: dict[str, Any],
) -> str:
    """
    Obtiene el correo usado exclusivamente para las notificaciones internas.
    La fuente es la columna virtual Aprobador_v de
    Documentos_Aprobadores_Actual.
    """
    return texto(
        aprobador.get("Aprobador_v")
    ).strip().lower()

def deduplicar_cadena_por_email(
    cadena: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Elimina correos repetidos y omite direcciones inválidas."""
    resultado: list[dict[str, Any]] = []
    vistos: set[str] = set()

    for fila in cadena:
        email = obtener_email_notificacion(fila)
        if not email:
            app.logger.warning(
                "Se omitió un responsable sin Aprobador_v: %s",
                fila.get("ID_APROBACION_ACTUAL"),
            )
            continue

        if not _EMAIL_RE.fullmatch(email):
            app.logger.warning(
                "Se omitió un responsable con correo inválido: %s",
                email,
            )
            continue

        if email in vistos:
            continue

        vistos.add(email)
        resultado.append(fila)

    return resultado


def buscar_eventos_documento(
    id_documento: str,
) -> list[dict[str, Any]]:
    selector = (
        f"FILTER({TABLA_EVENTOS}, "
        f"[ID_DOCUMENTO] = {literal_appsheet(id_documento)})"
    )
    return appsheet_find(TABLA_EVENTOS, selector)


def buscar_evento_por_id(
    id_evento: str,
) -> dict[str, Any]:
    selector = (
        f"FILTER({TABLA_EVENTOS}, "
        f"[ID_EVENTO] = {literal_appsheet(id_evento)})"
    )
    filas = appsheet_find(TABLA_EVENTOS, selector)
    if not filas:
        raise LookupError(f"No se encontró ID_EVENTO={id_evento}")
    return filas[0]


def parsear_fecha_appsheet(valor: Any) -> datetime:
    fecha_texto = texto(valor)
    if not fecha_texto:
        return datetime.min

    formatos = (
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%dT%H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
    )

    for formato in formatos:
        try:
            fecha = datetime.strptime(fecha_texto, formato)
            if fecha.tzinfo is not None:
                fecha = fecha.astimezone(CHILE_TZ).replace(tzinfo=None)
            return fecha
        except ValueError:
            continue

    return datetime.min

def formatear_fecha_historial(valor: Any) -> str:
    """
    Convierte la fecha recibida desde AppSheet al formato chileno
    utilizado en los emails.
    """
    fecha = parsear_fecha_appsheet(valor)

    if fecha == datetime.min:
        return texto(valor)

    return fecha.strftime("%d/%m/%Y %H:%M")

def construir_historial_comentarios(
    id_documento: str,
    id_evento_excluir: str = "",
    mapa_nombres: dict[str, str] | None = None,
) -> tuple[str, str]:
    """
    Construye una versión resumida del historial para los emails.

    Reglas:
    - Solo muestra eventos relevantes.
    - Solo muestra eventos con comentario.
    - Excluye el evento actual, porque ya aparece destacado arriba.
    - Limita el historial a los últimos eventos configurados.
    - Mantiene todos los eventos originales en Documento_Eventos.
    """
    eventos = buscar_eventos_documento(id_documento)

    if mapa_nombres is None:
        mapa_nombres = construir_mapa_nombres_usuarios(
            id_documento=id_documento,
        )

    eventos_relevantes: list[dict[str, Any]] = []

    for evento in eventos:
        id_evento = texto(evento.get("ID_EVENTO"))
        tipo_evento = texto(evento.get("TIPO_EVENTO"))
        comentario = texto(evento.get("COMENTARIO"))

        # El evento actual ya se muestra como "Último movimiento".
        if id_evento_excluir and id_evento == id_evento_excluir:
            continue

        # Omite eventos técnicos o administrativos.
        if tipo_evento not in EVENTOS_HISTORIAL_EMAIL:
            continue

        # No incorpora eventos sin comentario.
        if not comentario:
            continue

        eventos_relevantes.append(evento)

    # Ordena desde el evento más reciente al más antiguo.
    eventos_relevantes.sort(
        key=lambda fila: (
            parsear_fecha_appsheet(fila.get("FECHA_EVENTO")),
            texto(fila.get("ID_EVENTO")),
        ),
        reverse=True,
    )

    total_eventos_relevantes = len(eventos_relevantes)

    # Conserva únicamente los eventos relevantes más recientes.
    eventos_mostrados = eventos_relevantes[
        :MAX_EVENTOS_HISTORIAL_EMAIL
    ]

    cantidad_omitidos = (
        total_eventos_relevantes - len(eventos_mostrados)
    )

    lineas_texto: list[str] = []
    bloques_html: list[str] = []

    if cantidad_omitidos > 0:
        lineas_texto.append(
            f"Se muestran los últimos "
            f"{len(eventos_mostrados)} de "
            f"{total_eventos_relevantes} eventos relevantes. "
            f"El historial completo está disponible en AppSheet."
        )

        bloques_html.append(
            f"""
            <div style="
                margin:0 0 14px 0;
                padding:10px 12px;
                border-radius:6px;
                background:#f3f4f6;
                color:#4b5563;
                font-size:13px;
            ">
                Se muestran los últimos
                <strong>{len(eventos_mostrados)}</strong>
                de
                <strong>{total_eventos_relevantes}</strong>
                eventos relevantes.
                El historial completo está disponible en AppSheet.
            </div>
            """
        )

    for evento in eventos_mostrados:
        fecha = formatear_fecha_historial(
            evento.get("FECHA_EVENTO")
        )
        usuario = obtener_nombre_usuario_evento(
            usuario_evento=evento.get("USUARIO"),
            mapa_nombres=mapa_nombres,
        )
        tipo_evento = (
            texto(evento.get("TIPO_EVENTO"))
            or "Actualización"
        )
        comentario = texto(evento.get("COMENTARIO"))
        estado_nuevo = texto(evento.get("ESTADO_NUEVO"))

        movimiento = tipo_evento
        if estado_nuevo:
            movimiento += f" — {estado_nuevo}"

        # Versión compacta en texto plano.
        lineas_texto.append(
            f"{fecha} — {movimiento}\n"
            f"{usuario}\n"
            f"{comentario}"
        )

        # Versión compacta en HTML.
        bloques_html.append(
            """
            <div style="
                margin:0 0 10px 0;
                padding:9px 11px;
                border-left:3px solid #6b7280;
                background:#f9fafb;
            ">
                <div style="
                    font-size:13px;
                    font-weight:700;
                    color:#111827;
                ">
                    {fecha} — {movimiento}
                </div>

                <div style="
                    margin-top:2px;
                    font-size:12px;
                    color:#6b7280;
                ">
                    {usuario}
                </div>

                <div style="
                    margin-top:5px;
                    font-size:14px;
                    white-space:pre-wrap;
                ">
                    {comentario}
                </div>
            </div>
            """.format(
                fecha=html.escape(fecha),
                movimiento=html.escape(movimiento),
                usuario=html.escape(usuario),
                comentario=html.escape(comentario),
            )
        )

    if not eventos_mostrados:
        return (
            "No existen comentarios anteriores relevantes.",
            (
                "<p style='color:#6b7280;'>"
                "No existen comentarios anteriores relevantes."
                "</p>"
            ),
        )

    return (
        "\n\n".join(lineas_texto),
        "".join(bloques_html),
    )

def construir_email_notificacion(
    *,
    documento: dict[str, Any],
    destinatario: dict[str, Any],
    tipo_notificacion: str,
    movimiento: str,
    comentario_principal: str,
    historial_texto: str,
    historial_html: str,
    link_documento: str,
    link_appsheet: str,
) -> tuple[str, str, str]:
    if tipo_notificacion not in TIPOS_NOTIFICACION_VALIDOS:
        raise ValueError(
            f"TIPO_NOTIFICACION no válido: {tipo_notificacion!r}"
        )

    id_documento = texto(documento.get("ID_DOCUMENTO"))
    titulo = texto(documento.get("TITULO")) or id_documento
    tipo_documento = texto(documento.get("TIPO_DOCUMENTO"))
    estado = texto(documento.get("ESTADO"))
    version = texto(documento.get("VERSION_ACTUAL")) or "-"
    revision = texto(documento.get("REVISION_ACTUAL")) or "-"
    responsable = texto(
        documento.get("ENCARGADO_ACTUAL_NOMBRE")
    )
    nombre_destinatario = (
        texto(destinatario.get("NOMBRE"))
        or texto(destinatario.get("DESTINATARIO_NOMBRE"))
        or "usuario/a"
    )

    prefijo = {
        "Acción requerida": "Acción requerida",
        "Confirmación": "Confirmación",
        "Informativa": "Información",
        "Cierre": "Proceso documental",
    }[tipo_notificacion]

    # El asunto debe ser idéntico para todas las notificaciones del documento.
    asunto = f"Seguimiento documental — {titulo}"

    comentario_texto = (
        comentario_principal
        or "No se registró un comentario adicional."
    )

    lineas_enlaces = [f"Abrir en AppSheet: {link_appsheet}"]
    if tipo_notificacion == "Acción requerida" and link_documento:
        lineas_enlaces.insert(
            0,
            f"Abrir documento: {link_documento}",
        )

    cuerpo_texto = (
        f"{tipo_notificacion.upper()}\n\n"
        f"Estimado/a {nombre_destinatario}:\n\n"
        f"Se registró un movimiento en {NOMBRE_APLICACION}.\n\n"
        f"Documento: {titulo}\n"
        f"Tipo: {tipo_documento}\n"
        f"Estado actual: {estado}\n"
        f"Versión: V{version} — REV{revision}\n"
        f"Responsable actual: {responsable}\n\n"
        f"Último movimiento:\n{movimiento}\n\n"
        f"Comentario:\n{comentario_texto}\n\n"
        + "\n".join(lineas_enlaces)
        + "\n\nHISTORIAL DEL FLUJO\n\n"
        + historial_texto
        + f"\n\nEste mensaje fue enviado por {NOMBRE_APLICACION}."
    )

    botones_html: list[str] = []
    if tipo_notificacion == "Acción requerida" and link_documento:
        botones_html.append(
            """
            <a href="{url}" style="
                display:inline-block;
                padding:11px 18px;
                margin:4px 8px 4px 0;
                background:#4f46e5;
                color:#ffffff;
                text-decoration:none;
                border-radius:6px;
                font-weight:700;
            ">Abrir documento</a>
            """.format(url=html.escape(link_documento, quote=True))
        )

    botones_html.append(
        """
        <a href="{url}" style="
            display:inline-block;
            padding:11px 18px;
            margin:4px 8px 4px 0;
            background:#111827;
            color:#ffffff;
            text-decoration:none;
            border-radius:6px;
            font-weight:700;
        ">Abrir en AppSheet</a>
        """.format(url=html.escape(link_appsheet, quote=True))
    )

    cuerpo_html = """
    <html>
      <body style="
          margin:0;
          padding:0;
          background:#f3f4f6;
          font-family:Arial,Helvetica,sans-serif;
          color:#111827;
      ">
        <div style="
            max-width:720px;
            margin:24px auto;
            background:#ffffff;
            border:1px solid #e5e7eb;
            border-radius:10px;
            overflow:hidden;
        ">
          <div style="
              padding:20px 24px;
              background:#111827;
              color:#ffffff;
          ">
            <div style="font-size:13px; opacity:.85;">{aplicacion}</div>
            <div style="font-size:22px; font-weight:700; margin-top:4px;">
              {tipo_notificacion}
            </div>
          </div>

          <div style="padding:24px;">
            <p>Estimado/a <strong>{nombre_destinatario}</strong>:</p>
            <p>Se registró un movimiento en el flujo documental.</p>

            <table style="
                width:100%;
                border-collapse:collapse;
                margin:18px 0;
            ">
              <tr><td style="{td_label}">Documento</td><td style="{td_value}">{titulo}</td></tr>
              <tr><td style="{td_label}">Tipo</td><td style="{td_value}">{tipo_documento}</td></tr>
              <tr><td style="{td_label}">Estado</td><td style="{td_value}">{estado}</td></tr>
              <tr><td style="{td_label}">Versión</td><td style="{td_value}">V{version} — REV{revision}</td></tr>
              <tr><td style="{td_label}">Responsable actual</td><td style="{td_value}">{responsable}</td></tr>
            </table>

            <div style="
                padding:14px 16px;
                margin:18px 0;
                border-radius:8px;
                background:#eef2ff;
            ">
              <div style="font-weight:700; margin-bottom:5px;">
                Último movimiento
              </div>
              <div>{movimiento}</div>
            </div>

            <div style="
                padding:14px 16px;
                margin:18px 0;
                border-radius:8px;
                background:#fff7ed;
            ">
              <div style="font-weight:700; margin-bottom:5px;">
                Comentario
              </div>
              <div style="white-space:pre-wrap;">{comentario_principal}</div>
            </div>

            <div style="margin:22px 0;">
              {botones}
            </div>

            <h3 style="margin-top:28px;">Historial del flujo</h3>
            {historial_html}
          </div>

          <div style="
              padding:14px 24px;
              background:#f9fafb;
              color:#6b7280;
              font-size:12px;
          ">
            Mensaje automático de {aplicacion}.
          </div>
        </div>
      </body>
    </html>
    """.format(
        aplicacion=html.escape(NOMBRE_APLICACION),
        tipo_notificacion=html.escape(tipo_notificacion),
        nombre_destinatario=html.escape(nombre_destinatario),
        titulo=html.escape(titulo),
        tipo_documento=html.escape(tipo_documento),
        estado=html.escape(estado),
        version=html.escape(version),
        revision=html.escape(revision),
        responsable=html.escape(responsable),
        movimiento=html.escape(movimiento),
        comentario_principal=html.escape(comentario_texto),
        botones="".join(botones_html),
        historial_html=historial_html,
        td_label=(
            "padding:8px 10px;border-bottom:1px solid #e5e7eb;"
            "font-weight:700;width:35%;vertical-align:top;"
        ),
        td_value=(
            "padding:8px 10px;border-bottom:1px solid #e5e7eb;"
            "vertical-align:top;"
        ),
    )

    return asunto, cuerpo_texto, cuerpo_html


def construir_clave_idempotencia(
    id_evento: str,
    email: str,
    tipo_notificacion: str,
) -> str:
    id_evento_limpio = texto(id_evento)
    email_limpio = texto(email).lower()
    tipo_limpio = texto(tipo_notificacion)

    if not id_evento_limpio:
        raise ValueError(
            "ID_EVENTO es obligatorio para la idempotencia"
        )

    return (
        f"{id_evento_limpio}|"
        f"{email_limpio}|"
        f"{tipo_limpio}"
    )


def buscar_notificacion_por_clave(
    clave: str,
) -> dict[str, Any] | None:
    selector = (
        f"FILTER({TABLA_NOTIFICACIONES}, "
        f"[CLAVE_IDEMPOTENCIA] = {literal_appsheet(clave)})"
    )
    filas = appsheet_find(TABLA_NOTIFICACIONES, selector)
    if len(filas) > 1:
        raise RuntimeError(
            "Existen varias notificaciones con la misma "
            "CLAVE_IDEMPOTENCIA"
        )
    return filas[0] if filas else None


def crear_notificacion_pendiente(
    *,
    id_evento: str,
    id_documento: str,
    id_version: str,
    aprobador: dict[str, Any],
    tipo_notificacion: str,
    asunto: str,
    cuerpo: str,
    link_documento: str,
    link_appsheet: str,
) -> tuple[dict[str, Any], bool]:
    email = obtener_email_notificacion(aprobador)
    clave = construir_clave_idempotencia(
        id_evento=id_evento,
        email=email,
        tipo_notificacion=tipo_notificacion,
    )

    existente = buscar_notificacion_por_clave(clave)
    if existente:
        return existente, False

    fila = {
        "ID_NOTIFICACION": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_EVENTO": id_evento,
        "ID_VERSION": id_version,
        "ID_APROBACION_ACTUAL": texto(
            aprobador.get("ID_APROBACION_ACTUAL")
        ),
        "DESTINATARIO_EMAIL": email,
        "DESTINATARIO_NOMBRE": texto(
            aprobador.get("NOMBRE")
        ),
        "ORDEN_DESTINATARIO": aprobador.get("ORDEN", ""),
        "ROL_DESTINATARIO": texto(
            aprobador.get("ROL_FLUJO")
        ),
        "TIPO_NOTIFICACION": tipo_notificacion,
        "ASUNTO": asunto,
        "CUERPO": cuerpo,
        "LINK_DOCUMENTO": link_documento,
        "LINK_APPSHEET": link_appsheet,
        "ESTADO_ENVIO": "Pendiente",
        "INTENTOS": 0,
        "FECHA_CREACION": ahora_iso(),
        "CLAVE_IDEMPOTENCIA": clave,
    }

    appsheet_action(
        TABLA_NOTIFICACIONES,
        "Add",
        [fila],
    )
    return fila, True


def construir_rfc_message_id_notificacion(
    id_notificacion: str,
) -> str:
    """Genera un Message-ID RFC estable para una notificación."""
    identificador = re.sub(
        r"[^A-Za-z0-9_.-]+",
        "-",
        texto(id_notificacion),
    ).strip("-")
    if not identificador:
        raise ValueError("ID_NOTIFICACION vacío para construir Message-ID")

    dominio = "gmail.com"
    if "@" in GMAIL_SENDER_EMAIL:
        dominio = GMAIL_SENDER_EMAIL.rsplit("@", 1)[1].lower()

    return f"<notificacion-{identificador}@{dominio}>"


def buscar_ultima_notificacion_hilo(
    *,
    id_documento: str,
    destinatario_email: str,
    asunto: str,
    id_notificacion_excluir: str,
) -> dict[str, Any] | None:
    """
    Busca el último mensaje enviado del mismo documento, destinatario y asunto.

    Solo considera asuntos del nuevo formato. Esto evita intentar enlazar los
    correos históricos que fueron enviados antes de incorporar Message-ID.
    """
    selector = (
        f"FILTER({TABLA_NOTIFICACIONES}, "
        f"[ID_DOCUMENTO] = {literal_appsheet(id_documento)})"
    )
    filas = appsheet_find(TABLA_NOTIFICACIONES, selector)

    email_objetivo = texto(destinatario_email).lower()
    candidatas = [
        fila
        for fila in filas
        if texto(fila.get("ID_NOTIFICACION")) != id_notificacion_excluir
        and texto(fila.get("DESTINATARIO_EMAIL")).lower() == email_objetivo
        and texto(fila.get("ESTADO_ENVIO")) == "Enviada"
        and texto(fila.get("ASUNTO")) == asunto
        and texto(fila.get("GMAIL_THREAD_ID"))
    ]

    if not candidatas:
        return None

    candidatas.sort(
        key=lambda fila: (
            parsear_fecha_appsheet(fila.get("FECHA_ENVIO")),
            texto(fila.get("ID_NOTIFICACION")),
        ),
        reverse=True,
    )
    return candidatas[0]


@medir_operacion("gmail.enviar_notificacion_interna")
def enviar_email_notificacion(
    *,
    gmail_service: Any,
    destinatario: str,
    asunto: str,
    cuerpo_texto: str,
    cuerpo_html: str,
    rfc_message_id: str,
    thread_id: str = "",
    in_reply_to: str = "",
) -> dict[str, str]:
    email = texto(destinatario).lower()
    if not _EMAIL_RE.fullmatch(email):
        raise ValueError(
            f"Correo de notificación no válido: {destinatario!r}"
        )

    mensaje = EmailMessage()
    mensaje["To"] = email
    mensaje["From"] = GMAIL_SENDER_EMAIL
    mensaje["Subject"] = asunto
    mensaje["Message-ID"] = rfc_message_id

    # Gmail exige threadId, In-Reply-To, References y asunto coincidente para
    # incorporar el mensaje a una conversación existente.
    usar_hilo = bool(thread_id and in_reply_to)
    if usar_hilo:
        mensaje["In-Reply-To"] = in_reply_to
        mensaje["References"] = in_reply_to

    mensaje.set_content(cuerpo_texto)
    mensaje.add_alternative(cuerpo_html, subtype="html")

    raw = base64.urlsafe_b64encode(
        mensaje.as_bytes()
    ).decode("ascii")

    body_gmail: dict[str, str] = {"raw": raw}
    if usar_hilo:
        body_gmail["threadId"] = thread_id

    respuesta = (
        gmail_service.users()
        .messages()
        .send(
            userId="me",
            body=body_gmail,
        )
        .execute()
    )

    message_id = texto(respuesta.get("id"))
    if not message_id:
        raise RuntimeError(
            "Gmail no devolvió el ID del mensaje de notificación"
        )

    return {
        "message_id": message_id,
        "thread_id": texto(respuesta.get("threadId")),
        "rfc_message_id": rfc_message_id,
    }

def marcar_notificacion_enviada(
    *,
    id_notificacion: str,
    intentos: int,
    message_id: str,
    thread_id: str,
) -> None:
    appsheet_action(
        TABLA_NOTIFICACIONES,
        "Edit",
        [
            {
                "ID_NOTIFICACION": id_notificacion,
                "ESTADO_ENVIO": "Enviada",
                "INTENTOS": intentos + 1,
                "FECHA_ENVIO": ahora_iso(),
                "ERROR_ENVIO": "",
                "GMAIL_MESSAGE_ID": message_id,
                "GMAIL_THREAD_ID": thread_id,
            }
        ],
    )


def marcar_notificacion_error(
    *,
    id_notificacion: str,
    intentos: int,
    mensaje_error: str,
) -> None:
    appsheet_action(
        TABLA_NOTIFICACIONES,
        "Edit",
        [
            {
                "ID_NOTIFICACION": id_notificacion,
                "ESTADO_ENVIO": "Error",
                "INTENTOS": intentos + 1,
                "ERROR_ENVIO": texto(mensaje_error)[:1500],
            }
        ],
    )


def actualizar_resumen_notificacion_documento(
    *,
    id_documento: str,
    message_id: str,
    thread_id: str,
) -> None:
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "GMAIL_THREAD_ID_FLUJO": thread_id,
                "GMAIL_ULTIMO_MESSAGE_ID_FLUJO": message_id,
                "FECHA_ULTIMA_NOTIFICACION": ahora_iso(),
            }
        ],
    )


def procesar_notificacion_individual(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    aprobador: dict[str, Any],
    tipo_notificacion: str,
    movimiento: str,
    comentario_principal: str,
    link_documento: str,
) -> dict[str, Any]:
    """
    Crea, envía y registra una notificación sin propagar errores al flujo
    documental que la invoque.
    """
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    id_evento = texto(evento.get("ID_EVENTO"))
    id_version = (
        texto(evento.get("ID_VERSION"))
        or texto(documento.get("ID_VERSION_ACTUAL"))
    )
    email = obtener_email_notificacion(aprobador)

    # Por diseño, solo las notificaciones de Acción requerida guardan
    # y muestran el enlace directo al documento editable.
    link_documento = (
        normalizar_url_appsheet(link_documento)
        if tipo_notificacion == "Acción requerida"
        else ""
    )

    if not email or not _EMAIL_RE.fullmatch(email):
        return {
            "ok": False,
            "omitida": True,
            "destinatario": email,
            "mensaje": "Correo interno inválido o vacío.",
        }

    try:
        link_appsheet = construir_link_appsheet(id_documento)
        historial_texto, historial_html = (
          construir_historial_comentarios(
            id_documento=id_documento,
            id_evento_excluir=id_evento,
            )
        )

        asunto, cuerpo_texto, cuerpo_html = (
            construir_email_notificacion(
                documento=documento,
                destinatario=aprobador,
                tipo_notificacion=tipo_notificacion,
                movimiento=movimiento,
                comentario_principal=comentario_principal,
                historial_texto=historial_texto,
                historial_html=historial_html,
                link_documento=link_documento,
                link_appsheet=link_appsheet,
            )
        )

        notificacion, creada = crear_notificacion_pendiente(
            id_evento=id_evento,
            id_documento=id_documento,
            id_version=id_version,
            aprobador=aprobador,
            tipo_notificacion=tipo_notificacion,
            asunto=asunto,
            cuerpo=cuerpo_texto,
            link_documento=link_documento,
            link_appsheet=link_appsheet,
        )

        estado_existente = texto(
            notificacion.get("ESTADO_ENVIO")
        )
        id_notificacion = texto(
            notificacion.get("ID_NOTIFICACION")
        )
        intentos = 0
        try:
            intentos = entero(
                notificacion.get("INTENTOS") or 0,
                "INTENTOS",
            )
        except ValueError:
            intentos = 0

        if not creada and estado_existente == "Enviada":
            return {
                "ok": True,
                "omitida": True,
                "duplicada": True,
                "id_notificacion": id_notificacion,
                "destinatario": email,
                "mensaje": "La notificación ya había sido enviada.",
            }

        if (
            not creada
            and estado_existente
            not in ESTADOS_NOTIFICACION_REINTENTABLES
        ):
            return {
                "ok": False,
                "omitida": True,
                "id_notificacion": id_notificacion,
                "destinatario": email,
                "mensaje": (
                    "La notificación existente no permite reintento. "
                    f"Estado: {estado_existente!r}"
                ),
            }

        rfc_message_id = construir_rfc_message_id_notificacion(
            id_notificacion
        )
        notificacion_anterior = buscar_ultima_notificacion_hilo(
            id_documento=id_documento,
            destinatario_email=email,
            asunto=asunto,
            id_notificacion_excluir=id_notificacion,
        )

        thread_id_anterior = ""
        in_reply_to = ""
        if notificacion_anterior:
            thread_id_anterior = texto(
                notificacion_anterior.get("GMAIL_THREAD_ID")
            )
            id_notificacion_anterior = texto(
                notificacion_anterior.get("ID_NOTIFICACION")
            )
            if id_notificacion_anterior:
                in_reply_to = construir_rfc_message_id_notificacion(
                    id_notificacion_anterior
                )

        gmail_service = obtener_gmail_service()
        respuesta_gmail = enviar_email_notificacion(
            gmail_service=gmail_service,
            destinatario=email,
            asunto=asunto,
            cuerpo_texto=cuerpo_texto,
            cuerpo_html=cuerpo_html,
            rfc_message_id=rfc_message_id,
            thread_id=thread_id_anterior,
            in_reply_to=in_reply_to,
        )

        marcar_notificacion_enviada(
            id_notificacion=id_notificacion,
            intentos=intentos,
            message_id=respuesta_gmail["message_id"],
            thread_id=respuesta_gmail["thread_id"],
        )

        try:
            actualizar_resumen_notificacion_documento(
                id_documento=id_documento,
                message_id=respuesta_gmail["message_id"],
                thread_id=respuesta_gmail["thread_id"],
            )
        except Exception:
            traceback.print_exc()

        return {
            "ok": True,
            "omitida": False,
            "id_notificacion": id_notificacion,
            "destinatario": email,
            "tipo_notificacion": tipo_notificacion,
            "message_id": respuesta_gmail["message_id"],
            "thread_id": respuesta_gmail["thread_id"],
        }

    except Exception as exc:
        traceback.print_exc()

        # Si ya alcanzamos a crear la fila, registra el error en ella.
        try:
            id_notificacion_local = texto(
                locals().get("id_notificacion")
            )
            intentos_local = locals().get("intentos", 0)
            if id_notificacion_local:
                marcar_notificacion_error(
                    id_notificacion=id_notificacion_local,
                    intentos=int(intentos_local),
                    mensaje_error=str(exc),
                )
        except Exception:
            traceback.print_exc()

        return {
            "ok": False,
            "omitida": False,
            "destinatario": email,
            "tipo_notificacion": tipo_notificacion,
            "error": str(exc),
        }


def notificar_destinatarios_internos_legacy(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    destinatarios: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Procesa una lista de especificaciones. Cada elemento debe contener:
    aprobador, tipo_notificacion, movimiento, comentario_principal y
    link_documento.
    """
    resultados: list[dict[str, Any]] = []
    correos_procesados: set[str] = set()

    for especificacion in destinatarios:
        aprobador = especificacion.get("aprobador")
        if not isinstance(aprobador, dict):
            resultados.append(
                {
                    "ok": False,
                    "omitida": True,
                    "error": "Especificación sin aprobador válido.",
                }
            )
            continue

        email = obtener_email_notificacion(aprobador)
        if not email:
            resultados.append(
                {
                    "ok": False,
                    "omitida": True,
                    "destinatario": "",
                    "tipo_notificacion": texto(
                        especificacion.get("tipo_notificacion")
                    ),
                    "mensaje": (
                        "El responsable no tiene Aprobador_v para enviar "
                        "la notificación."
                    ),
                }
            )
            continue

        if email in correos_procesados:
            continue

        correos_procesados.add(email)
        resultados.append(
            procesar_notificacion_individual(
                documento=documento,
                evento=evento,
                aprobador=aprobador,
                tipo_notificacion=texto(
                    especificacion.get("tipo_notificacion")
                ),
                movimiento=texto(
                    especificacion.get("movimiento")
                ),
                comentario_principal=texto(
                    especificacion.get("comentario_principal")
                ),
                link_documento=normalizar_url_appsheet(
                    especificacion.get("link_documento")
                ),
            )
        )

    return resultados



def construir_mapa_nombres_desde_aprobadores(
    aprobadores: list[dict[str, Any]],
) -> dict[str, str]:
    """Construye el mapa de nombres sin consultar nuevamente AppSheet."""
    mapa: dict[str, str] = {}

    for fila in aprobadores:
        nombre = texto(fila.get("NOMBRE")).strip()
        if not nombre:
            continue

        for columna_email in ("APROBADOR", "Aprobador_v"):
            email = texto(fila.get(columna_email)).strip().lower()
            if email:
                mapa[email] = nombre

    if GMAIL_SENDER_EMAIL:
        mapa.setdefault(
            GMAIL_SENDER_EMAIL.strip().lower(),
            NOMBRE_APLICACION,
        )

    return mapa


def buscar_notificaciones_documento(
    id_documento: str,
) -> list[dict[str, Any]]:
    """Obtiene una sola vez todas las notificaciones del documento."""
    selector = (
        f"FILTER({TABLA_NOTIFICACIONES}, "
        f"[ID_DOCUMENTO] = {literal_appsheet(id_documento)})"
    )
    return appsheet_find(TABLA_NOTIFICACIONES, selector)


def seleccionar_ultima_notificacion_hilo_en_memoria(
    *,
    filas: list[dict[str, Any]],
    destinatario_email: str,
    asunto: str,
    id_notificacion_excluir: str,
) -> dict[str, Any] | None:
    """Busca el hilo anterior usando las filas ya cargadas en memoria."""
    email_objetivo = texto(destinatario_email).lower()
    candidatas = [
        fila
        for fila in filas
        if texto(fila.get("ID_NOTIFICACION"))
        != id_notificacion_excluir
        and texto(fila.get("DESTINATARIO_EMAIL")).lower()
        == email_objetivo
        and texto(fila.get("ESTADO_ENVIO")) == "Enviada"
        and texto(fila.get("ASUNTO")) == asunto
        and texto(fila.get("GMAIL_THREAD_ID"))
    ]

    if not candidatas:
        return None

    candidatas.sort(
        key=lambda fila: (
            parsear_fecha_appsheet(fila.get("FECHA_ENVIO")),
            texto(fila.get("ID_NOTIFICACION")),
        ),
        reverse=True,
    )
    return candidatas[0]


def construir_fila_notificacion_pendiente(
    *,
    id_evento: str,
    id_documento: str,
    id_version: str,
    aprobador: dict[str, Any],
    tipo_notificacion: str,
    asunto: str,
    cuerpo: str,
    link_documento: str,
    link_appsheet: str,
    clave_idempotencia: str,
) -> dict[str, Any]:
    """Prepara una fila para agregarla junto con las demás en un solo Add."""
    return {
        "ID_NOTIFICACION": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_EVENTO": id_evento,
        "ID_VERSION": id_version,
        "ID_APROBACION_ACTUAL": texto(
            aprobador.get("ID_APROBACION_ACTUAL")
        ),
        "DESTINATARIO_EMAIL": obtener_email_notificacion(aprobador),
        "DESTINATARIO_NOMBRE": texto(aprobador.get("NOMBRE")),
        "ORDEN_DESTINATARIO": aprobador.get("ORDEN", ""),
        "ROL_DESTINATARIO": texto(aprobador.get("ROL_FLUJO")),
        "TIPO_NOTIFICACION": tipo_notificacion,
        "ASUNTO": asunto,
        "CUERPO": cuerpo,
        "LINK_DOCUMENTO": link_documento,
        "LINK_APPSHEET": link_appsheet,
        "ESTADO_ENVIO": "Pendiente",
        "INTENTOS": 0,
        "FECHA_CREACION": ahora_iso(),
        "CLAVE_IDEMPOTENCIA": clave_idempotencia,
    }


def notificar_destinatarios_internos_optimizado(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    destinatarios: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Envía las notificaciones internas reduciendo llamadas a AppSheet.

    Optimiza cuatro puntos sin alterar el flujo documental:
    - historial y enlace AppSheet calculados una sola vez;
    - una sola lectura de Documento_Notificaciones;
    - un Add y un Edit por lote para las notificaciones;
    - un solo servicio Gmail y un solo Edit final de Documentos.
    """
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    id_evento = texto(evento.get("ID_EVENTO"))
    id_version = (
        texto(evento.get("ID_VERSION"))
        or texto(documento.get("ID_VERSION_ACTUAL"))
    )

    resultados_por_indice: dict[int, dict[str, Any]] = {}
    especificaciones_validas: list[dict[str, Any]] = []
    correos_procesados: set[str] = set()
    aprobadores_contexto: list[dict[str, Any]] = []

    for indice, especificacion in enumerate(destinatarios):
        aprobador = especificacion.get("aprobador")
        if not isinstance(aprobador, dict):
            resultados_por_indice[indice] = {
                "ok": False,
                "omitida": True,
                "error": "Especificación sin aprobador válido.",
            }
            continue

        aprobadores_contexto.append(aprobador)
        email = obtener_email_notificacion(aprobador)
        tipo_notificacion = texto(
            especificacion.get("tipo_notificacion")
        )

        if not email:
            resultados_por_indice[indice] = {
                "ok": False,
                "omitida": True,
                "destinatario": "",
                "tipo_notificacion": tipo_notificacion,
                "mensaje": (
                    "El responsable no tiene Aprobador_v para enviar "
                    "la notificación."
                ),
            }
            continue

        if not _EMAIL_RE.fullmatch(email):
            resultados_por_indice[indice] = {
                "ok": False,
                "omitida": True,
                "destinatario": email,
                "tipo_notificacion": tipo_notificacion,
                "mensaje": "Correo interno inválido.",
            }
            continue

        if email in correos_procesados:
            continue

        correos_procesados.add(email)
        especificaciones_validas.append(
            {
                "indice": indice,
                "especificacion": especificacion,
                "aprobador": aprobador,
                "email": email,
                "tipo_notificacion": tipo_notificacion,
            }
        )

    if not especificaciones_validas:
        return [
            resultados_por_indice[indice]
            for indice in sorted(resultados_por_indice)
        ]

    try:
        link_appsheet = construir_link_appsheet(id_documento)
        mapa_nombres = construir_mapa_nombres_desde_aprobadores(
            aprobadores_contexto
        )
        historial_texto, historial_html = construir_historial_comentarios(
            id_documento=id_documento,
            id_evento_excluir=id_evento,
            mapa_nombres=mapa_nombres,
        )
        notificaciones_existentes = buscar_notificaciones_documento(
            id_documento
        )
    except Exception as exc:
        traceback.print_exc()
        for item in especificaciones_validas:
            resultados_por_indice[item["indice"]] = {
                "ok": False,
                "omitida": False,
                "destinatario": item["email"],
                "tipo_notificacion": item["tipo_notificacion"],
                "error": str(exc),
            }
        return [
            resultados_por_indice[indice]
            for indice in sorted(resultados_por_indice)
        ]

    notificaciones_por_clave: dict[str, list[dict[str, Any]]] = {}
    for fila in notificaciones_existentes:
        clave = texto(fila.get("CLAVE_IDEMPOTENCIA"))
        if clave:
            notificaciones_por_clave.setdefault(clave, []).append(fila)

    nuevas_filas: list[dict[str, Any]] = []
    preparadas: list[dict[str, Any]] = []

    for item in especificaciones_validas:
        indice = item["indice"]
        especificacion = item["especificacion"]
        aprobador = item["aprobador"]
        email = item["email"]
        tipo_notificacion = item["tipo_notificacion"]

        link_documento = (
            normalizar_url_appsheet(
                especificacion.get("link_documento")
            )
            if tipo_notificacion == "Acción requerida"
            else ""
        )

        try:
            asunto, cuerpo_texto, cuerpo_html = construir_email_notificacion(
                documento=documento,
                destinatario=aprobador,
                tipo_notificacion=tipo_notificacion,
                movimiento=texto(especificacion.get("movimiento")),
                comentario_principal=texto(
                    especificacion.get("comentario_principal")
                ),
                historial_texto=historial_texto,
                historial_html=historial_html,
                link_documento=link_documento,
                link_appsheet=link_appsheet,
            )
            clave = construir_clave_idempotencia(
                id_evento=id_evento,
                email=email,
                tipo_notificacion=tipo_notificacion,
            )
        except Exception as exc:
            resultados_por_indice[indice] = {
                "ok": False,
                "omitida": False,
                "destinatario": email,
                "tipo_notificacion": tipo_notificacion,
                "error": str(exc),
            }
            continue

        coincidencias = notificaciones_por_clave.get(clave, [])
        if len(coincidencias) > 1:
            resultados_por_indice[indice] = {
                "ok": False,
                "omitida": True,
                "destinatario": email,
                "tipo_notificacion": tipo_notificacion,
                "error": (
                    "Existen varias notificaciones con la misma "
                    "CLAVE_IDEMPOTENCIA."
                ),
            }
            continue

        creada = not coincidencias
        if creada:
            notificacion = construir_fila_notificacion_pendiente(
                id_evento=id_evento,
                id_documento=id_documento,
                id_version=id_version,
                aprobador=aprobador,
                tipo_notificacion=tipo_notificacion,
                asunto=asunto,
                cuerpo=cuerpo_texto,
                link_documento=link_documento,
                link_appsheet=link_appsheet,
                clave_idempotencia=clave,
            )
            nuevas_filas.append(notificacion)
            notificaciones_por_clave[clave] = [notificacion]
        else:
            notificacion = coincidencias[0]

        estado_existente = texto(notificacion.get("ESTADO_ENVIO"))
        id_notificacion = texto(
            notificacion.get("ID_NOTIFICACION")
        )
        try:
            intentos = entero(
                notificacion.get("INTENTOS") or 0,
                "INTENTOS",
            )
        except ValueError:
            intentos = 0

        if not creada and estado_existente == "Enviada":
            resultados_por_indice[indice] = {
                "ok": True,
                "omitida": True,
                "duplicada": True,
                "id_notificacion": id_notificacion,
                "destinatario": email,
                "mensaje": "La notificación ya había sido enviada.",
            }
            continue

        if (
            not creada
            and estado_existente
            not in ESTADOS_NOTIFICACION_REINTENTABLES
        ):
            resultados_por_indice[indice] = {
                "ok": False,
                "omitida": True,
                "id_notificacion": id_notificacion,
                "destinatario": email,
                "mensaje": (
                    "La notificación existente no permite reintento. "
                    f"Estado: {estado_existente!r}"
                ),
            }
            continue

        preparadas.append(
            {
                **item,
                "notificacion": notificacion,
                "creada": creada,
                "id_notificacion": id_notificacion,
                "intentos": intentos,
                "asunto": asunto,
                "cuerpo_texto": cuerpo_texto,
                "cuerpo_html": cuerpo_html,
            }
        )

    if nuevas_filas:
        try:
            appsheet_action(
                TABLA_NOTIFICACIONES,
                "Add",
                nuevas_filas,
            )
            notificaciones_existentes.extend(nuevas_filas)
        except Exception as exc:
            traceback.print_exc()
            ids_nuevos = {
                texto(fila.get("ID_NOTIFICACION"))
                for fila in nuevas_filas
            }
            preparadas_restantes: list[dict[str, Any]] = []
            for item in preparadas:
                if item["id_notificacion"] in ids_nuevos:
                    resultados_por_indice[item["indice"]] = {
                        "ok": False,
                        "omitida": False,
                        "destinatario": item["email"],
                        "tipo_notificacion": item[
                            "tipo_notificacion"
                        ],
                        "error": (
                            "No fue posible crear la notificación en "
                            f"AppSheet: {exc}"
                        ),
                    }
                else:
                    preparadas_restantes.append(item)
            preparadas = preparadas_restantes

    if not preparadas:
        return [
            resultados_por_indice[indice]
            for indice in sorted(resultados_por_indice)
        ]

    filas_estado: list[dict[str, Any]] = []
    ultimo_envio_exitoso: dict[str, str] | None = None

    try:
        gmail_service = obtener_gmail_service()
    except Exception as exc:
        traceback.print_exc()
        gmail_service = None
        for item in preparadas:
            filas_estado.append(
                {
                    "ID_NOTIFICACION": item["id_notificacion"],
                    "ESTADO_ENVIO": "Error",
                    "INTENTOS": item["intentos"] + 1,
                    "ERROR_ENVIO": texto(str(exc))[:1500],
                }
            )
            resultados_por_indice[item["indice"]] = {
                "ok": False,
                "omitida": False,
                "destinatario": item["email"],
                "tipo_notificacion": item["tipo_notificacion"],
                "error": str(exc),
            }

    if gmail_service is not None:
        for item in preparadas:
            id_notificacion = item["id_notificacion"]
            email = item["email"]
            tipo_notificacion = item["tipo_notificacion"]
            intentos = item["intentos"]

            notificacion_anterior = (
                seleccionar_ultima_notificacion_hilo_en_memoria(
                    filas=notificaciones_existentes,
                    destinatario_email=email,
                    asunto=item["asunto"],
                    id_notificacion_excluir=id_notificacion,
                )
            )

            thread_id_anterior = ""
            in_reply_to = ""
            if notificacion_anterior:
                thread_id_anterior = texto(
                    notificacion_anterior.get("GMAIL_THREAD_ID")
                )
                id_notificacion_anterior = texto(
                    notificacion_anterior.get("ID_NOTIFICACION")
                )
                if id_notificacion_anterior:
                    in_reply_to = (
                        construir_rfc_message_id_notificacion(
                            id_notificacion_anterior
                        )
                    )

            try:
                respuesta_gmail = enviar_email_notificacion(
                    gmail_service=gmail_service,
                    destinatario=email,
                    asunto=item["asunto"],
                    cuerpo_texto=item["cuerpo_texto"],
                    cuerpo_html=item["cuerpo_html"],
                    rfc_message_id=(
                        construir_rfc_message_id_notificacion(
                            id_notificacion
                        )
                    ),
                    thread_id=thread_id_anterior,
                    in_reply_to=in_reply_to,
                )
                fecha_envio = ahora_iso()
                fila_estado = {
                    "ID_NOTIFICACION": id_notificacion,
                    "ESTADO_ENVIO": "Enviada",
                    "INTENTOS": intentos + 1,
                    "FECHA_ENVIO": fecha_envio,
                    "ERROR_ENVIO": "",
                    "GMAIL_MESSAGE_ID": respuesta_gmail[
                        "message_id"
                    ],
                    "GMAIL_THREAD_ID": respuesta_gmail[
                        "thread_id"
                    ],
                }
                filas_estado.append(fila_estado)
                item["notificacion"].update(fila_estado)

                ultimo_envio_exitoso = {
                    "message_id": respuesta_gmail["message_id"],
                    "thread_id": respuesta_gmail["thread_id"],
                }
                resultados_por_indice[item["indice"]] = {
                    "ok": True,
                    "omitida": False,
                    "id_notificacion": id_notificacion,
                    "destinatario": email,
                    "tipo_notificacion": tipo_notificacion,
                    "message_id": respuesta_gmail["message_id"],
                    "thread_id": respuesta_gmail["thread_id"],
                }
            except Exception as exc:
                traceback.print_exc()
                filas_estado.append(
                    {
                        "ID_NOTIFICACION": id_notificacion,
                        "ESTADO_ENVIO": "Error",
                        "INTENTOS": intentos + 1,
                        "ERROR_ENVIO": texto(str(exc))[:1500],
                    }
                )
                resultados_por_indice[item["indice"]] = {
                    "ok": False,
                    "omitida": False,
                    "destinatario": email,
                    "tipo_notificacion": tipo_notificacion,
                    "error": str(exc),
                }

    if filas_estado:
        try:
            appsheet_action(
                TABLA_NOTIFICACIONES,
                "Edit",
                filas_estado,
            )
        except Exception as exc:
            traceback.print_exc()
            for resultado in resultados_por_indice.values():
                if resultado.get("id_notificacion"):
                    resultado["advertencia_registro"] = (
                        "El correo fue procesado, pero no se pudo actualizar "
                        "Documento_Notificaciones: " + str(exc)
                    )

    if ultimo_envio_exitoso:
        try:
            actualizar_resumen_notificacion_documento(
                id_documento=id_documento,
                message_id=ultimo_envio_exitoso["message_id"],
                thread_id=ultimo_envio_exitoso["thread_id"],
            )
        except Exception:
            traceback.print_exc()

    return [
        resultados_por_indice[indice]
        for indice in sorted(resultados_por_indice)
    ]


def notificar_destinatarios_internos(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    destinatarios: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Selecciona la implementación optimizada o la anterior por variable."""
    usar_lote = os.getenv(
        "USE_BATCH_NOTIFICATIONS",
        "true",
    ).strip().lower() not in {"0", "false", "no", "off"}

    if not usar_lote:
        return notificar_destinatarios_internos_legacy(
            documento=documento,
            evento=evento,
            destinatarios=destinatarios,
        )

    return notificar_destinatarios_internos_optimizado(
        documento=documento,
        evento=evento,
        destinatarios=destinatarios,
    )


def buscar_evento_envio_revision_actual(
    *,
    id_documento: str,
    id_version: str,
    id_aprobacion_actual: str,
) -> dict[str, Any] | None:
    """
    Recupera el evento exacto que originó la revisión vigente.

    Se usa en reintentos del Bot para completar únicamente las
    notificaciones pendientes o con error.
    """
    candidatos = [
        evento
        for evento in buscar_eventos_documento(id_documento)
        if texto(evento.get("TIPO_EVENTO")) == "Enviado a revisión"
        and texto(evento.get("ID_VERSION")) == id_version
        and texto(evento.get("ID_APROBACION_ACTUAL"))
        == id_aprobacion_actual
    ]

    if not candidatos:
        return None

    candidatos.sort(
        key=lambda evento: (
            parsear_fecha_appsheet(evento.get("FECHA_EVENTO")),
            texto(evento.get("ID_EVENTO")),
        )
    )
    return candidatos[-1]


def construir_especificaciones_envio_revision(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    cadena: list[dict[str, Any]],
    aprobador_anterior: dict[str, Any],
    aprobador_actual: dict[str, Any],
) -> list[dict[str, Any]]:
    """
    Define la prioridad de destinatarios para un envío a revisión.

    El orden es deliberado: Acción requerida, Confirmación e Informativa.
    Si dos filas comparten Aprobador_v, la notificación de mayor prioridad
    gana porque notificar_destinatarios_internos deduplica en ese orden.
    """
    comentario_evento = (
        texto(evento.get("COMENTARIO"))
        or "El documento fue enviado a revisión."
    )
    link_documento = normalizar_url_appsheet(
        documento.get("GOOGLE_DOC_URL")
    )

    especificaciones: list[dict[str, Any]] = [
        {
            "aprobador": aprobador_actual,
            "tipo_notificacion": "Acción requerida",
            "movimiento": "Documento asignado para revisión",
            "comentario_principal": comentario_evento,
            "link_documento": link_documento,
        },
        {
            "aprobador": aprobador_anterior,
            "tipo_notificacion": "Confirmación",
            "movimiento": "Documento enviado a revisión",
            "comentario_principal": comentario_evento,
            "link_documento": "",
        },
    ]

    for integrante in cadena:
        especificaciones.append(
            {
                "aprobador": integrante,
                "tipo_notificacion": "Informativa",
                "movimiento": "El documento avanzó a revisión",
                "comentario_principal": comentario_evento,
                "link_documento": "",
            }
        )

    return especificaciones


def ejecutar_notificaciones_envio_revision(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    cadena: list[dict[str, Any]],
    aprobador_anterior: dict[str, Any],
    aprobador_actual: dict[str, Any],
) -> list[dict[str, Any]]:
    """Envía y registra las notificaciones internas del primer avance."""
    especificaciones = construir_especificaciones_envio_revision(
        documento=documento,
        evento=evento,
        cadena=cadena,
        aprobador_anterior=aprobador_anterior,
        aprobador_actual=aprobador_actual,
    )
    return notificar_destinatarios_internos(
        documento=documento,
        evento=evento,
        destinatarios=especificaciones,
    )


def reanudar_notificaciones_envio_revision(
    *,
    documento: dict[str, Any],
    datos_solicitud: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[str]]:
    """
    Completa notificaciones de un envío a revisión ya ejecutado.

    No repite correos enviados. Si el evento faltó por una falla parcial, lo
    reconstruye desde la versión vigente y continúa de forma idempotente.
    """
    advertencias: list[str] = []
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    id_version = texto(documento.get("ID_VERSION_ACTUAL"))
    id_aprobacion_actual = texto(
        documento.get("ID_APROBACION_ACTUAL")
    )
    numero_version = entero(
        documento.get("VERSION_ACTUAL"),
        "VERSION_ACTUAL",
    )

    if not id_version or not id_aprobacion_actual:
        advertencias.append(
            "No fue posible identificar la versión o el encargado actual "
            "para reanudar las notificaciones."
        )
        return [], advertencias

    version = buscar_version_por_id(id_version)
    if texto(version.get("MOTIVO_CREACION")) != "Envío a revisión":
        advertencias.append(
            "El documento está En revisión, pero la versión vigente no "
            "corresponde al envío inicial a revisión. No se ejecutaron "
            "notificaciones desde /enviar-revision."
        )
        return [], advertencias

    aprobador_actual = buscar_aprobacion_actual(
        id_aprobacion_actual
    )
    cadena = buscar_cadena_actual_documento(
        id_documento=id_documento,
        numero_version=numero_version,
    )
    orden_actual = entero(
        aprobador_actual.get("ORDEN"),
        "ORDEN",
    )

    anteriores = [
        fila
        for fila in cadena
        if entero(fila.get("ORDEN"), "ORDEN") < orden_actual
    ]
    enviados = [
        fila
        for fila in anteriores
        if texto(fila.get("RESULTADO")) == "Enviado"
    ]
    candidatos_anterior = enviados or anteriores
    if not candidatos_anterior:
        advertencias.append(
            "No se encontró al responsable anterior para reconstruir las "
            "notificaciones del envío a revisión."
        )
        return [], advertencias

    aprobador_anterior = max(
        candidatos_anterior,
        key=lambda fila: entero(fila.get("ORDEN"), "ORDEN"),
    )

    evento = buscar_evento_envio_revision_actual(
        id_documento=id_documento,
        id_version=id_version,
        id_aprobacion_actual=id_aprobacion_actual,
    )

    if evento is None:
        evento = crear_evento_envio_revision(
            id_documento=id_documento,
            id_version=id_version,
            id_aprobacion_actual=id_aprobacion_actual,
            usuario=(
                texto(documento.get("ULTIMO_ENVIADO_POR"))
                or texto(datos_solicitud.get("usuario"))
                or texto(aprobador_anterior.get("APROBADOR"))
            ),
            fecha=(
                texto(documento.get("FECHA_ULTIMO_ENVIO"))
                or texto(version.get("FECHA_CREACION"))
                or ahora_iso()
            ),
            comentario=(
                texto(version.get("COMENTARIO_CAMBIO"))
                or texto(datos_solicitud.get("comentario"))
            ),
            nombre_archivo=(
                texto(version.get("NOMBRE_ARCHIVO"))
                or texto(documento.get("TITULO"))
                or id_documento
            ),
        )
        advertencias.append(
            "El evento de envío a revisión faltaba y fue reconstruido "
            "antes de reanudar las notificaciones."
        )

    resultados = ejecutar_notificaciones_envio_revision(
        documento=documento,
        evento=evento,
        cadena=cadena,
        aprobador_anterior=aprobador_anterior,
        aprobador_actual=aprobador_actual,
    )

    fallidas = [
        resultado
        for resultado in resultados
        if not resultado.get("ok")
    ]
    if fallidas:
        advertencias.append(
            f"{len(fallidas)} notificación(es) siguen omitidas o con "
            "error. Revisa Documento_Notificaciones."
        )

    return resultados, advertencias



def buscar_evento_revision_aprobada_actual(
    *,
    id_documento: str,
    id_version: str,
    id_aprobacion_evento: str,
) -> dict[str, Any] | None:
    """Recupera el evento exacto de la aprobación ya procesada."""
    candidatos = [
        evento
        for evento in buscar_eventos_documento(id_documento)
        if texto(evento.get("TIPO_EVENTO")) == "Revisión aprobada"
        and texto(evento.get("ID_VERSION")) == id_version
        and texto(evento.get("ID_APROBACION_ACTUAL"))
        == id_aprobacion_evento
    ]

    if not candidatos:
        return None

    candidatos.sort(
        key=lambda evento: (
            parsear_fecha_appsheet(evento.get("FECHA_EVENTO")),
            texto(evento.get("ID_EVENTO")),
        )
    )
    return candidatos[-1]


def construir_especificaciones_aprobacion_revision(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    cadena: list[dict[str, Any]],
    aprobador_aprueba: dict[str, Any],
    aprobador_actual: dict[str, Any] | None,
    ultimo_aprobador: bool,
) -> list[dict[str, Any]]:
    """
    Define los destinatarios de una aprobación.

    Aprobación intermedia:
    - siguiente responsable: Acción requerida;
    - quien aprobó: Confirmación;
    - resto de la cadena: Informativa.

    Aprobación final:
    - quien aprobó: Confirmación;
    - resto de la cadena: Informativa.

    No se asigna Acción requerida al quedar Listo para firma porque el
    documento ya no mantiene un ID_APROBACION_ACTUAL interno.
    """
    comentario_evento = (
        texto(evento.get("COMENTARIO"))
        or "La revisión fue aprobada."
    )

    especificaciones: list[dict[str, Any]] = []

    if not ultimo_aprobador and aprobador_actual is not None:
        especificaciones.append(
            {
                "aprobador": aprobador_actual,
                "tipo_notificacion": "Acción requerida",
                "movimiento": "Documento asignado para continuar la revisión",
                "comentario_principal": comentario_evento,
                "link_documento": normalizar_url_appsheet(
                    documento.get("GOOGLE_DOC_URL")
                ),
            }
        )

    especificaciones.append(
        {
            "aprobador": aprobador_aprueba,
            "tipo_notificacion": "Confirmación",
            "movimiento": (
                "Revisión final aprobada"
                if ultimo_aprobador
                else "Revisión aprobada y enviada al siguiente responsable"
            ),
            "comentario_principal": comentario_evento,
            "link_documento": "",
        }
    )

    for integrante in cadena:
        especificaciones.append(
            {
                "aprobador": integrante,
                "tipo_notificacion": "Informativa",
                "movimiento": (
                    "Documento listo para firma"
                    if ultimo_aprobador
                    else "El documento avanzó al siguiente responsable"
                ),
                "comentario_principal": comentario_evento,
                "link_documento": "",
            }
        )

    return especificaciones


def ejecutar_notificaciones_aprobacion_revision(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    cadena: list[dict[str, Any]],
    aprobador_aprueba: dict[str, Any],
    aprobador_actual: dict[str, Any] | None,
    ultimo_aprobador: bool,
) -> list[dict[str, Any]]:
    especificaciones = construir_especificaciones_aprobacion_revision(
        documento=documento,
        evento=evento,
        cadena=cadena,
        aprobador_aprueba=aprobador_aprueba,
        aprobador_actual=aprobador_actual,
        ultimo_aprobador=ultimo_aprobador,
    )
    return notificar_destinatarios_internos(
        documento=documento,
        evento=evento,
        destinatarios=especificaciones,
    )


def reanudar_notificaciones_aprobacion_revision(
    *,
    documento: dict[str, Any],
    id_aprobacion_aprueba: str,
    datos_solicitud: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[str]]:
    """Completa notificaciones de una aprobación ya procesada."""
    advertencias: list[str] = []
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    id_version = texto(documento.get("ID_VERSION_ACTUAL"))
    numero_version = entero(
        documento.get("VERSION_ACTUAL"),
        "VERSION_ACTUAL",
    )
    ultimo_aprobador = texto(documento.get("ESTADO")) == "Listo para firma"

    cadena = buscar_cadena_documento_version(
        id_documento=id_documento,
        numero_version=numero_version,
    )
    if not cadena:
        advertencias.append(
            "No se encontró la cadena de aprobación para reanudar las "
            "notificaciones."
        )
        return [], advertencias

    aprobador_actual: dict[str, Any] | None = None
    if not ultimo_aprobador:
        id_aprobacion_actual = texto(
            documento.get("ID_APROBACION_ACTUAL")
        )
        if not id_aprobacion_actual:
            advertencias.append(
                "El documento sigue En revisión, pero no tiene encargado "
                "actual para reanudar las notificaciones."
            )
            return [], advertencias
        aprobador_actual = buscar_aprobacion_actual(
            id_aprobacion_actual
        )

    aprobador_aprueba: dict[str, Any] | None = None
    if id_aprobacion_aprueba:
        aprobador_aprueba = buscar_aprobacion_actual(
            id_aprobacion_aprueba
        )
    else:
        aprobados = [
            fila
            for fila in cadena
            if texto(fila.get("RESULTADO")) == "Aprobado"
        ]
        if aprobador_actual is not None:
            orden_actual = entero(
                aprobador_actual.get("ORDEN"),
                "ORDEN",
            )
            aprobados_anteriores = [
                fila
                for fila in aprobados
                if entero(fila.get("ORDEN"), "ORDEN") < orden_actual
            ]
            aprobados = aprobados_anteriores or aprobados

        if aprobados:
            aprobador_aprueba = max(
                aprobados,
                key=lambda fila: entero(fila.get("ORDEN"), "ORDEN"),
            )

    if aprobador_aprueba is None:
        advertencias.append(
            "No se pudo identificar al responsable que aprobó la revisión."
        )
        return [], advertencias

    orden_aprueba = entero(
        aprobador_aprueba.get("ORDEN"),
        "ORDEN",
    )
    orden_siguiente = (
        None
        if ultimo_aprobador or aprobador_actual is None
        else entero(aprobador_actual.get("ORDEN"), "ORDEN")
    )
    id_aprobacion_evento = (
        texto(aprobador_aprueba.get("ID_APROBACION_ACTUAL"))
        if ultimo_aprobador
        else texto(aprobador_actual.get("ID_APROBACION_ACTUAL"))
    )

    evento = buscar_evento_revision_aprobada_actual(
        id_documento=id_documento,
        id_version=id_version,
        id_aprobacion_evento=id_aprobacion_evento,
    )

    if evento is None:
        version = buscar_version_por_id(id_version)
        evento = crear_evento_revision_aprobada(
            id_documento=id_documento,
            id_version=id_version,
            id_aprobacion_actual=id_aprobacion_evento,
            usuario=(
                texto(documento.get("ULTIMO_ENVIADO_POR"))
                or texto(datos_solicitud.get("usuario"))
                or texto(aprobador_aprueba.get("APROBADOR"))
            ),
            fecha=(
                texto(documento.get("FECHA_ULTIMO_ENVIO"))
                or texto(version.get("FECHA_CREACION"))
                or ahora_iso()
            ),
            comentario=(
                texto(aprobador_aprueba.get("COMENTARIO"))
                or texto(version.get("COMENTARIO_CAMBIO"))
                or texto(datos_solicitud.get("comentario"))
            ),
            orden_actual=orden_aprueba,
            orden_siguiente=orden_siguiente,
        )
        advertencias.append(
            "El evento de aprobación faltaba y fue reconstruido antes de "
            "reanudar las notificaciones."
        )

    resultados = ejecutar_notificaciones_aprobacion_revision(
        documento=documento,
        evento=evento,
        cadena=cadena,
        aprobador_aprueba=aprobador_aprueba,
        aprobador_actual=aprobador_actual,
        ultimo_aprobador=ultimo_aprobador,
    )

    fallidas = [
        resultado
        for resultado in resultados
        if not resultado.get("ok")
    ]
    if fallidas:
        advertencias.append(
            f"{len(fallidas)} notificación(es) siguen omitidas o con "
            "error. Revisa Documento_Notificaciones."
        )

    return resultados, advertencias



def buscar_integrante_cadena_por_usuario(
    cadena: list[dict[str, Any]],
    usuario: str,
) -> dict[str, Any] | None:
    """Busca un integrante usando su correo operativo o de notificación."""
    correo = texto(usuario).strip().lower()
    if not correo:
        return None

    for integrante in cadena:
        correos = {
            texto(integrante.get("APROBADOR")).strip().lower(),
            texto(integrante.get("Aprobador_v")).strip().lower(),
        }
        correos.discard("")
        if correo in correos:
            return integrante

    return None


def buscar_evento_rechazo_revision_actual(
    *,
    id_documento: str,
    id_aprobacion_rechaza: str,
) -> dict[str, Any] | None:
    candidatos = [
        evento
        for evento in buscar_eventos_documento(id_documento)
        if texto(evento.get("TIPO_EVENTO")) == "Revisión rechazada"
        and texto(evento.get("ID_APROBACION_ACTUAL"))
        == id_aprobacion_rechaza
    ]
    if not candidatos:
        return None

    candidatos.sort(
        key=lambda evento: (
            parsear_fecha_appsheet(evento.get("FECHA_EVENTO")),
            texto(evento.get("ID_EVENTO")),
        )
    )
    return candidatos[-1]


def crear_evento_rechazo_revision_reconstruido(
    *,
    documento: dict[str, Any],
    aprobador_rechaza: dict[str, Any],
    comentario: str,
    usuario: str,
) -> dict[str, Any]:
    evento = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": texto(documento.get("ID_DOCUMENTO")),
        "ID_VERSION": texto(
            aprobador_rechaza.get("ID_VERSION_TRABAJADA")
        ),
        "ID_APROBACION_ACTUAL": texto(
            aprobador_rechaza.get("ID_APROBACION_ACTUAL")
        ),
        "TIPO_EVENTO": "Revisión rechazada",
        "ESTADO_ANTERIOR": (
            "En revisión - Orden "
            + texto(aprobador_rechaza.get("ORDEN"))
        ),
        "ESTADO_NUEVO": texto(documento.get("ESTADO")),
        "USUARIO": (
            texto(usuario)
            or texto(aprobador_rechaza.get("APROBADOR"))
        ),
        "FECHA_EVENTO": (
            texto(aprobador_rechaza.get("FECHA_RESPUESTA"))
            or texto(documento.get("FECHA_ULTIMO_ENVIO"))
            or ahora_iso()
        ),
        "COMENTARIO": (
            texto(comentario)
            or texto(aprobador_rechaza.get("COMENTARIO"))
            or "La revisión fue rechazada."
        ),
    }
    appsheet_action(TABLA_EVENTOS, "Add", [evento])
    return evento


def construir_especificaciones_rechazo_revision(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    cadena: list[dict[str, Any]],
    aprobador_rechaza: dict[str, Any],
    aprobador_destino: dict[str, Any],
) -> list[dict[str, Any]]:
    comentario_evento = (
        texto(evento.get("COMENTARIO"))
        or "La revisión fue rechazada y el documento volvió al responsable anterior."
    )

    especificaciones: list[dict[str, Any]] = [
        {
            "aprobador": aprobador_destino,
            "tipo_notificacion": "Acción requerida",
            "movimiento": "Documento devuelto para corrección",
            "comentario_principal": comentario_evento,
            "link_documento": normalizar_url_appsheet(
                documento.get("GOOGLE_DOC_URL")
            ),
        },
        {
            "aprobador": aprobador_rechaza,
            "tipo_notificacion": "Confirmación",
            "movimiento": "Rechazo registrado y documento devuelto",
            "comentario_principal": comentario_evento,
            "link_documento": "",
        },
    ]

    for integrante in cadena:
        especificaciones.append(
            {
                "aprobador": integrante,
                "tipo_notificacion": "Informativa",
                "movimiento": "Documento devuelto por observaciones",
                "comentario_principal": comentario_evento,
                "link_documento": "",
            }
        )

    return especificaciones


def ejecutar_notificaciones_rechazo_revision(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    cadena: list[dict[str, Any]],
    aprobador_rechaza: dict[str, Any],
    aprobador_destino: dict[str, Any],
) -> list[dict[str, Any]]:
    return notificar_destinatarios_internos(
        documento=documento,
        evento=evento,
        destinatarios=construir_especificaciones_rechazo_revision(
            documento=documento,
            evento=evento,
            cadena=cadena,
            aprobador_rechaza=aprobador_rechaza,
            aprobador_destino=aprobador_destino,
        ),
    )


def reanudar_notificaciones_rechazo_revision(
    *,
    documento: dict[str, Any],
    id_aprobacion_rechaza: str,
    datos_solicitud: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[str]]:
    advertencias: list[str] = []
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    id_aprobacion_destino = texto(
        documento.get("ID_APROBACION_ACTUAL")
    )
    numero_version = entero(
        documento.get("VERSION_ACTUAL"),
        "VERSION_ACTUAL",
    )

    if not id_aprobacion_rechaza or not id_aprobacion_destino:
        return [], [
            "No fue posible identificar al responsable que rechazó o al "
            "responsable que recibió el documento."
        ]

    aprobador_rechaza = buscar_aprobacion_actual(
        id_aprobacion_rechaza
    )
    aprobador_destino = buscar_aprobacion_actual(
        id_aprobacion_destino
    )
    cadena = buscar_cadena_documento_version(
        id_documento=id_documento,
        numero_version=numero_version,
    )

    evento = buscar_evento_rechazo_revision_actual(
        id_documento=id_documento,
        id_aprobacion_rechaza=id_aprobacion_rechaza,
    )
    if evento is None:
        evento = crear_evento_rechazo_revision_reconstruido(
            documento=documento,
            aprobador_rechaza=aprobador_rechaza,
            comentario=texto(datos_solicitud.get("comentario")),
            usuario=texto(datos_solicitud.get("usuario")),
        )
        advertencias.append(
            "El evento de rechazo faltaba y fue reconstruido antes de "
            "reanudar las notificaciones."
        )

    resultados = ejecutar_notificaciones_rechazo_revision(
        documento=documento,
        evento=evento,
        cadena=cadena,
        aprobador_rechaza=aprobador_rechaza,
        aprobador_destino=aprobador_destino,
    )
    fallidas = [r for r in resultados if not r.get("ok")]
    if fallidas:
        advertencias.append(
            f"{len(fallidas)} notificación(es) siguen omitidas o con error. "
            "Revisa Documento_Notificaciones."
        )
    return resultados, advertencias


def buscar_evento_reinicio_firma_actual(
    *,
    id_documento: str,
    id_version_observada: str,
) -> dict[str, Any] | None:
    candidatos = [
        evento
        for evento in buscar_eventos_documento(id_documento)
        if texto(evento.get("TIPO_EVENTO")) == "Proceso reiniciado"
        and (
            not id_version_observada
            or texto(evento.get("ID_VERSION")) == id_version_observada
        )
    ]
    if not candidatos:
        return None
    candidatos.sort(
        key=lambda evento: (
            parsear_fecha_appsheet(evento.get("FECHA_EVENTO")),
            texto(evento.get("ID_EVENTO")),
        )
    )
    return candidatos[-1]


def crear_evento_reinicio_firma_reconstruido(
    *,
    documento: dict[str, Any],
    id_version_observada: str,
    usuario: str,
    comentario: str,
) -> dict[str, Any]:
    evento = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": texto(documento.get("ID_DOCUMENTO")),
        "ID_VERSION": id_version_observada,
        "ID_APROBACION_ACTUAL": "",
        "TIPO_EVENTO": "Proceso reiniciado",
        "ESTADO_ANTERIOR": "En firma",
        "ESTADO_NUEVO": "Borrador",
        "USUARIO": texto(usuario),
        "FECHA_EVENTO": (
            texto(documento.get("FECHA_ULTIMO_ENVIO")) or ahora_iso()
        ),
        "COMENTARIO": (
            "El documento recibió observaciones durante la firma. Motivo: "
            + (texto(comentario) or "Sin detalle adicional.")
        ),
    }
    appsheet_action(TABLA_EVENTOS, "Add", [evento])
    return evento


def construir_especificaciones_reinicio_firma(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    cadena: list[dict[str, Any]],
    aprobador_destino: dict[str, Any],
    aprobador_confirma: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    comentario_evento = (
        texto(evento.get("COMENTARIO"))
        or "El documento recibió observaciones durante la firma."
    )
    especificaciones: list[dict[str, Any]] = [
        {
            "aprobador": aprobador_destino,
            "tipo_notificacion": "Acción requerida",
            "movimiento": "Proceso reiniciado por observaciones de firma",
            "comentario_principal": comentario_evento,
            "link_documento": normalizar_url_appsheet(
                documento.get("GOOGLE_DOC_URL")
            ),
        }
    ]

    if aprobador_confirma is not None:
        especificaciones.append(
            {
                "aprobador": aprobador_confirma,
                "tipo_notificacion": "Confirmación",
                "movimiento": "Observaciones de firma registradas",
                "comentario_principal": comentario_evento,
                "link_documento": "",
            }
        )

    for integrante in cadena:
        especificaciones.append(
            {
                "aprobador": integrante,
                "tipo_notificacion": "Informativa",
                "movimiento": "El proceso volvió a borrador",
                "comentario_principal": comentario_evento,
                "link_documento": "",
            }
        )

    return especificaciones


def ejecutar_notificaciones_reinicio_firma(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    cadena: list[dict[str, Any]],
    aprobador_destino: dict[str, Any],
    aprobador_confirma: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    return notificar_destinatarios_internos(
        documento=documento,
        evento=evento,
        destinatarios=construir_especificaciones_reinicio_firma(
            documento=documento,
            evento=evento,
            cadena=cadena,
            aprobador_destino=aprobador_destino,
            aprobador_confirma=aprobador_confirma,
        ),
    )


def reanudar_notificaciones_reinicio_firma(
    *,
    documento: dict[str, Any],
    id_version_observada: str,
    datos_solicitud: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[str]]:
    advertencias: list[str] = []
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    numero_version = entero(
        documento.get("VERSION_ACTUAL"),
        "VERSION_ACTUAL",
    )
    id_aprobacion_destino = texto(
        documento.get("ID_APROBACION_ACTUAL")
    )
    if not id_aprobacion_destino:
        return [], [
            "No fue posible identificar al responsable del nuevo borrador."
        ]

    cadena = buscar_cadena_documento_version(
        id_documento=id_documento,
        numero_version=numero_version,
    )
    aprobador_destino = buscar_aprobacion_actual(
        id_aprobacion_destino
    )
    usuario = (
        texto(datos_solicitud.get("usuario"))
        or texto(documento.get("ULTIMO_ENVIADO_POR"))
    )
    aprobador_confirma = buscar_integrante_cadena_por_usuario(
        cadena,
        usuario,
    )

    evento = buscar_evento_reinicio_firma_actual(
        id_documento=id_documento,
        id_version_observada=id_version_observada,
    )
    if evento is None:
        evento = crear_evento_reinicio_firma_reconstruido(
            documento=documento,
            id_version_observada=id_version_observada,
            usuario=usuario,
            comentario=texto(datos_solicitud.get("comentario")),
        )
        advertencias.append(
            "El evento de reinicio faltaba y fue reconstruido antes de "
            "reanudar las notificaciones."
        )

    resultados = ejecutar_notificaciones_reinicio_firma(
        documento=documento,
        evento=evento,
        cadena=cadena,
        aprobador_destino=aprobador_destino,
        aprobador_confirma=aprobador_confirma,
    )
    fallidas = [r for r in resultados if not r.get("ok")]
    if fallidas:
        advertencias.append(
            f"{len(fallidas)} notificación(es) siguen omitidas o con error. "
            "Revisa Documento_Notificaciones."
        )
    return resultados, advertencias


def buscar_evento_cierre_proceso(
    *,
    id_documento: str,
    id_version: str,
) -> dict[str, Any] | None:
    candidatos = [
        evento
        for evento in buscar_eventos_documento(id_documento)
        if texto(evento.get("TIPO_EVENTO")) == "Proceso terminado"
        and (
            not id_version
            or texto(evento.get("ID_VERSION")) == id_version
        )
    ]
    if not candidatos:
        return None
    candidatos.sort(
        key=lambda evento: (
            parsear_fecha_appsheet(evento.get("FECHA_EVENTO")),
            texto(evento.get("ID_EVENTO")),
        )
    )
    return candidatos[-1]


def crear_evento_cierre_reconstruido(
    *,
    documento: dict[str, Any],
    usuario: str,
) -> dict[str, Any]:
    evento = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": texto(documento.get("ID_DOCUMENTO")),
        "ID_VERSION": texto(documento.get("ID_VERSION_ACTUAL")),
        "ID_APROBACION_ACTUAL": "",
        "TIPO_EVENTO": "Proceso terminado",
        "ESTADO_ANTERIOR": "En firma",
        "ESTADO_NUEVO": "Proceso terminado",
        "USUARIO": (
            texto(usuario)
            or texto(documento.get("CARGADO_POR"))
        ),
        "FECHA_EVENTO": (
            texto(documento.get("FECHA_CIERRE")) or ahora_iso()
        ),
        "COMENTARIO": (
            "El documento quedó cerrado con su PDF firmado definitivo."
        ),
    }
    appsheet_action(TABLA_EVENTOS, "Add", [evento])
    return evento


def construir_especificaciones_cierre_proceso(
    *,
    evento: dict[str, Any],
    cadena: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    comentario_evento = (
        texto(evento.get("COMENTARIO"))
        or "El proceso documental terminó correctamente."
    )
    return [
        {
            "aprobador": integrante,
            "tipo_notificacion": "Cierre",
            "movimiento": "Proceso documental finalizado",
            "comentario_principal": comentario_evento,
            "link_documento": "",
        }
        for integrante in cadena
    ]


def ejecutar_notificaciones_cierre_proceso(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    cadena: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    return notificar_destinatarios_internos(
        documento=documento,
        evento=evento,
        destinatarios=construir_especificaciones_cierre_proceso(
            evento=evento,
            cadena=cadena,
        ),
    )


def reanudar_notificaciones_cierre_proceso(
    *,
    documento: dict[str, Any],
    datos_solicitud: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[str]]:
    advertencias: list[str] = []
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    id_version = texto(documento.get("ID_VERSION_ACTUAL"))
    numero_version = entero(
        documento.get("VERSION_ACTUAL"),
        "VERSION_ACTUAL",
    )
    cadena = buscar_cadena_documento_version(
        id_documento=id_documento,
        numero_version=numero_version,
    )
    if not cadena:
        return [], [
            "No se encontró la cadena interna para enviar el cierre."
        ]

    evento = buscar_evento_cierre_proceso(
        id_documento=id_documento,
        id_version=id_version,
    )
    if evento is None:
        evento = crear_evento_cierre_reconstruido(
            documento=documento,
            usuario=texto(datos_solicitud.get("usuario")),
        )
        advertencias.append(
            "El evento de cierre faltaba y fue reconstruido antes de "
            "reanudar las notificaciones."
        )

    resultados = ejecutar_notificaciones_cierre_proceso(
        documento=documento,
        evento=evento,
        cadena=cadena,
    )
    fallidas = [r for r in resultados if not r.get("ok")]
    if fallidas:
        advertencias.append(
            f"{len(fallidas)} notificación(es) siguen omitidas o con error. "
            "Revisa Documento_Notificaciones."
        )
    return resultados, advertencias


def construir_evento_diagnostico(
    documento: dict[str, Any],
) -> dict[str, Any]:
    """Crea un evento virtual cuando el documento aún no tiene bitácora."""
    return {
        "ID_EVENTO": f"diagnostico-{texto(documento.get('ID_DOCUMENTO'))}",
        "ID_DOCUMENTO": texto(documento.get("ID_DOCUMENTO")),
        "ID_VERSION": texto(documento.get("ID_VERSION_ACTUAL")),
        "ID_APROBACION_ACTUAL": texto(
            documento.get("ID_APROBACION_ACTUAL")
        ),
        "TIPO_EVENTO": "Diagnóstico de notificación",
        "ESTADO_ANTERIOR": "",
        "ESTADO_NUEVO": texto(documento.get("ESTADO")),
        "USUARIO": "Sistema",
        "FECHA_EVENTO": ahora_iso(),
        "COMENTARIO": (
            "Vista previa técnica. No se creó ni envió una notificación."
        ),
    }


def seleccionar_evento_diagnostico(
    *,
    id_documento: str,
    id_evento: str,
    documento: dict[str, Any],
) -> dict[str, Any]:
    if id_evento:
        evento = buscar_evento_por_id(id_evento)
        if texto(evento.get("ID_DOCUMENTO")) != id_documento:
            raise ValueError(
                "El evento indicado no pertenece al documento"
            )
        return evento

    eventos = buscar_eventos_documento(id_documento)
    if not eventos:
        return construir_evento_diagnostico(documento)

    eventos.sort(
        key=lambda fila: parsear_fecha_appsheet(
            fila.get("FECHA_EVENTO")
        )
    )
    return eventos[-1]


@app.route("/diagnostico-notificacion", methods=["POST"])
def diagnostico_notificacion():
    """
    Construye una vista previa. No crea filas y no envía correos.
    """
    try:
        validar_configuracion()
        validar_token()

        data = request.get_json(silent=True) or {}
        id_documento = texto(data.get("id_documento"))
        id_aprobacion_actual = texto(
            data.get("id_aprobacion_actual")
        )
        id_evento = texto(data.get("id_evento"))
        tipo_notificacion = (
            texto(data.get("tipo_notificacion"))
            or "Acción requerida"
        )

        if not id_documento:
            return {"error": "Falta id_documento"}, 400
        if tipo_notificacion not in TIPOS_NOTIFICACION_VALIDOS:
            return {
                "error": (
                    "tipo_notificacion debe ser Acción requerida, "
                    "Informativa, Confirmación o Cierre"
                )
            }, 400

        documento = buscar_documento(id_documento)

        id_aprobacion_actual = (
            id_aprobacion_actual
            or texto(documento.get("ID_APROBACION_ACTUAL"))
        )
        if not id_aprobacion_actual:
            raise ValueError(
                "No se pudo identificar un encargado para el diagnóstico"
            )

        aprobador = buscar_aprobacion_actual(
            id_aprobacion_actual
        )
        if texto(aprobador.get("ID_DOCUMENTO")) != id_documento:
            raise ValueError(
                "El encargado indicado no pertenece al documento"
            )

        evento = seleccionar_evento_diagnostico(
            id_documento=id_documento,
            id_evento=id_evento,
            documento=documento,
        )

        id_version = (
            texto(aprobador.get("ID_VERSION_TRABAJADA"))
            or texto(documento.get("ID_VERSION_ACTUAL"))
        )
        link_documento = normalizar_url_appsheet(
            documento.get("GOOGLE_DOC_URL")
        )

        if id_version:
            try:
                version = buscar_version_por_id(id_version)
                link_documento_version = normalizar_url_appsheet(
                    version.get("GOOGLE_DOC_URL")
                )
                if link_documento_version:
                    link_documento = link_documento_version
            except Exception:
                traceback.print_exc()

        link_appsheet = construir_link_appsheet(id_documento)
        historial_texto, historial_html = (
            construir_historial_comentarios(
                id_documento=id_documento,
                id_evento_excluir=texto(
                    evento.get("ID_EVENTO")
                ),
            )
        )

        movimiento = (
            texto(data.get("movimiento"))
            or texto(evento.get("TIPO_EVENTO"))
            or "Actualización del flujo documental"
        )
        comentario_principal = (
            texto(data.get("comentario_principal"))
            or texto(evento.get("COMENTARIO"))
        )

        asunto, cuerpo_texto, cuerpo_html = (
            construir_email_notificacion(
                documento=documento,
                destinatario=aprobador,
                tipo_notificacion=tipo_notificacion,
                movimiento=movimiento,
                comentario_principal=comentario_principal,
                historial_texto=historial_texto,
                historial_html=historial_html,
                link_documento=link_documento,
                link_appsheet=link_appsheet,
            )
        )

        return jsonify(
            {
                "ok": True,
                "solo_diagnostico": True,
                "no_envia_email": True,
                "no_crea_notificacion": True,
                "id_documento": id_documento,
                "id_evento": texto(evento.get("ID_EVENTO")),
                "id_aprobacion_actual": id_aprobacion_actual,
                "destinatario_email": obtener_email_notificacion(
                    aprobador
                ),
                "destinatario_nombre": texto(
                    aprobador.get("NOMBRE")
                ),
                "tipo_notificacion": tipo_notificacion,
                "link_documento": link_documento,
                "link_appsheet": link_appsheet,
                "asunto": asunto,
                "movimiento": movimiento,
                "comentario_principal": comentario_principal,
                "historial_texto": historial_texto,
                "cuerpo_texto": cuerpo_texto,
                "cuerpo_html": cuerpo_html,
            }
        )

    except PermissionError as exc:
        return {"error": str(exc)}, 403

    except (ValueError, LookupError) as exc:
        return {"error": str(exc)}, 400

    except Exception as exc:
        traceback.print_exc()
        return {"error": str(exc)}, 500


def marcar_envio_firma_en_proceso(
    id_documento: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "ESTADO_FIRMA": "Enviando",
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "OBSERVACION_ACTUAL": "",
            }
        ],
    )


def actualizar_documento_enviado_firma(
    *,
    id_documento: str,
    usuario: str,
    destinatarios: list[str],
    mensaje_adicional: str,
    message_id: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "ESTADO": "En firma",
                "ESTADO_FIRMA": "Pendiente",
                "DESTINATARIOS_FIRMA": ", ".join(destinatarios),
                "MENSAJE_ADICIONAL_FIRMA": mensaje_adicional,
                "ENVIADO_FIRMA_POR": usuario,
                "EMAIL_FIRMA_MESSAGE_ID": message_id,
                "FECHA_ENVIO_FIRMA": fecha,
                "ULTIMO_ENVIADO_POR": usuario,
                "FECHA_ULTIMO_ENVIO": fecha,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "ACCION_SOLICITADA": "",
                "OBSERVACION_ACTUAL": "",
            }
        ],
    )


def restaurar_documento_tras_error_envio_firma(
    id_documento: str,
    mensaje: str,
) -> None:
    try:
        appsheet_action(
            TABLA_DOCUMENTOS,
            "Edit",
            [
                {
                    "ID_DOCUMENTO": id_documento,
                    "ESTADO_FIRMA": "No iniciado",
                    "ACCION_SOLICITADA": "",
                    "OBSERVACION_ACTUAL": mensaje[:1000],
                    "FECHA_ULTIMA_ACTUALIZACION": ahora_iso(),
                }
            ],
        )
    except Exception:
        traceback.print_exc()


def crear_evento_enviado_firma(
    *,
    id_documento: str,
    id_version: str,
    usuario: str,
    fecha: str,
    destinatarios: list[str],
    pdf_nombre: str,
    docx_nombre: str,
    message_id: str,
) -> None:
    appsheet_action(
        TABLA_EVENTOS,
        "Add",
        [
            {
                "ID_EVENTO": nuevo_id(),
                "ID_DOCUMENTO": id_documento,
                "ID_VERSION": id_version,
                "ID_APROBACION_ACTUAL": "",
                "TIPO_EVENTO": "Enviado a firma",
                "ESTADO_ANTERIOR": "Listo para firma",
                "ESTADO_NUEVO": "En firma",
                "USUARIO": usuario,
                "FECHA_EVENTO": fecha,
                "COMENTARIO": (
                    f"Se enviaron {pdf_nombre} y {docx_nombre} a "
                    f"{', '.join(destinatarios)}. "
                    f"Gmail message ID: {message_id}."
                ),
            }
        ],
    )


@app.route("/enviar-firma", methods=["POST"])
def enviar_firma():
    id_documento = ""
    correo_enviado = False
    message_id = ""
    fecha_envio = ""
    datos_finales: dict[str, Any] = {}

    try:
        validar_configuracion()
        validar_token()

        data = request.get_json(silent=True) or {}
        id_documento = texto(data.get("id_documento"))
        usuario = texto(data.get("usuario"))
        destinatarios_entrada = data.get("destinatarios")
        mensaje_adicional = texto(data.get("mensaje_adicional"))

        if not id_documento:
            return {"error": "Falta id_documento"}, 400

        documento = buscar_documento(id_documento)

        estado = texto(documento.get("ESTADO"))
        estado_firma = texto(documento.get("ESTADO_FIRMA"))
        message_id_existente = texto(
            documento.get("EMAIL_FIRMA_MESSAGE_ID")
        )

        if estado == "En firma" and message_id_existente:
            return jsonify(
                {
                    "ok": True,
                    "ya_procesado": True,
                    "id_documento": id_documento,
                    "estado": estado,
                    "estado_firma": estado_firma,
                    "message_id": message_id_existente,
                    "fecha_envio_firma": texto(
                        documento.get("FECHA_ENVIO_FIRMA")
                    ),
                }
            )

        if estado_firma == "Enviando":
            return {
                "error": (
                    "Ya existe un envío en proceso para este documento. "
                    "Revisa el registro antes de volver a intentarlo."
                )
            }, 409

        if estado != "Listo para firma":
            raise ValueError(
                "Solo se puede enviar a firma un documento en estado "
                f"Listo para firma. Estado actual: {estado!r}"
            )

        if estado_firma not in {"", "No iniciado", "Error"}:
            raise ValueError(
                "El estado de firma no permite iniciar el envío. "
                f"Estado actual: {estado_firma!r}"
            )

        pdf_id = texto(documento.get("PDF_PARA_FIRMA_ID"))
        google_doc_id = texto(documento.get("GOOGLE_DOC_ID"))
        id_version = texto(documento.get("ID_VERSION_ACTUAL"))
        if not pdf_id:
            raise ValueError("Documentos no tiene PDF_PARA_FIRMA_ID")
        if not google_doc_id:
            raise ValueError(
                "Documentos no tiene GOOGLE_DOC_ID para exportar el DOCX"
            )
        if not id_version:
            raise ValueError("Documentos no tiene ID_VERSION_ACTUAL")

        usuario_registrado = texto(
            documento.get("ENVIADO_FIRMA_POR")
        )
        usuario = usuario or usuario_registrado
        if not usuario or not _EMAIL_RE.fullmatch(usuario.lower()):
            raise ValueError("El usuario que envía a firma no es válido")
        if (
            usuario_registrado
            and usuario_registrado.lower() != usuario.lower()
        ):
            raise PermissionError(
                "El usuario del webhook no coincide con quien solicitó "
                "el envío a firma"
            )

        if destinatarios_entrada in (None, ""):
            destinatarios_entrada = documento.get(
                "DESTINATARIOS_FIRMA"
            )
        destinatarios = normalizar_destinatarios(
            destinatarios_entrada
        )
        if not destinatarios:
            raise ValueError("No se indicaron destinatarios para la firma")

        if not mensaje_adicional:
            mensaje_adicional = texto(
                documento.get("MENSAJE_ADICIONAL_FIRMA")
            )

        id_plantilla = texto(documento.get("ID_PLANTILLA"))
        plantilla = buscar_plantilla(id_plantilla)

        fecha_envio = ahora_iso()
        asunto, cuerpo, cuerpo_html = construir_email_firma_externo(
            documento=documento,
            plantilla=plantilla,
            usuario=usuario,
            mensaje_adicional=mensaje_adicional,
            fecha_envio=fecha_envio,
        )

        marcar_envio_firma_en_proceso(
            id_documento=id_documento,
            fecha=fecha_envio,
        )

        drive_service = obtener_drive_service()
        gmail_service = obtener_gmail_service()
        pdf_bytes, pdf_nombre = descargar_pdf_drive(
            drive_service=drive_service,
            file_id=pdf_id,
        )
        docx_bytes, docx_nombre = exportar_docx_drive(
            drive_service=drive_service,
            google_doc_id=google_doc_id,
            nombre_base=pdf_nombre,
        )

        respuesta_gmail = enviar_email_con_pdf(
            gmail_service=gmail_service,
            destinatarios=destinatarios,
            asunto=asunto,
            cuerpo=cuerpo,
            pdf_bytes=pdf_bytes,
            pdf_nombre=pdf_nombre,
            docx_bytes=docx_bytes,
            docx_nombre=docx_nombre,
            reply_to=usuario,
            cuerpo_html=cuerpo_html,
        )
        correo_enviado = True
        message_id = respuesta_gmail["message_id"]

        datos_finales = {
            "id_documento": id_documento,
            "usuario": usuario,
            "destinatarios": destinatarios,
            "mensaje_adicional": mensaje_adicional,
            "message_id": message_id,
            "fecha": fecha_envio,
        }
        actualizar_documento_enviado_firma(**datos_finales)

        advertencias: list[str] = []
        try:
            crear_evento_enviado_firma(
                id_documento=id_documento,
                id_version=id_version,
                usuario=usuario,
                fecha=fecha_envio,
                destinatarios=destinatarios,
                pdf_nombre=pdf_nombre,
                docx_nombre=docx_nombre,
                message_id=message_id,
            )
        except Exception as exc_evento:
            traceback.print_exc()
            advertencias.append(
                "El correo se envió, pero no se pudo crear el evento: "
                f"{exc_evento}"
            )

        return jsonify(
            {
                "ok": True,
                "ya_procesado": False,
                "id_documento": id_documento,
                "estado": "En firma",
                "estado_firma": "Pendiente",
                "message_id": message_id,
                "destinatarios": destinatarios,
                "pdf_id": pdf_id,
                "pdf_nombre": pdf_nombre,
                "google_doc_id": google_doc_id,
                "docx_nombre": docx_nombre,
                "advertencias": advertencias,
            }
        )

    except PermissionError as exc:
        if id_documento and not correo_enviado:
            restaurar_documento_tras_error_envio_firma(
                id_documento,
                str(exc),
            )
        return {"error": str(exc)}, 403

    except (ValueError, LookupError) as exc:
        if id_documento and not correo_enviado:
            restaurar_documento_tras_error_envio_firma(
                id_documento,
                str(exc),
            )
        return {"error": str(exc)}, 400

    except Exception as exc:
        traceback.print_exc()

        # Si Gmail ya respondió con éxito, nunca se reinicia a "No iniciado",
        # porque un reintento automático podría enviar el correo dos veces.
        if id_documento and correo_enviado and datos_finales:
            try:
                actualizar_documento_enviado_firma(**datos_finales)
            except Exception:
                traceback.print_exc()
        elif id_documento:
            restaurar_documento_tras_error_envio_firma(
                id_documento,
                str(exc),
            )

        return {
            "error": str(exc),
            "correo_enviado": correo_enviado,
            "message_id": message_id,
        }, 500


# -----------------------------------------------------------------------------
# Flujo: registrar PDF firmado y cerrar documento
# -----------------------------------------------------------------------------


def extraer_drive_id(valor: str) -> str:
    """Extrae un ID de Drive desde una URL conocida; si no hay, devuelve vacío."""
    valor = texto(valor)
    if not valor:
        return ""

    patrones = (
        r"/d/([A-Za-z0-9_-]{20,})",
        r"/file/d/([A-Za-z0-9_-]{20,})",
        r"/document/d/([A-Za-z0-9_-]{20,})",
    )
    for patron in patrones:
        coincidencia = re.search(patron, valor)
        if coincidencia:
            return coincidencia.group(1)

    try:
        consulta = parse_qs(urlparse(valor).query)
        candidato = texto((consulta.get("id") or [""])[0])
        if re.fullmatch(r"[A-Za-z0-9_-]{20,}", candidato):
            return candidato
    except Exception:
        pass

    return ""


def nombre_archivo_desde_valor_appsheet(valor: str) -> str:
    """Obtiene el nombre final desde un File relativo o una URL."""
    valor = unquote(texto(valor)).replace("\\", "/")
    if not valor:
        return ""

    if "://" in valor:
        ruta = urlparse(valor).path
    else:
        ruta = valor.split("?", 1)[0].split("#", 1)[0]

    nombre = PurePosixPath(ruta).name.strip()
    return nombre


EXTENSIONES_ARCHIVO_OBSERVACION_FIRMA = {".pdf", ".doc", ".docx"}


def validar_archivo_observacion_firma(valor: str) -> tuple[str, str]:
    """
    Valida el archivo opcional recibido como respaldo de observaciones de firma.

    El valor corresponde al File/path que AppSheet ya guardó. Cloud Run no
    descarga ni transforma este archivo: solo conserva su referencia histórica
    en Documento_Versiones.
    """
    ruta = texto(valor)
    if not ruta:
        return "", ""

    nombre = nombre_archivo_desde_valor_appsheet(ruta)
    if not nombre:
        raise ValueError(
            "No se pudo determinar el nombre del archivo de observaciones"
        )

    extension = PurePosixPath(nombre.lower()).suffix
    if extension not in EXTENSIONES_ARCHIVO_OBSERVACION_FIRMA:
        permitidas = ", ".join(
            sorted(EXTENSIONES_ARCHIVO_OBSERVACION_FIRMA)
        )
        raise ValueError(
            "El archivo de observaciones debe ser PDF, DOC o DOCX. "
            f"Extensiones permitidas: {permitidas}"
        )

    return ruta, nombre


def marcar_version_observada_firma(
    *,
    id_version: str,
    fecha: str,
    usuario: str,
    comentario: str,
    archivo_observacion_firma: str = "",
) -> tuple[str, str]:
    """
    Cierra la versión enviada a firma como Observada y, si existe, conserva
    el archivo devuelto por el firmante externo como respaldo histórico.

    Este archivo NO se usa para crear la nueva versión. El nuevo borrador se
    sigue copiando exclusivamente desde GOOGLE_DOC_ID de la versión Para firma.
    """
    ruta_archivo, nombre_archivo = validar_archivo_observacion_firma(
        archivo_observacion_firma
    )

    cambios: dict[str, Any] = {
        "ID_VERSION": id_version,
        "ESTADO_VERSION": "Observada",
        "FECHA_CIERRE": fecha,
        "COMENTARIO_OBSERVACION_FIRMA": comentario,
    }

    if ruta_archivo:
        cambios.update(
            {
                "ARCHIVO_OBSERVACION_FIRMA": ruta_archivo,
                "NOMBRE_ARCHIVO_OBSERVACION_FIRMA": nombre_archivo,
                "FECHA_ARCHIVO_OBSERVACION_FIRMA": fecha,
                "CARGADO_POR_OBSERVACION_FIRMA": usuario,
            }
        )

    appsheet_action(
        TABLA_VERSIONES,
        "Edit",
        [cambios],
    )

    return ruta_archivo, nombre_archivo


def obtener_metadata_pdf_drive(
    drive_service: Any,
    file_id: str,
) -> dict[str, Any]:
    archivo = (
        drive_service.files()
        .get(
            fileId=file_id,
            fields=(
                "id,name,mimeType,size,createdTime,modifiedTime,parents,"
                "webViewLink,webContentLink,md5Checksum"
            ),
            supportsAllDrives=True,
        )
        .execute()
    )

    nombre = texto(archivo.get("name"))
    mime_type = texto(archivo.get("mimeType")).lower()
    tamano = entero(archivo.get("size") or 0, "size")

    if mime_type != "application/pdf" and not nombre.lower().endswith(".pdf"):
        raise ValueError(
            f"El archivo cargado no es PDF: {nombre!r} ({mime_type!r})"
        )
    if tamano <= 0:
        raise ValueError("El PDF firmado está vacío")

    return archivo


def buscar_pdf_cargado_appsheet(
    drive_service: Any,
    valor_pdf: str,
) -> dict[str, Any]:
    """
    Localiza el archivo que AppSheet guardó en Drive.

    AppSheet almacena en la columna File el nombre o ruta relativa. Primero se
    acepta una URL de Drive; en caso contrario se busca por el nombre exacto.
    """
    drive_id = extraer_drive_id(valor_pdf)
    if drive_id:
        return obtener_metadata_pdf_drive(drive_service, drive_id)

    nombre = nombre_archivo_desde_valor_appsheet(valor_pdf)
    if not nombre:
        raise ValueError("No se pudo obtener el nombre desde PDF_FIRMADO")
    if not nombre.lower().endswith(".pdf"):
        raise ValueError("PDF_FIRMADO debe corresponder a un archivo .pdf")

    nombre_q = escapar_consulta_drive(nombre)
    consulta = f"name = '{nombre_q}' and trashed = false"
    respuesta = (
        drive_service.files()
        .list(
            q=consulta,
            fields=(
                "files(id,name,mimeType,size,createdTime,modifiedTime,parents,"
                "webViewLink,webContentLink,md5Checksum)"
            ),
            spaces="drive",
            corpora="user",
            orderBy="modifiedTime desc",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            pageSize=20,
        )
        .execute()
    )
    archivos = respuesta.get("files", [])

    if not archivos:
        raise LookupError(
            "No se encontró en Google Drive el archivo cargado en "
            f"PDF_FIRMADO: {nombre!r}. Confirma que AppSheet y Cloud Run "
            "usen una cuenta con acceso al mismo Drive."
        )

    archivos_pdf = [
        archivo
        for archivo in archivos
        if texto(archivo.get("mimeType")).lower() == "application/pdf"
        or texto(archivo.get("name")).lower().endswith(".pdf")
    ]

    if len(archivos_pdf) > 1:
        ids = ", ".join(texto(fila.get("id")) for fila in archivos_pdf[:5])
        raise RuntimeError(
            "Se encontraron varios archivos PDF con el mismo nombre en Drive. "
            f"Nombre={nombre!r}; IDs={ids}. Renombra o elimina los duplicados."
        )

    if not archivos_pdf:
        raise ValueError(f"El archivo encontrado no es PDF: {nombre!r}")

    return obtener_metadata_pdf_drive(
        drive_service,
        texto(archivos_pdf[0].get("id")),
    )


def copiar_pdf_firmado_o_reutilizar(
    drive_service: Any,
    source_file_id: str,
    folder_id: str,
    nombre_final: str,
) -> dict[str, str]:
    """Copia el PDF firmado a la carpeta del documento con nombre canónico."""
    existente = buscar_archivo_en_carpeta(
        drive_service=drive_service,
        folder_id=folder_id,
        nombre_archivo=nombre_final,
    )
    if existente:
        metadata = obtener_metadata_pdf_drive(
            drive_service,
            existente["id"],
        )
        return {
            "id": texto(metadata.get("id")),
            "name": texto(metadata.get("name")) or nombre_final,
            "url": texto(metadata.get("webViewLink"))
            or texto(metadata.get("webContentLink"))
            or f"https://drive.google.com/file/d/{metadata['id']}/view",
        }

    copia = (
        drive_service.files()
        .copy(
            fileId=source_file_id,
            body={
                "name": nombre_final,
                "parents": [folder_id],
            },
            fields=(
                "id,name,mimeType,size,webViewLink,webContentLink,md5Checksum"
            ),
            supportsAllDrives=True,
        )
        .execute()
    )

    file_id = texto(copia.get("id"))
    if not file_id:
        raise RuntimeError("Drive no devolvió el ID del PDF firmado copiado")

    return {
        "id": file_id,
        "name": texto(copia.get("name")) or nombre_final,
        "url": texto(copia.get("webViewLink"))
        or texto(copia.get("webContentLink"))
        or f"https://drive.google.com/file/d/{file_id}/view",
    }


def buscar_version_firmada(
    id_documento: str,
    numero_version: int,
) -> dict[str, Any] | None:
    versiones = buscar_versiones_documento(id_documento)
    encontradas = [
        fila
        for fila in versiones
        if entero(fila.get("NUMERO_VERSION"), "NUMERO_VERSION")
        == numero_version
        and texto(fila.get("ETAPA")).lower() == "firmado"
    ]

    if len(encontradas) > 1:
        raise RuntimeError(
            "Existen varias filas Firmado para la misma versión documental"
        )
    return encontradas[0] if encontradas else None


def marcar_registro_firma_en_proceso(
    id_documento: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "ESTADO_FIRMA": "Procesando",
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
            }
        ],
    )


def crear_version_firmada(
    *,
    id_version: str,
    id_documento: str,
    id_version_origen: str,
    numero_version: int,
    numero_revision: int,
    nombre_archivo: str,
    google_doc_id: str,
    google_doc_url: str,
    pdf_id: str,
    pdf_url: str,
    comentario: str,
    usuario: str,
    fecha: str,
) -> None:
    google_doc_id_limpio = texto(google_doc_id)
    google_doc_url_limpia = normalizar_url_appsheet(google_doc_url)

    # Para esta versión, el ID de Google Docs es la fuente más confiable.
    # Si AppSheet devolvió una URL enriquecida o inválida, la reconstruimos.
    if google_doc_id_limpio:
        google_doc_url_limpia = (
            f"https://docs.google.com/document/d/"
            f"{google_doc_id_limpio}/edit"
        )

    pdf_url_limpia = normalizar_url_appsheet(pdf_url)
    if not pdf_url_limpia and texto(pdf_id):
        pdf_url_limpia = (
            f"https://drive.google.com/file/d/{texto(pdf_id)}/view"
        )

    app.logger.info(
        "Creando versión firmada: google_doc_id=%s, google_doc_url=%s",
        google_doc_id_limpio,
        google_doc_url_limpia,
    )

    appsheet_action(
        TABLA_VERSIONES,
        "Add",
        [
            {
                "ID_VERSION": id_version,
                "ID_DOCUMENTO": id_documento,
                "ID_VERSION_ORIGEN": id_version_origen,
                "NUMERO_VERSION": numero_version,
                "NUMERO_REVISION": numero_revision,
                "ETAPA": "Firmado",
                "ESTADO_VERSION": "Firmada",
                "NOMBRE_ARCHIVO": nombre_archivo,
                "GOOGLE_DOC_ID": google_doc_id_limpio,
                "GOOGLE_DOC_URL": google_doc_url_limpia,
                "PDF_VERSION_ID": texto(pdf_id),
                "PDF_VERSION_URL": pdf_url_limpia,
                "ID_APROBACION_RESPONSABLE": "",
                "ORDEN_RESPONSABLE": "",
                "MOTIVO_CREACION": "Carga de documento firmado",
                "COMENTARIO_CAMBIO": comentario,
                "CREADO_POR": usuario,
                "FECHA_CREACION": fecha,
                "FECHA_CIERRE": fecha,
            }
        ],
    )


def actualizar_documento_proceso_terminado(
    *,
    id_documento: str,
    id_version_final: str,
    pdf_final: dict[str, str],
    usuario: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "ESTADO": "Proceso terminado",
                "ESTADO_FIRMA": "Firmado",
                "ID_VERSION_ACTUAL": id_version_final,
                "PDF_FIRMADO_ID": pdf_final["id"],
                "PDF_FIRMADO_URL": pdf_final["url"],
                "FECHA_FIRMA_COMPLETA": fecha,
                "CARGADO_POR": usuario,
                "ULTIMO_ENVIADO_POR": usuario,
                "FECHA_ULTIMO_ENVIO": fecha,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "FECHA_CIERRE": fecha,
                "ACCION_SOLICITADA": "",
                "OBSERVACION_ACTUAL": "",
            }
        ],
    )


def restaurar_documento_tras_error_registro_firma(
    id_documento: str,
    mensaje: str,
) -> None:
    try:
        appsheet_action(
            TABLA_DOCUMENTOS,
            "Edit",
            [
                {
                    "ID_DOCUMENTO": id_documento,
                    "ESTADO_FIRMA": "Pendiente",
                    "ACCION_SOLICITADA": "",
                    "OBSERVACION_ACTUAL": mensaje[:1000],
                    "FECHA_ULTIMA_ACTUALIZACION": ahora_iso(),
                }
            ],
        )
    except Exception:
        traceback.print_exc()


def crear_eventos_cierre_firma(
    *,
    id_documento: str,
    id_version: str,
    usuario: str,
    fecha: str,
    nombre_origen: str,
    nombre_final: str,
    comentario: str,
) -> dict[str, Any]:
    """Crea los eventos de cierre y devuelve el evento Proceso terminado."""
    detalle_carga = (
        f"Se cargó {nombre_origen} y se archivó como {nombre_final}."
    )
    if comentario:
        detalle_carga += f" Comentario: {comentario}"

    evento_carga: dict[str, Any] = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION": id_version,
        "ID_APROBACION_ACTUAL": "",
        "TIPO_EVENTO": "PDF firmado cargado",
        "ESTADO_ANTERIOR": "En firma",
        "ESTADO_NUEVO": "En firma",
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": detalle_carga,
    }
    evento_cierre: dict[str, Any] = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION": id_version,
        "ID_APROBACION_ACTUAL": "",
        "TIPO_EVENTO": "Proceso terminado",
        "ESTADO_ANTERIOR": "En firma",
        "ESTADO_NUEVO": "Proceso terminado",
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": (
            "El documento quedó cerrado con su PDF firmado definitivo."
        ),
    }

    appsheet_action(
        TABLA_EVENTOS,
        "Add",
        [evento_carga, evento_cierre],
    )
    return evento_cierre


@app.route("/registrar-firma", methods=["POST"])
def registrar_firma():
    id_documento = ""
    pdf_final_creado = False
    documento_cerrado = False

    try:
        validar_configuracion()
        validar_token()

        data = request.get_json(silent=True) or {}
        id_documento = texto(data.get("id_documento"))
        usuario = texto(data.get("usuario"))
        valor_pdf_firmado = texto(data.get("pdf_firmado"))
        comentario = texto(data.get("comentario"))

        if not id_documento:
            return {"error": "Falta id_documento"}, 400

        documento = buscar_documento(id_documento)
        estado = texto(documento.get("ESTADO"))
        estado_firma = texto(documento.get("ESTADO_FIRMA"))
        pdf_final_existente = texto(documento.get("PDF_FIRMADO_ID"))

        if estado == "Proceso terminado" and pdf_final_existente:
            notificaciones_reintento: list[dict[str, Any]] = []
            advertencias_reintento: list[str] = []
            try:
                (
                    notificaciones_reintento,
                    advertencias_reintento,
                ) = reanudar_notificaciones_cierre_proceso(
                    documento=documento,
                    datos_solicitud=data,
                )
            except Exception as exc_notificacion:
                traceback.print_exc()
                advertencias_reintento.append(
                    "El proceso ya estaba cerrado, pero no se pudieron "
                    "reanudar sus notificaciones: "
                    f"{exc_notificacion}"
                )

            return jsonify(
                {
                    "ok": True,
                    "ya_procesado": True,
                    "id_documento": id_documento,
                    "estado": estado,
                    "estado_firma": estado_firma,
                    "pdf_firmado_id": pdf_final_existente,
                    "pdf_firmado_url": texto(
                        documento.get("PDF_FIRMADO_URL")
                    ),
                    "notificaciones": notificaciones_reintento,
                    "advertencias": advertencias_reintento,
                }
            )

        if estado_firma == "Procesando":
            return {
                "error": (
                    "Ya existe un cierre de firma en proceso. Revisa el registro "
                    "antes de volver a intentarlo."
                )
            }, 409

        if estado != "En firma":
            raise ValueError(
                "Solo se puede registrar la firma cuando el documento está "
                f"En firma. Estado actual: {estado!r}"
            )
        if estado_firma not in {"Pendiente", "Error"}:
            raise ValueError(
                "El estado de firma no permite cerrar el documento. "
                f"Estado actual: {estado_firma!r}"
            )

        valor_pdf_firmado = valor_pdf_firmado or texto(
            documento.get("PDF_FIRMADO")
        )
        if not valor_pdf_firmado:
            raise ValueError("Debe cargar un archivo en PDF_FIRMADO")

        usuario_registrado = texto(documento.get("CARGADO_POR"))
        usuario = usuario or usuario_registrado
        if not usuario or not _EMAIL_RE.fullmatch(usuario.lower()):
            raise ValueError("CARGADO_POR debe ser un correo válido")
        if (
            usuario_registrado
            and usuario_registrado.lower() != usuario.lower()
        ):
            raise PermissionError(
                "El usuario del webhook no coincide con CARGADO_POR"
            )

        id_version_origen = texto(documento.get("ID_VERSION_ACTUAL"))
        if not id_version_origen:
            raise ValueError("Documentos no tiene ID_VERSION_ACTUAL")
        version_origen = buscar_version_por_id(id_version_origen)

        numero_version = entero(
            documento.get("VERSION_ACTUAL"),
            "VERSION_ACTUAL",
        )
        numero_revision = entero(
            documento.get("REVISION_ACTUAL") or 0,
            "REVISION_ACTUAL",
        )

        id_plantilla = texto(documento.get("ID_PLANTILLA"))
        plantilla = buscar_plantilla(id_plantilla)
        folder_id = texto(plantilla.get("CARPETA_DESTINO_ID"))
        if not folder_id:
            raise ValueError("La plantilla no tiene CARPETA_DESTINO_ID")

        fecha = ahora_iso()
        marcar_registro_firma_en_proceso(id_documento, fecha)

        drive_service = obtener_drive_service()
        pdf_subido = buscar_pdf_cargado_appsheet(
            drive_service,
            valor_pdf_firmado,
        )

        titulo = limpiar_nombre_archivo(texto(documento.get("TITULO")))
        nombre_final = f"{titulo}_V{numero_version:02d}_FIRMADO.pdf"
        pdf_final = copiar_pdf_firmado_o_reutilizar(
            drive_service=drive_service,
            source_file_id=texto(pdf_subido.get("id")),
            folder_id=folder_id,
            nombre_final=nombre_final,
        )
        pdf_final_creado = True

        version_firmada_existente = buscar_version_firmada(
            id_documento,
            numero_version,
        )
        if version_firmada_existente:
            id_version_final = texto(
                version_firmada_existente.get("ID_VERSION")
            )
        else:
            id_version_final = nuevo_id()
            actualizar_estado_version(
                id_version=id_version_origen,
                estado_version="Cerrada",
                fecha_cierre=fecha,
            )
            google_doc_id_origen = texto(
                version_origen.get("GOOGLE_DOC_ID")
            )
            google_doc_url_origen = normalizar_url_appsheet(
                version_origen.get("GOOGLE_DOC_URL")
            )

            crear_version_firmada(
                id_version=id_version_final,
                id_documento=id_documento,
                id_version_origen=id_version_origen,
                numero_version=numero_version,
                numero_revision=numero_revision,
                nombre_archivo=nombre_final,
                google_doc_id=google_doc_id_origen,
                google_doc_url=google_doc_url_origen,
                pdf_id=pdf_final["id"],
                pdf_url=pdf_final["url"],
                comentario=comentario,
                usuario=usuario,
                fecha=fecha,
            )

        actualizar_documento_proceso_terminado(
            id_documento=id_documento,
            id_version_final=id_version_final,
            pdf_final=pdf_final,
            usuario=usuario,
            fecha=fecha,
        )
        documento_cerrado = True

        advertencias: list[str] = []
        evento_cierre: dict[str, Any] | None = None
        try:
            evento_cierre = crear_eventos_cierre_firma(
                id_documento=id_documento,
                id_version=id_version_final,
                usuario=usuario,
                fecha=fecha,
                nombre_origen=texto(pdf_subido.get("name")),
                nombre_final=pdf_final["name"],
                comentario=comentario,
            )
        except Exception as exc_evento:
            traceback.print_exc()
            advertencias.append(
                "El documento se cerró, pero no se pudieron crear todos los "
                f"eventos: {exc_evento}"
            )

        notificaciones: list[dict[str, Any]] = []
        if evento_cierre is not None:
            try:
                documento_actualizado = buscar_documento(id_documento)
                cadena = buscar_cadena_documento_version(
                    id_documento=id_documento,
                    numero_version=numero_version,
                )
                notificaciones = ejecutar_notificaciones_cierre_proceso(
                    documento=documento_actualizado,
                    evento=evento_cierre,
                    cadena=cadena,
                )
                fallidas = [
                    resultado
                    for resultado in notificaciones
                    if not resultado.get("ok")
                ]
                if fallidas:
                    advertencias.append(
                        f"{len(fallidas)} notificación(es) de cierre quedaron "
                        "omitidas o con error. Revisa Documento_Notificaciones."
                    )
            except Exception as exc_notificacion:
                traceback.print_exc()
                advertencias.append(
                    "El documento se cerró correctamente, pero falló el "
                    f"proceso de notificaciones internas: {exc_notificacion}"
                )

        return jsonify(
            {
                "ok": True,
                "ya_procesado": False,
                "id_documento": id_documento,
                "estado": "Proceso terminado",
                "estado_firma": "Firmado",
                "id_version_final": id_version_final,
                "pdf_firmado_id": pdf_final["id"],
                "pdf_firmado_url": pdf_final["url"],
                "pdf_firmado_nombre": pdf_final["name"],
                "notificaciones": notificaciones,
                "advertencias": advertencias,
            }
        )

    except PermissionError as exc:
        if id_documento:
            restaurar_documento_tras_error_registro_firma(
                id_documento,
                str(exc),
            )
        return {"error": str(exc)}, 403

    except (ValueError, LookupError) as exc:
        if id_documento:
            restaurar_documento_tras_error_registro_firma(
                id_documento,
                str(exc),
            )
        return {"error": str(exc)}, 400

    except Exception as exc:
        traceback.print_exc()
        if id_documento:
            restaurar_documento_tras_error_registro_firma(
                id_documento,
                str(exc),
            )
        return {
            "error": str(exc),
            "pdf_final_creado": pdf_final_creado,
        }, 500


# -----------------------------------------------------------------------------
# Flujo: observaciones/rechazo después del envío a firma
# -----------------------------------------------------------------------------


def actualizar_documento_reinicio_firma(
    *,
    id_documento: str,
    numero_version: int,
    id_version: str,
    copia: dict[str, str],
    primer_encargado: dict[str, Any],
    usuario: str,
    fecha: str,
) -> None:
    """Apunta la cabecera a un nuevo borrador y reinicia el flujo completo."""
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "ESTADO": "Borrador",
                "VERSION_ACTUAL": numero_version,
                "REVISION_ACTUAL": 0,
                "ID_VERSION_ACTUAL": id_version,
                "GOOGLE_DOC_ID": copia["id"],
                "GOOGLE_DOC_URL": copia["url"],
                "ORDEN_ACTUAL": primer_encargado["ORDEN"],
                "ID_APROBACION_ACTUAL": primer_encargado[
                    "ID_APROBACION_ACTUAL"
                ],
                "ENCARGADO_ACTUAL_NOMBRE": primer_encargado.get(
                    "NOMBRE", ""
                ),
                "ENCARGADO_ACTUAL_EMAIL": primer_encargado.get(
                    "APROBADOR", ""
                ),
                # El resultado del ciclo de firma anterior queda en el
                # historial. La cabecera ya no debe apuntar a su PDF.
                "PDF_PARA_FIRMA_ID": "",
                "PDF_PARA_FIRMA_URL": "",
                "ESTADO_FIRMA": "Observado",
                "PDF_FIRMADO": "",
                "PDF_FIRMADO_ID": "",
                "PDF_FIRMADO_URL": "",
                "FECHA_ENVIO_FIRMA": "",
                "FECHA_FIRMA_COMPLETA": "",
                "CARGADO_POR": "",
                "DESTINATARIOS_FIRMA": "",
                "MENSAJE_ADICIONAL_FIRMA": "",
                "ENVIADO_FIRMA_POR": "",
                # Campo transitorio usado por el formulario de rechazo. La
                # referencia histórica ya quedó copiada a Documento_Versiones.
                "ARCHIVO_OBSERVACION_FIRMA_TEMP": "",
                "EMAIL_FIRMA_MESSAGE_ID": "",
                "FECHA_CIERRE": "",
                "ULTIMO_ENVIADO_POR": usuario,
                "FECHA_ULTIMO_ENVIO": fecha,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "OBSERVACION_ACTUAL": "",
                "ACCION_SOLICITADA": "",
            }
        ],
    )


def crear_eventos_reinicio_firma(
    *,
    id_documento: str,
    id_version_observada: str,
    id_version_nueva: str,
    id_aprobacion_nueva: str,
    usuario: str,
    fecha: str,
    comentario: str,
    numero_version_nueva: int,
    nombre_archivo: str,
    nombre_archivo_observacion: str = "",
) -> tuple[list[str], dict[str, Any] | None]:
    """Crea la bitácora del reinicio y devuelve el evento principal."""
    advertencias: list[str] = []
    evento_reinicio: dict[str, Any] = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION": id_version_observada,
        "ID_APROBACION_ACTUAL": "",
        "TIPO_EVENTO": "Proceso reiniciado",
        "ESTADO_ANTERIOR": "En firma",
        "ESTADO_NUEVO": "Borrador",
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": (
            "El documento recibió observaciones durante la firma. "
            f"Motivo: {comentario}"
            + (
                f" Archivo de respaldo recibido: {nombre_archivo_observacion}."
                if nombre_archivo_observacion
                else ""
            )
        ),
    }
    evento_version: dict[str, Any] = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION": id_version_nueva,
        "ID_APROBACION_ACTUAL": id_aprobacion_nueva,
        "TIPO_EVENTO": "Nueva versión creada",
        "ESTADO_ANTERIOR": "Observado en firma",
        "ESTADO_NUEVO": "Borrador",
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": (
            f"Se creó la versión {numero_version_nueva}: "
            f"{nombre_archivo}. La cadena de aprobación se reinició "
            "desde el primer responsable."
        ),
    }

    try:
        appsheet_action(
            TABLA_EVENTOS,
            "Add",
            [evento_reinicio, evento_version],
        )
        return advertencias, evento_reinicio
    except Exception as exc:
        traceback.print_exc()
        advertencias.append(
            "La transición terminó, pero no se pudieron crear los eventos: "
            f"{exc}"
        )
        return advertencias, None


@app.route("/rechazar-firma", methods=["POST"])
def rechazar_firma():
    id_documento = ""

    try:
        validar_configuracion()
        validar_token()

        data = request.get_json(silent=True) or {}
        id_documento = texto(data.get("id_documento"))
        id_version_solicitud = texto(data.get("id_version_firma"))
        usuario = texto(data.get("usuario"))
        comentario = texto(data.get("comentario"))
        archivo_observacion_firma = texto(
            data.get("archivo_observacion_firma")
        )

        if not id_documento:
            return {"error": "Falta id_documento"}, 400
        if not comentario:
            return {
                "error": (
                    "El comentario es obligatorio para reiniciar el proceso "
                    "por observaciones durante la firma"
                )
            }, 400

        documento = buscar_documento(id_documento)
        id_version_documento = texto(documento.get("ID_VERSION_ACTUAL"))
        id_version_origen = id_version_solicitud or id_version_documento
        if not id_version_origen:
            raise ValueError("No se pudo identificar la versión enviada a firma")

        version_origen = buscar_version_por_id(id_version_origen)
        numero_version_anterior = entero(
            version_origen.get("NUMERO_VERSION"),
            "NUMERO_VERSION",
        )
        estado_version_origen = texto(
            version_origen.get("ESTADO_VERSION")
        )
        etapa_origen = texto(version_origen.get("ETAPA"))

        numero_version_documento = entero(
            documento.get("VERSION_ACTUAL"),
            "VERSION_ACTUAL",
        )
        estado_documento = texto(documento.get("ESTADO"))

        # Reintento del mismo webhook después de una transición ya terminada.
        # También intenta completar notificaciones pendientes o con error.
        if (
            estado_version_origen == "Observada"
            and numero_version_documento > numero_version_anterior
            and estado_documento in {"Borrador", "En revisión"}
        ):
            # Si AppSheet reintenta el webhook con un archivo que no alcanzó a
            # quedar registrado, completamos solamente el respaldo histórico.
            if archivo_observacion_firma:
                ruta_validada, nombre_validado = validar_archivo_observacion_firma(
                    archivo_observacion_firma
                )
                ruta_guardada = texto(
                    version_origen.get("ARCHIVO_OBSERVACION_FIRMA")
                )
                if ruta_validada and ruta_guardada != ruta_validada:
                    marcar_version_observada_firma(
                        id_version=id_version_origen,
                        fecha=ahora_iso(),
                        usuario=usuario or texto(documento.get("ENVIADO_FIRMA_POR")),
                        comentario=comentario,
                        archivo_observacion_firma=ruta_validada,
                    )
                    version_origen["ARCHIVO_OBSERVACION_FIRMA"] = ruta_validada
                    version_origen[
                        "NOMBRE_ARCHIVO_OBSERVACION_FIRMA"
                    ] = nombre_validado

            notificaciones_reintento: list[dict[str, Any]] = []
            advertencias_reintento: list[str] = []
            try:
                documento = buscar_documento(id_documento)
                (
                    notificaciones_reintento,
                    advertencias_reintento,
                ) = reanudar_notificaciones_reinicio_firma(
                    documento=documento,
                    id_version_observada=id_version_origen,
                    datos_solicitud=data,
                )
            except Exception as exc_notificacion:
                traceback.print_exc()
                advertencias_reintento.append(
                    "El reinicio ya estaba procesado, pero no se pudieron "
                    "reanudar sus notificaciones: "
                    f"{exc_notificacion}"
                )

            return jsonify(
                {
                    "ok": True,
                    "ya_procesado": True,
                    "id_documento": id_documento,
                    "estado": texto(documento.get("ESTADO")),
                    "estado_firma": texto(
                        documento.get("ESTADO_FIRMA")
                    ),
                    "numero_version": entero(
                        documento.get("VERSION_ACTUAL"),
                        "VERSION_ACTUAL",
                    ),
                    "numero_revision": entero(
                        documento.get("REVISION_ACTUAL") or 0,
                        "REVISION_ACTUAL",
                    ),
                    "id_version": texto(
                        documento.get("ID_VERSION_ACTUAL")
                    ),
                    "google_doc_id": texto(
                        documento.get("GOOGLE_DOC_ID")
                    ),
                    "google_doc_url": normalizar_url_appsheet(
                        documento.get("GOOGLE_DOC_URL")
                    ),
                    "archivo_observacion_firma": texto(
                        version_origen.get("ARCHIVO_OBSERVACION_FIRMA")
                    ),
                    "nombre_archivo_observacion_firma": texto(
                        version_origen.get(
                            "NOMBRE_ARCHIVO_OBSERVACION_FIRMA"
                        )
                    ),
                    "notificaciones": notificaciones_reintento,
                    "advertencias": advertencias_reintento,
                }
            )

        estado_firma = texto(documento.get("ESTADO_FIRMA"))
        if estado_documento != "En firma":
            raise ValueError(
                "Solo se puede rechazar por observaciones un documento En firma. "
                f"Estado actual: {estado_documento!r}"
            )
        if estado_firma != "Pendiente":
            raise ValueError(
                "El estado de firma debe ser Pendiente para reiniciar el proceso. "
                f"Estado actual: {estado_firma!r}"
            )
        if id_version_documento != id_version_origen:
            raise ValueError(
                "La versión enviada por AppSheet ya no coincide con la versión "
                "actual del documento"
            )
        if etapa_origen != "Para firma":
            raise ValueError(
                "La versión que se intenta observar no corresponde a la etapa "
                f"Para firma. Etapa encontrada: {etapa_origen!r}"
            )

        usuario_envio = texto(documento.get("ENVIADO_FIRMA_POR"))
        usuario = usuario or usuario_envio
        if not usuario or not _EMAIL_RE.fullmatch(usuario.lower()):
            raise ValueError("El usuario que reinicia la aprobación no es válido")
        if usuario_envio and usuario_envio.lower() != usuario.lower():
            raise PermissionError(
                "Solo el usuario que envió el documento a firma puede "
                "reiniciar la aprobación"
            )

        google_doc_id_origen = texto(
            version_origen.get("GOOGLE_DOC_ID")
        ) or texto(documento.get("GOOGLE_DOC_ID"))
        if not google_doc_id_origen:
            raise ValueError(
                "La versión Para firma no tiene GOOGLE_DOC_ID para crear "
                "el nuevo borrador"
            )

        cadena_anterior = buscar_cadena_documento_version(
            id_documento=id_documento,
            numero_version=numero_version_anterior,
        )
        if not cadena_anterior:
            raise ValueError(
                "No se encontró la cadena de aprobación de la versión enviada "
                "a firma"
            )

        id_plantilla = texto(documento.get("ID_PLANTILLA"))
        plantilla = buscar_plantilla(id_plantilla)
        folder_id = texto(plantilla.get("CARPETA_DESTINO_ID"))
        if not folder_id:
            raise ValueError("La plantilla no tiene CARPETA_DESTINO_ID")

        numero_version_nueva = numero_version_anterior + 1
        numero_revision_nueva = 0
        titulo = texto(documento.get("TITULO")) or f"Documento_{id_documento}"
        nombre_archivo = limpiar_nombre_archivo(
            f"{titulo}_V{numero_version_nueva:02d}_BORRADOR"
        )
        fecha = ahora_iso()
        drive_service = obtener_drive_service()

        version_existente = buscar_version_numero_revision(
            id_documento=id_documento,
            numero_version=numero_version_nueva,
            numero_revision=numero_revision_nueva,
        )

        if version_existente:
            id_version_nueva = texto(version_existente.get("ID_VERSION"))
            copia = {
                "id": texto(version_existente.get("GOOGLE_DOC_ID")),
                "url": normalizar_url_appsheet(
                    version_existente.get("GOOGLE_DOC_URL")
                ),
                "name": (
                    texto(version_existente.get("NOMBRE_ARCHIVO"))
                    or nombre_archivo
                ),
            }
            if not copia["id"]:
                raise RuntimeError(
                    "La nueva versión existente no tiene GOOGLE_DOC_ID"
                )
        else:
            id_version_nueva = nuevo_id()
            copia = copiar_archivo_o_reutilizar(
                drive_service=drive_service,
                source_file_id=google_doc_id_origen,
                folder_id=folder_id,
                nombre_archivo=nombre_archivo,
            )

        # La versión Para firma queda congelada. El nuevo borrador comienza
        # nuevamente con edición exclusiva para el primer responsable.
        for fila in cadena_anterior:
            email = texto(fila.get("APROBADOR"))
            if email:
                asegurar_permiso_rol(
                    drive_service=drive_service,
                    file_id=google_doc_id_origen,
                    email=email,
                    role="reader",
                )

        primer_anterior = cadena_anterior[0]
        email_primero = texto(primer_anterior.get("APROBADOR"))
        if not email_primero:
            raise ValueError("El primer responsable no tiene correo")

        permission_id_primero = asegurar_permiso_rol(
            drive_service=drive_service,
            file_id=copia["id"],
            email=email_primero,
            role="writer",
        )

        cadena_nueva_existente = buscar_cadena_documento_version(
            id_documento=id_documento,
            numero_version=numero_version_nueva,
        )

        if cadena_nueva_existente:
            primeros = [
                fila
                for fila in cadena_nueva_existente
                if entero(fila.get("ORDEN"), "ORDEN")
                == entero(primer_anterior.get("ORDEN"), "ORDEN")
            ]
            if len(primeros) != 1:
                raise RuntimeError(
                    "La cadena nueva existente no tiene un único primer "
                    "responsable"
                )
            primer_nuevo = primeros[0]
            actualizar_destino_cadena_reutilizada(
                destino=primer_nuevo,
                id_version_nueva=id_version_nueva,
                permission_id_destino=permission_id_primero,
                estado_destino="En elaboración",
                fecha=fecha,
            )
        else:
            filas_nuevas, primer_nuevo = construir_cadena_nueva_por_rechazo(
                cadena_anterior=cadena_anterior,
                id_documento=id_documento,
                id_plantilla=id_plantilla,
                numero_version_anterior=numero_version_anterior,
                numero_version_nueva=numero_version_nueva,
                indice_destino=0,
                id_version_nueva=id_version_nueva,
                permission_id_destino=permission_id_primero,
                fecha=fecha,
            )
            appsheet_action(
                TABLA_APROBADORES_ACTUAL,
                "Add",
                filas_nuevas,
            )

        if not version_existente:
            crear_registro_version_por_aprobacion(
                id_version=id_version_nueva,
                id_documento=id_documento,
                id_version_origen=id_version_origen,
                numero_version=numero_version_nueva,
                numero_revision=0,
                etapa="Borrador",
                nombre_archivo=copia["name"],
                google_doc_id=copia["id"],
                google_doc_url=copia["url"],
                id_aprobacion_responsable=primer_nuevo[
                    "ID_APROBACION_ACTUAL"
                ],
                orden_responsable=entero(
                    primer_nuevo.get("ORDEN"),
                    "ORDEN",
                ),
                motivo_creacion="Reinicio por observaciones de firma",
                comentario=comentario,
                creado_por=usuario,
                fecha_creacion=fecha,
            )

        (
            archivo_observacion_guardado,
            nombre_archivo_observacion,
        ) = marcar_version_observada_firma(
            id_version=id_version_origen,
            fecha=fecha,
            usuario=usuario,
            comentario=comentario,
            archivo_observacion_firma=archivo_observacion_firma,
        )

        actualizar_documento_reinicio_firma(
            id_documento=id_documento,
            numero_version=numero_version_nueva,
            id_version=id_version_nueva,
            copia=copia,
            primer_encargado=primer_nuevo,
            usuario=usuario,
            fecha=fecha,
        )

        advertencias, evento_reinicio = crear_eventos_reinicio_firma(
            id_documento=id_documento,
            id_version_observada=id_version_origen,
            id_version_nueva=id_version_nueva,
            id_aprobacion_nueva=primer_nuevo[
                "ID_APROBACION_ACTUAL"
            ],
            usuario=usuario,
            fecha=fecha,
            comentario=comentario,
            numero_version_nueva=numero_version_nueva,
            nombre_archivo=copia["name"],
            nombre_archivo_observacion=nombre_archivo_observacion,
        )

        notificaciones: list[dict[str, Any]] = []
        if evento_reinicio is not None:
            try:
                documento_actualizado = buscar_documento(id_documento)
                cadena_nueva = buscar_cadena_documento_version(
                    id_documento=id_documento,
                    numero_version=numero_version_nueva,
                )
                aprobador_destino = buscar_aprobacion_actual(
                    texto(documento_actualizado.get("ID_APROBACION_ACTUAL"))
                )
                aprobador_confirma = buscar_integrante_cadena_por_usuario(
                    cadena_nueva,
                    usuario,
                )
                notificaciones = ejecutar_notificaciones_reinicio_firma(
                    documento=documento_actualizado,
                    evento=evento_reinicio,
                    cadena=cadena_nueva,
                    aprobador_destino=aprobador_destino,
                    aprobador_confirma=aprobador_confirma,
                )
                fallidas = [
                    resultado
                    for resultado in notificaciones
                    if not resultado.get("ok")
                ]
                if fallidas:
                    advertencias.append(
                        f"{len(fallidas)} notificación(es) quedaron omitidas "
                        "o con error. Revisa Documento_Notificaciones."
                    )
            except Exception as exc_notificacion:
                traceback.print_exc()
                advertencias.append(
                    "El proceso se reinició correctamente, pero falló el "
                    f"envío de notificaciones internas: {exc_notificacion}"
                )

        return jsonify(
            {
                "ok": True,
                "ya_procesado": False,
                "id_documento": id_documento,
                "estado": "Borrador",
                "estado_firma": "Observado",
                "numero_version": numero_version_nueva,
                "numero_revision": 0,
                "id_version": id_version_nueva,
                "google_doc_id": copia["id"],
                "google_doc_url": copia["url"],
                "nombre_archivo": copia["name"],
                "archivo_observacion_firma": archivo_observacion_guardado,
                "nombre_archivo_observacion_firma": nombre_archivo_observacion,
                "orden_actual": primer_nuevo["ORDEN"],
                "id_aprobacion_actual": primer_nuevo[
                    "ID_APROBACION_ACTUAL"
                ],
                "encargado_actual": primer_nuevo.get("NOMBRE", ""),
                "encargado_email": email_primero,
                "notificaciones": notificaciones,
                "advertencias": advertencias,
            }
        )

    except PermissionError as exc:
        if id_documento:
            registrar_error_transicion(id_documento, str(exc))
        return {"error": str(exc)}, 403

    except (ValueError, LookupError) as exc:
        if id_documento:
            registrar_error_transicion(id_documento, str(exc))
        return {"error": str(exc)}, 400

    except Exception as exc:
        traceback.print_exc()
        if id_documento:
            registrar_error_transicion(id_documento, str(exc))
        return {"error": str(exc)}, 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)
