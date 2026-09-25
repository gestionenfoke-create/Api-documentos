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

# Firma Simple por paquete - fase definitiva de envío/respuesta/rechazo/cierre.


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

TABLA_REVISIONES_EXTERNAS = os.environ.get(
    "TABLA_REVISIONES_EXTERNAS",
    "Documento_Revisiones_Externas",
)

TABLA_REVISION_EXTERNA_DETALLE = os.environ.get(
    "TABLA_REVISION_EXTERNA_DETALLE",
    "Documento_Revision_Externa_Detalle",
)

# Tablas de la fase Notaría
TABLA_NOTARIAS = os.environ.get(
    "TABLA_NOTARIAS",
    "Notarias",
)

TABLA_DOCUMENTOS_PRIME = os.environ.get(
    "TABLA_DOCUMENTOS_PRIME",
    "Documentos_prime",
)

TABLA_ENVIOS_NOTARIA = os.environ.get(
    "TABLA_ENVIOS_NOTARIA",
    "Documento_Envios_Notaria",
)

TABLA_ENVIOS_NOTARIA_DETALLE = os.environ.get(
    "TABLA_ENVIOS_NOTARIA_DETALLE",
    "Documento_Envios_Notaria_Detalle",
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



# -----------------------------------------------------------------------------
# Fase 1: tipo de firma y jerarquía documental
# -----------------------------------------------------------------------------

TIPOS_FIRMA_VALIDOS = {"Simple", "Notarial"}


def normalizar_tipo_firma(valor: Any) -> str:
    """Normaliza el tipo de firma configurado en plantilla/documento."""
    valor_texto = texto(valor).strip()
    if not valor_texto:
        # Compatibilidad con documentos/plantillas existentes creados antes
        # de incorporar TIPO_FIRMA.
        return "Simple"

    normalizado = valor_texto.casefold()
    equivalencias = {
        "simple": "Simple",
        "notarial": "Notarial",
    }
    if normalizado not in equivalencias:
        raise ValueError(
            "TIPO_FIRMA debe ser Simple o Notarial. "
            f"Valor recibido: {valor_texto!r}"
        )
    return equivalencias[normalizado]


def buscar_todos_documentos() -> list[dict[str, Any]]:
    """
    Obtiene los documentos en una sola llamada a AppSheet para construir la
    jerarquía en memoria. Evita una llamada API por cada hijo/nivel.
    """
    selector = f"FILTER({TABLA_DOCUMENTOS}, TRUE)"
    return appsheet_find(TABLA_DOCUMENTOS, selector)


def construir_indice_documentos(
    documentos: list[dict[str, Any]],
) -> dict[str, dict[str, Any]]:
    indice: dict[str, dict[str, Any]] = {}
    for fila in documentos:
        id_documento = texto(fila.get("ID_DOCUMENTO"))
        if not id_documento:
            continue
        if id_documento in indice:
            raise ValueError(
                f"ID_DOCUMENTO duplicado en {TABLA_DOCUMENTOS}: {id_documento}"
            )
        indice[id_documento] = fila
    return indice


def obtener_documento_raiz_desde_indice(
    id_documento: str,
    indice: dict[str, dict[str, Any]],
) -> tuple[dict[str, Any], list[str]]:
    """
    Sube por PADRE hasta encontrar la raíz. Devuelve también el camino desde
    el documento solicitado hacia la raíz. Detecta referencias inexistentes y
    ciclos A -> B -> C -> A.
    """
    actual = texto(id_documento)
    if actual not in indice:
        raise LookupError(
            f"No se encontró ID_DOCUMENTO={actual} en {TABLA_DOCUMENTOS}"
        )

    camino: list[str] = []
    visitados: set[str] = set()

    while True:
        if actual in visitados:
            ciclo = " -> ".join(camino + [actual])
            raise ValueError(
                "Se detectó una relación circular en PADRE: " + ciclo
            )

        visitados.add(actual)
        camino.append(actual)
        fila = indice[actual]
        padre = texto(fila.get("PADRE"))

        if not padre:
            return fila, camino

        if padre == actual:
            raise ValueError(
                f"El documento {actual} no puede tenerse a sí mismo como PADRE"
            )

        if padre not in indice:
            raise ValueError(
                f"El documento {actual} referencia PADRE={padre}, "
                "pero ese documento no existe"
            )

        actual = padre


def construir_jerarquia_desde_raiz(
    id_raiz: str,
    indice: dict[str, dict[str, Any]],
) -> list[tuple[dict[str, Any], int]]:
    """Devuelve raíz + descendientes en orden jerárquico (padre antes que hijos)."""
    hijos_por_padre: dict[str, list[dict[str, Any]]] = {}

    for fila in indice.values():
        id_fila = texto(fila.get("ID_DOCUMENTO"))
        padre = texto(fila.get("PADRE"))
        if not padre:
            continue
        # Inconsistencias ajenas al árbol solicitado no deben impedir validar
        # un paquete correcto. Si afectan al documento solicitado, la subida
        # hacia la raíz las detecta antes de llegar aquí.
        if padre == id_fila or padre not in indice:
            continue
        hijos_por_padre.setdefault(padre, []).append(fila)

    for lista_hijos in hijos_por_padre.values():
        lista_hijos.sort(
            key=lambda fila: (
                texto(fila.get("TITULO")).casefold(),
                texto(fila.get("ID_DOCUMENTO")),
            )
        )

    resultado: list[tuple[dict[str, Any], int]] = []
    visitados: set[str] = set()
    pila_activa: set[str] = set()

    def recorrer(id_actual: str, nivel: int) -> None:
        if id_actual in pila_activa:
            raise ValueError(
                f"Se detectó una relación circular en la jerarquía desde {id_raiz}"
            )
        if id_actual in visitados:
            return

        pila_activa.add(id_actual)
        visitados.add(id_actual)
        fila_actual = indice[id_actual]
        resultado.append((fila_actual, nivel))

        for hijo in hijos_por_padre.get(id_actual, []):
            recorrer(texto(hijo.get("ID_DOCUMENTO")), nivel + 1)

        pila_activa.remove(id_actual)

    recorrer(id_raiz, 0)
    return resultado


ESTADO_LISTO_FIRMA = "Listo para firma"
ESTADO_LISTO_REVISION_EXTERNA = "Listo para revisión externa"
ESTADOS_APROBACION_INTERNA_COMPLETA = {
    ESTADO_LISTO_FIRMA,
    ESTADO_LISTO_REVISION_EXTERNA,
}


def obtener_tipo_firma_efectivo_documento(id_documento: str) -> str:
    """
    Devuelve el TIPO_FIRMA que gobierna al documento. En una jerarquía,
    siempre manda el TIPO_FIRMA de la raíz del paquete.
    """
    documentos = buscar_todos_documentos()
    indice = construir_indice_documentos(documentos)
    raiz, _ = obtener_documento_raiz_desde_indice(id_documento, indice)
    return normalizar_tipo_firma(raiz.get("TIPO_FIRMA"))


def configuracion_salida_aprobacion(tipo_firma: str) -> dict[str, str]:
    """Define la salida de la aprobación interna según el tipo de firma."""
    tipo = normalizar_tipo_firma(tipo_firma)
    if tipo == "Notarial":
        return {
            "tipo_firma": "Notarial",
            "estado_documento": ESTADO_LISTO_REVISION_EXTERNA,
            "etapa_version": "Para revisión externa",
            "sufijo_archivo": "PARA_REVISION_EXTERNA",
            "motivo_creacion": "Preparación para revisión externa",
            "tipo_evento_preparado": "Preparado para revisión externa",
            "movimiento_responsable": "Documento listo para gestionar revisión externa",
            "movimiento_informativo": "Documento listo para revisión externa",
        }
    return {
        "tipo_firma": "Simple",
        "estado_documento": ESTADO_LISTO_FIRMA,
        "etapa_version": "Para firma",
        "sufijo_archivo": "PARA_FIRMA",
        "motivo_creacion": "Preparación para firma",
        "tipo_evento_preparado": "Preparado para firma",
        "movimiento_responsable": "Documento listo para gestionar firma",
        "movimiento_informativo": "Documento listo para firma",
    }


ROL_RESPONSABLE_FIRMAS = "responsable de firmas"
ESTADO_GESTION_FIRMA = "Gestión de firma"


def es_responsable_firmas(fila: dict[str, Any] | None) -> bool:
    """Indica si una fila de la cadena corresponde al responsable operativo de firma."""
    if not fila:
        return False
    return texto(fila.get("ROL_FLUJO")).strip().lower() == ROL_RESPONSABLE_FIRMAS


def obtener_responsable_firmas_cadena(
    cadena: list[dict[str, Any]],
    *,
    contexto: str = "cadena vigente",
) -> dict[str, Any]:
    """Exige exactamente un Responsable de firmas dentro de la cadena indicada."""
    responsables = [fila for fila in cadena if es_responsable_firmas(fila)]
    if not responsables:
        raise ValueError(
            f"{contexto}: no existe un integrante con ROL_FLUJO='Responsable de firmas'"
        )
    if len(responsables) > 1:
        raise ValueError(
            f"{contexto}: existen {len(responsables)} integrantes con "
            "ROL_FLUJO='Responsable de firmas'; debe existir exactamente uno"
        )
    responsable = responsables[0]
    if not texto(responsable.get("APROBADOR")):
        raise ValueError("El Responsable de firmas no tiene APROBADOR/correo")
    return responsable


def obtener_primer_responsable_aprobacion(
    cadena: list[dict[str, Any]],
    *,
    contexto: str = "cadena vigente",
) -> dict[str, Any]:
    """Devuelve el primer integrante que participa realmente de la aprobación interna."""
    aprobadores = [fila for fila in cadena if not es_responsable_firmas(fila)]
    if not aprobadores:
        raise ValueError(
            f"{contexto}: no existe ningún integrante de aprobación distinto del "
            "Responsable de firmas"
        )
    return min(aprobadores, key=lambda fila: entero(fila.get("ORDEN"), "ORDEN"))


def validar_responsable_firmas_actual(documento: dict[str, Any]) -> tuple[bool, str]:
    """Valida que el responsable actual de un Listo para firma sea el rol de firmas."""
    id_actual = texto(documento.get("ID_APROBACION_ACTUAL"))
    if not id_actual:
        return False, "No tiene responsable actual de firma"
    try:
        fila = buscar_aprobacion_actual(id_actual)
    except Exception as exc:
        return False, f"No se pudo resolver el responsable actual de firma: {exc}"
    if not es_responsable_firmas(fila):
        return False, (
            "El responsable actual no tiene ROL_FLUJO='Responsable de firmas'"
        )
    if texto(fila.get("ESTADO")) != ESTADO_GESTION_FIRMA:
        return False, (
            f"El Responsable de firmas está en estado {texto(fila.get('ESTADO'))!r}; "
            f"se requiere {ESTADO_GESTION_FIRMA!r}"
        )
    if not es_verdadero(fila.get("CADENA_ACTIVA")):
        return False, "La fila del Responsable de firmas no está activa"
    return True, ""


def evaluar_documento_para_paquete(
    documento: dict[str, Any],
    *,
    nivel: int,
    id_raiz: str,
) -> dict[str, Any]:
    """
    Evalúa la preparación interna usando el flujo vigente de Fase 1.

    Durante esta fase tanto Simple como Notarial siguen llegando a
    "Listo para firma" al finalizar la cadena. En fases posteriores la regla
    se ampliará a "Listo para revisión externa" para la ruta Notarial.
    """
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    estado = texto(documento.get("ESTADO"))
    id_version = texto(documento.get("ID_VERSION_ACTUAL"))
    google_doc_id = texto(documento.get("GOOGLE_DOC_ID"))
    pdf_id = texto(documento.get("PDF_PARA_FIRMA_ID"))
    id_aprobacion_actual = texto(documento.get("ID_APROBACION_ACTUAL"))

    motivos: list[str] = []

    if estado != "Listo para firma":
        motivos.append(
            f"Estado {estado or '<vacío>'}; se requiere 'Listo para firma'"
        )
    responsable_firma_valido = False
    if estado == "Listo para firma":
        responsable_firma_valido, motivo_responsable = validar_responsable_firmas_actual(
            documento
        )
        if not responsable_firma_valido:
            motivos.append(motivo_responsable)
    elif id_aprobacion_actual:
        motivos.append("La cadena interna todavía tiene un responsable de aprobación")
    if not id_version:
        motivos.append("No tiene ID_VERSION_ACTUAL")
    if not google_doc_id:
        motivos.append("No tiene GOOGLE_DOC_ID")
    if not pdf_id:
        motivos.append("No tiene PDF_PARA_FIRMA_ID")

    proyecto = texto(documento.get("ID_PROYECTO"))
    padre = texto(documento.get("PADRE"))

    return {
        "id_documento": id_documento,
        "titulo": texto(documento.get("TITULO")),
        "nivel": nivel,
        "es_raiz": id_documento == id_raiz,
        "padre": padre,
        "id_proyecto": proyecto,
        "tipo_firma_documento": normalizar_tipo_firma(
            documento.get("TIPO_FIRMA")
        ),
        "estado": estado,
        "estado_firma": texto(documento.get("ESTADO_FIRMA")),
        "numero_version": texto(documento.get("VERSION_ACTUAL")),
        "numero_revision": texto(documento.get("REVISION_ACTUAL")),
        "id_version_actual": id_version,
        "google_doc_id": google_doc_id,
        "pdf_para_firma_id": pdf_id,
        "cadena_aprobacion_concluida": responsable_firma_valido,
        "responsable_firmas_asignado": responsable_firma_valido,
        "listo_para_paquete": not motivos,
        "motivos": motivos,
    }


def validar_paquete_documental_datos(id_documento: str) -> dict[str, Any]:
    """Construye y valida el paquete controlado por el documento raíz."""
    documentos = buscar_todos_documentos()
    indice = construir_indice_documentos(documentos)
    raiz, camino_a_raiz = obtener_documento_raiz_desde_indice(
        id_documento,
        indice,
    )
    id_raiz = texto(raiz.get("ID_DOCUMENTO"))
    jerarquia = construir_jerarquia_desde_raiz(id_raiz, indice)

    proyecto_raiz = texto(raiz.get("ID_PROYECTO"))
    evaluados: list[dict[str, Any]] = []
    errores_jerarquia: list[str] = []

    for documento, nivel in jerarquia:
        evaluado = evaluar_documento_para_paquete(
            documento,
            nivel=nivel,
            id_raiz=id_raiz,
        )

        proyecto_hijo = evaluado["id_proyecto"]
        if (
            proyecto_raiz
            and proyecto_hijo
            and proyecto_hijo != proyecto_raiz
        ):
            mensaje = (
                f"{evaluado['titulo'] or evaluado['id_documento']} pertenece "
                f"al proyecto {proyecto_hijo}, distinto del proyecto raíz "
                f"{proyecto_raiz}"
            )
            evaluado["motivos"].append(mensaje)
            evaluado["listo_para_paquete"] = False
            errores_jerarquia.append(mensaje)

        evaluados.append(evaluado)

    pendientes = [
        fila for fila in evaluados if not fila["listo_para_paquete"]
    ]
    solicitado_es_raiz = texto(id_documento) == id_raiz
    tipo_firma_paquete = normalizar_tipo_firma(raiz.get("TIPO_FIRMA"))

    return {
        "id_documento_solicitado": texto(id_documento),
        "id_documento_raiz": id_raiz,
        "titulo_raiz": texto(raiz.get("TITULO")),
        "tipo_firma_paquete": tipo_firma_paquete,
        "solicitado_es_raiz": solicitado_es_raiz,
        "camino_a_raiz": camino_a_raiz,
        "cantidad_documentos": len(evaluados),
        "paquete_listo": not pendientes and not errores_jerarquia,
        "puede_iniciar_envio_externo": (
            solicitado_es_raiz and not pendientes and not errores_jerarquia
        ),
        "cantidad_pendientes": len(pendientes),
        "documentos_pendientes": [
            {
                "id_documento": fila["id_documento"],
                "titulo": fila["titulo"],
                "estado": fila["estado"],
                "motivos": fila["motivos"],
            }
            for fila in pendientes
        ],
        "errores_jerarquia": errores_jerarquia,
        "documentos": evaluados,
    }


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
    primer_base = obtener_primer_responsable_aprobacion(
        cadena_plantilla,
        contexto=f"Plantilla {id_plantilla}",
    )
    obtener_responsable_firmas_cadena(
        cadena_plantilla,
        contexto=f"Plantilla {id_plantilla}",
    )
    orden_primero = entero(primer_base.get("ORDEN"), "ORDEN")
    primer_actual: dict[str, Any] | None = None

    for fila_base in cadena_plantilla:
        id_aprobacion_actual = nuevo_id()
        orden = entero(fila_base.get("ORDEN"), "ORDEN")
        aprobador = texto(fila_base.get("APROBADOR"))
        nombre = texto(fila_base.get("NOMBRE"))
        rol_flujo = texto(fila_base.get("ROL_FLUJO"))

        if not aprobador:
            raise ValueError(
                f"El responsable de orden {orden} no tiene APROBADOR"
            )

        es_primero = orden == orden_primero and not es_responsable_firmas(fila_base)
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
            "ESTADO": "En elaboración" if es_primero else "Pendiente",
            "CADENA_ACTIVA": True,
        }

        if es_primero:
            fila_actual.update(
                {
                    "ID_VERSION_TRABAJADA": id_version,
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
    tipo_firma: str,
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
        # Snapshot: el documento conserva el TIPO_FIRMA que tenía la plantilla
        # al momento de su creación, aunque la plantilla cambie más adelante.
        "TIPO_FIRMA": tipo_firma,
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


@app.route("/validar-paquete-documental", methods=["POST"])
def validar_paquete_documental():
    """
    Fase 1: diagnóstico de jerarquía y preparación del paquete documental.

    No modifica datos. Puede recibir id_documento o id_documento_raiz. Si se
    envía un hijo, identifica la raíz y devuelve el paquete completo, pero
    puede_iniciar_envio_externo será False porque solo la raíz administra la
    salida externa.
    """
    try:
        validar_configuracion()
        validar_token()

        data = request.get_json(silent=True) or {}
        id_documento = texto(
            data.get("id_documento_raiz") or data.get("id_documento")
        )
        if not id_documento:
            return {"error": "Falta id_documento o id_documento_raiz"}, 400

        resultado = validar_paquete_documental_datos(id_documento)
        return jsonify({"ok": True, **resultado})

    except PermissionError as exc:
        return {"error": str(exc)}, 403
    except (ValueError, LookupError) as exc:
        return {"error": str(exc)}, 400
    except Exception as exc:
        traceback.print_exc()
        return {"error": str(exc)}, 500


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
                "ACCION_SOLICITADA": "",
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
    # Evita NameError en manejadores de error heredados si la creación falla
    # antes de completar el documento.
    documento_cerrado = False

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
        tipo_firma = normalizar_tipo_firma(plantilla.get("TIPO_FIRMA"))

        template_id = texto(plantilla.get("GOOGLE_DOC_TEMPLATE_ID"))
        folder_id = texto(plantilla.get("CARPETA_DESTINO_ID"))

        if not template_id:
            raise ValueError("La plantilla no tiene GOOGLE_DOC_TEMPLATE_ID")

        if not folder_id:
            raise ValueError("La plantilla no tiene CARPETA_DESTINO_ID")

        # La plantilla debe tener exactamente un responsable operativo de firmas,
        # pero ese rol no participa como etapa de aprobación.
        obtener_responsable_firmas_cadena(
            cadena_plantilla,
            contexto=f"Plantilla {id_plantilla}",
        )
        primer_base = obtener_primer_responsable_aprobacion(
            cadena_plantilla,
            contexto=f"Plantilla {id_plantilla}",
        )
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
            tipo_firma=tipo_firma,
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
                "tipo_firma": tipo_firma,
                "padre": texto(documento.get("PADRE")),
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
            and not es_responsable_firmas(fila)
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


def cerrar_cadena_para_gestion_externa(
    cadena_actual: list[dict[str, Any]],
    aprobacion_actual: dict[str, Any],
    responsable_firmas: dict[str, Any],
    id_version_salida_externa: str,
    comentario: str,
    fecha: str,
) -> None:
    """Cierra la aprobación interna y deja activo al Responsable de firmas."""
    filas: list[dict[str, Any]] = []
    id_actual = texto(aprobacion_actual.get("ID_APROBACION_ACTUAL"))
    id_responsable_firmas = texto(
        responsable_firmas.get("ID_APROBACION_ACTUAL")
    )

    for fila in cadena_actual:
        id_fila = texto(fila.get("ID_APROBACION_ACTUAL"))
        cambios: dict[str, Any] = {
            "ID_APROBACION_ACTUAL": id_fila,
            "CADENA_ACTIVA": id_fila == id_responsable_firmas,
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
        elif id_fila == id_responsable_firmas:
            cambios.update(
                {
                    "ESTADO": ESTADO_GESTION_FIRMA,
                    "RESULTADO": "",
                    "COMENTARIO": "",
                    "ID_VERSION_TRABAJADA": id_version_salida_externa,
                    "FECHA_INICIO": fecha,
                    "FECHA_RESPUESTA": "",
                }
            )

        filas.append(cambios)

    appsheet_action(TABLA_APROBADORES_ACTUAL, "Edit", filas)


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


def actualizar_documento_listo_salida_externa(
    *,
    id_documento: str,
    numero_version: int,
    numero_revision: int,
    id_version: str,
    copia: dict[str, str],
    pdf: dict[str, str],
    responsable_firmas: dict[str, Any],
    tipo_firma: str,
    usuario: str,
    fecha: str,
) -> None:
    config = configuracion_salida_aprobacion(tipo_firma)
    cambios: dict[str, Any] = {
        "ID_DOCUMENTO": id_documento,
        "ESTADO": config["estado_documento"],
        "VERSION_ACTUAL": numero_version,
        "REVISION_ACTUAL": numero_revision,
        "ID_VERSION_ACTUAL": id_version,
        "GOOGLE_DOC_ID": copia["id"],
        "GOOGLE_DOC_URL": copia["url"],
        "ORDEN_ACTUAL": responsable_firmas["ORDEN"],
        "ID_APROBACION_ACTUAL": responsable_firmas["ID_APROBACION_ACTUAL"],
        "ENCARGADO_ACTUAL_NOMBRE": responsable_firmas.get("NOMBRE", ""),
        "ENCARGADO_ACTUAL_EMAIL": responsable_firmas.get("APROBADOR", ""),
        "ULTIMO_ENVIADO_POR": usuario,
        "FECHA_ULTIMO_ENVIO": fecha,
        "FECHA_ULTIMA_ACTUALIZACION": fecha,
        "OBSERVACION_ACTUAL": "",
        "ACCION_SOLICITADA": "",
    }

    if config["tipo_firma"] == "Simple":
        cambios.update(
            {
                "PDF_PARA_FIRMA_ID": pdf["id"],
                "PDF_PARA_FIRMA_URL": pdf["url"],
                "ESTADO_FIRMA": "No iniciado",
            }
        )
    else:
        # En la ruta Notarial el PDF se conserva en Documento_Versiones.
        # PDF_PARA_FIRMA pertenece al flujo Simple y no debe representar
        # la revisión externa notarial.
        cambios.update(
            {
                "PDF_PARA_FIRMA_ID": "",
                "PDF_PARA_FIRMA_URL": "",
            }
        )

    appsheet_action(TABLA_DOCUMENTOS, "Edit", [cambios])


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
    estado_final: str | None = None,
) -> dict[str, Any]:
    if orden_siguiente is None:
        estado_nuevo = estado_final or ESTADO_LISTO_FIRMA
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
    tipo_firma: str = "Simple",
) -> None:
    config = configuracion_salida_aprobacion(tipo_firma)
    appsheet_action(
        TABLA_EVENTOS,
        "Add",
        [
            {
                "ID_EVENTO": nuevo_id(),
                "ID_DOCUMENTO": id_documento,
                "ID_VERSION": id_version,
                "ID_APROBACION_ACTUAL": id_aprobacion_actual,
                "TIPO_EVENTO": config["tipo_evento_preparado"],
                "ESTADO_ANTERIOR": "En revisión",
                "ESTADO_NUEVO": config["estado_documento"],
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

        if estado_documento in ESTADOS_APROBACION_INTERNA_COMPLETA:
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
                    f"El documento ya estaba {estado_documento}, pero no se "
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
            and not es_responsable_firmas(fila)
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

        # Último aprobador: la aprobación interna termina. El TIPO_FIRMA
        # efectivo lo gobierna la raíz de la jerarquía, no el hijo aislado.
        tipo_firma_efectivo = obtener_tipo_firma_efectivo_documento(id_documento)
        config_salida = configuracion_salida_aprobacion(tipo_firma_efectivo)
        estado_salida = config_salida["estado_documento"]

        # El Responsable de firmas queda como responsable operativo tanto para
        # Firma Simple como para la futura gestión de revisión externa Notarial.
        responsable_firmas = obtener_responsable_firmas_cadena(
            cadena_actual,
            contexto=f"Documento {id_documento} versión {numero_version}",
        )

        sufijo_salida = config_salida["sufijo_archivo"]
        nombre_archivo = limpiar_nombre_archivo(
            f"{titulo}_V{numero_version:02d}_{sufijo_salida}"
        )
        nombre_pdf = limpiar_nombre_archivo(
            f"{titulo}_V{numero_version:02d}_{sufijo_salida}.pdf"
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
                etapa=config_salida["etapa_version"],
                nombre_archivo=copia["name"],
                google_doc_id=copia["id"],
                google_doc_url=copia["url"],
                pdf_version_id=pdf["id"],
                pdf_version_url=pdf["url"],
                id_aprobacion_responsable=responsable_firmas[
                    "ID_APROBACION_ACTUAL"
                ],
                orden_responsable=entero(
                    responsable_firmas.get("ORDEN"),
                    "ORDEN",
                ),
                motivo_creacion=config_salida["motivo_creacion"],
                comentario=comentario,
                creado_por=usuario,
                fecha_creacion=fecha,
            )

        actualizar_estado_version(
            id_version=id_version_actual,
            estado_version="Aprobada",
            fecha_cierre=fecha,
        )

        cerrar_cadena_para_gestion_externa(
            cadena_actual=cadena_actual,
            aprobacion_actual=aprobacion_actual,
            responsable_firmas=responsable_firmas,
            id_version_salida_externa=id_version_nueva,
            comentario=comentario,
            fecha=fecha,
        )

        actualizar_documento_listo_salida_externa(
            id_documento=id_documento,
            numero_version=numero_version,
            numero_revision=numero_revision_nueva,
            id_version=id_version_nueva,
            copia=copia,
            pdf=pdf,
            responsable_firmas=responsable_firmas,
            tipo_firma=tipo_firma_efectivo,
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
                estado_final=estado_salida,
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
                id_aprobacion_actual=responsable_firmas[
                    "ID_APROBACION_ACTUAL"
                ],
                usuario=usuario,
                fecha=fecha,
                nombre_archivo=copia["name"],
                nombre_pdf=pdf["name"],
                tipo_firma=tipo_firma_efectivo,
            )
        except Exception as exc_evento_firma:
            traceback.print_exc()
            advertencias.append(
                f"El documento quedó {estado_salida}, pero no se pudo crear "
                f"el evento técnico {config_salida['tipo_evento_preparado']}: "
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
                        aprobador_actual=responsable_firmas,
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
                "estado": estado_salida,
                "tipo_firma_paquete": tipo_firma_efectivo,
                "etapa_version": config_salida["etapa_version"],
                "numero_version": numero_version,
                "numero_revision": numero_revision_nueva,
                "id_version": id_version_nueva,
                "google_doc_id": copia["id"],
                "google_doc_url": copia["url"],
                "nombre_archivo": copia["name"],
                "pdf_id": pdf["id"],
                "pdf_url": pdf["url"],
                "pdf_nombre": pdf["name"],
                "responsable_firmas_nombre": responsable_firmas.get("NOMBRE", ""),
                "responsable_firmas_email": responsable_firmas.get("APROBADOR", ""),
                "id_aprobacion_actual": responsable_firmas.get(
                    "ID_APROBACION_ACTUAL", ""
                ),
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

        if es_responsable_firmas(fila_anterior):
            # El responsable de firmas es operativo, no una aprobación heredable.
            # En cada nueva versión queda pendiente hasta que finalice nuevamente
            # toda la cadena interna.
            fila_nueva["ESTADO"] = "Pendiente"
        elif indice < indice_destino:
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


def cerrar_gestion_firma_anterior_por_observacion(
    *,
    cadena_anterior: list[dict[str, Any]],
    comentario: str,
    fecha: str,
) -> None:
    """Desactiva la gestión de firma de la versión observada antes de reiniciar."""
    filas: list[dict[str, Any]] = []
    for fila in cadena_anterior:
        if not es_verdadero(fila.get("CADENA_ACTIVA")):
            continue
        cambios: dict[str, Any] = {
            "ID_APROBACION_ACTUAL": texto(fila.get("ID_APROBACION_ACTUAL")),
            "CADENA_ACTIVA": False,
        }
        if es_responsable_firmas(fila):
            cambios.update(
                {
                    "ESTADO": "Cerrado",
                    "RESULTADO": "Observado en firma",
                    "COMENTARIO": comentario,
                    "FECHA_RESPUESTA": fecha,
                }
            )
        filas.append(cambios)
    if filas:
        appsheet_action(TABLA_APROBADORES_ACTUAL, "Edit", filas)


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


def construir_email_firma_paquete_externo(
    *,
    documento: dict[str, Any],
    plantilla: dict[str, Any],
    usuario: str,
    mensaje_adicional: str,
    fecha_envio: str,
    documentos_paquete: list[dict[str, Any]],
) -> tuple[str, str, str]:
    """
    Construye el correo de firma simple para uno o varios documentos.

    Para un paquete de un solo documento conserva exactamente el correo
    histórico. Cuando hay jerarquía, el documento raíz sigue siendo el
    documento principal y se agrega una relación explícita de todos los
    documentos incluidos en el mismo envío.
    """
    asunto, cuerpo_texto, cuerpo_html = construir_email_firma_externo(
        documento=documento,
        plantilla=plantilla,
        usuario=usuario,
        mensaje_adicional=mensaje_adicional,
        fecha_envio=fecha_envio,
    )

    if len(documentos_paquete) <= 1:
        return asunto, cuerpo_texto, cuerpo_html

    cantidad = len(documentos_paquete)
    titulo_raiz = (
        texto(documento.get("TITULO"))
        or f"Documento {texto(documento.get('ID_DOCUMENTO'))}"
    )

    lineas_documentos: list[str] = []
    for indice, fila in enumerate(documentos_paquete, start=1):
        titulo = texto(fila.get("titulo")) or texto(fila.get("id_documento"))
        version = texto(fila.get("numero_version"))
        revision = texto(fila.get("numero_revision"))
        detalle = titulo
        if version:
            detalle += f" — V{version}"
            if revision:
                detalle += f" / Rev. {revision}"
        lineas_documentos.append(f"{indice}. {detalle}")

    listado_texto = "\n".join(lineas_documentos)
    bloque_texto = (
        "\n\nDOCUMENTOS INCLUIDOS EN EL PAQUETE\n"
        f"{listado_texto}\n"
    )

    cuerpo_texto = cuerpo_texto.replace(
        f'Adjuntamos el documento "{titulo_raiz}" para su revisión y firma.',
        (
            f'Adjuntamos el documento principal "{titulo_raiz}" junto con '
            f"{cantidad - 1} documento(s) relacionado(s), para su revisión "
            "y firma en un único envío."
        ),
    )
    cuerpo_texto = cuerpo_texto.replace(
        "\n\n¿QUÉ DEBE HACER?",
        bloque_texto + "\n¿QUÉ DEBE HACER?",
        1,
    )
    cuerpo_texto = cuerpo_texto.replace(
        "2. Firmar el documento utilizando el archivo PDF.",
        "2. Firmar los documentos PDF que correspondan.",
    )
    cuerpo_texto = cuerpo_texto.replace(
        "3. Responder este mismo correo adjuntando el PDF firmado.",
        "3. Responder este mismo correo adjuntando los PDF firmados que correspondan.",
    )

    filas_html = []
    for indice, fila in enumerate(documentos_paquete, start=1):
        titulo = html.escape(
            texto(fila.get("titulo")) or texto(fila.get("id_documento"))
        )
        version = html.escape(texto(fila.get("numero_version")))
        revision = html.escape(texto(fila.get("numero_revision")))
        detalle_version = ""
        if version:
            detalle_version = f"V{version}"
            if revision:
                detalle_version += f" / Rev. {revision}"
        detalle_html = (
            f'<span style="color:#6b7280; font-size:12px;">'
            f'{html.escape(detalle_version)}</span>'
            if detalle_version
            else ""
        )
        filas_html.append(
            f"""
            <tr>
              <td style="
                  padding:10px 12px;
                  width:34px;
                  border-top:1px solid #e5e7eb;
                  color:#4338ca;
                  font-weight:700;
                  vertical-align:top;
              ">{indice}.</td>
              <td style="
                  padding:10px 12px;
                  border-top:1px solid #e5e7eb;
                  color:#111827;
                  font-size:14px;
                  line-height:1.45;
              ">
                <strong>{titulo}</strong><br>
                {detalle_html}
              </td>
            </tr>
            """
        )

    bloque_html = f"""
                    <div style="
                        margin-top:22px;
                        border:1px solid #e5e7eb;
                        border-radius:8px;
                        overflow:hidden;
                    ">
                      <div style="
                          padding:12px 16px;
                          background:#eef2ff;
                          color:#3730a3;
                          font-size:14px;
                          font-weight:700;
                      ">Documentos incluidos en el paquete ({cantidad})</div>
                      <table role="presentation" width="100%" cellspacing="0" cellpadding="0"
                             style="width:100%; border-collapse:collapse;">
                        {''.join(filas_html)}
                      </table>
                    </div>
    """

    cuerpo_html = cuerpo_html.replace(
        "Documento adjunto para revisión y firma",
        "Paquete documental adjunto para revisión y firma",
    )
    cuerpo_html = cuerpo_html.replace(
        "Se adjuntan el PDF para firma y una copia editable en formato Microsoft Word.",
        (
            f"Se adjuntan los archivos PDF y Microsoft Word correspondientes "
            f"a {cantidad} documentos relacionados."
        ),
    )
    cuerpo_html = cuerpo_html.replace(
        ">Resumen del documento</div>",
        ">Resumen del documento principal</div>",
    )
    marcador_acciones = """                    <div style="
                        margin-top:22px;
                        padding:17px 18px;"""
    if marcador_acciones in cuerpo_html:
        cuerpo_html = cuerpo_html.replace(
            marcador_acciones,
            bloque_html + "\n" + marcador_acciones,
            1,
        )

    cuerpo_html = cuerpo_html.replace(
        "Firmar el documento utilizando el archivo PDF.",
        "Firmar los documentos PDF que correspondan.",
    )
    cuerpo_html = cuerpo_html.replace(
        "Responder este mismo correo adjuntando el PDF firmado.",
        "Responder este mismo correo adjuntando los PDF firmados que correspondan.",
    )
    cuerpo_html = cuerpo_html.replace(
        "Puede responder directamente a este mensaje con el documento firmado.",
        "Puede responder directamente a este mensaje con los documentos firmados.",
    )

    return asunto, cuerpo_texto, cuerpo_html


@medir_operacion("gmail.enviar_firma_externa")
def enviar_email_con_adjuntos(
    gmail_service: Any,
    destinatarios: list[str],
    asunto: str,
    cuerpo: str,
    adjuntos: list[dict[str, Any]],
    reply_to: str = "",
    cuerpo_html: str = "",
) -> dict[str, str]:
    """
    Envía un correo externo con una cantidad variable de adjuntos.

    Cada adjunto debe contener:
    - contenido: bytes
    - nombre: nombre del archivo
    - maintype: application
    - subtype: pdf o el MIME subtype de DOCX
    """
    if not adjuntos:
        raise ValueError("El correo de firma no tiene archivos adjuntos")

    mensaje = EmailMessage()
    mensaje["To"] = ", ".join(destinatarios)
    mensaje["From"] = GMAIL_SENDER_EMAIL
    mensaje["Subject"] = asunto

    if reply_to and _EMAIL_RE.fullmatch(reply_to.lower()):
        mensaje["Reply-To"] = reply_to

    mensaje.set_content(cuerpo)
    if cuerpo_html:
        mensaje.add_alternative(cuerpo_html, subtype="html")

    nombres_vistos: set[str] = set()
    for adjunto in adjuntos:
        contenido = adjunto.get("contenido")
        nombre = texto(adjunto.get("nombre"))
        maintype = texto(adjunto.get("maintype")) or "application"
        subtype = texto(adjunto.get("subtype"))

        if not isinstance(contenido, bytes) or not contenido:
            raise ValueError(
                f"El adjunto {nombre or '<sin nombre>'} no contiene bytes válidos"
            )
        if not nombre:
            raise ValueError("Se recibió un adjunto sin nombre")
        if not subtype:
            raise ValueError(f"El adjunto {nombre} no tiene MIME subtype")

        # Evita nombres repetidos dentro del mismo mensaje.
        nombre_final = nombre
        contador = 2
        base, extension = os.path.splitext(nombre)
        while nombre_final.casefold() in nombres_vistos:
            nombre_final = f"{base}_{contador}{extension}"
            contador += 1
        nombres_vistos.add(nombre_final.casefold())

        mensaje.add_attachment(
            contenido,
            maintype=maintype,
            subtype=subtype,
            filename=nombre_final,
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
        "label_ids": [texto(x) for x in (respuesta.get("labelIds") or []) if texto(x)],
    }


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
    """Compatibilidad con el envío histórico de un único documento."""
    return enviar_email_con_adjuntos(
        gmail_service=gmail_service,
        destinatarios=destinatarios,
        asunto=asunto,
        cuerpo=cuerpo,
        adjuntos=[
            {
                "contenido": pdf_bytes,
                "nombre": pdf_nombre,
                "maintype": "application",
                "subtype": "pdf",
            },
            {
                "contenido": docx_bytes,
                "nombre": docx_nombre,
                "maintype": "application",
                "subtype": (
                    "vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
            },
        ],
        reply_to=reply_to,
        cuerpo_html=cuerpo_html,
    )

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
    - Responsable de firmas: Acción requerida;
    - quien aprobó: Confirmación;
    - resto de la cadena: Informativa.
    """
    comentario_evento = (
        texto(evento.get("COMENTARIO"))
        or "La revisión fue aprobada."
    )

    especificaciones: list[dict[str, Any]] = []
    estado_documento = texto(documento.get("ESTADO"))
    es_salida_notarial = (
        ultimo_aprobador
        and estado_documento == ESTADO_LISTO_REVISION_EXTERNA
    )
    movimiento_final_responsable = (
        "Documento listo para gestionar revisión externa"
        if es_salida_notarial
        else "Documento listo para gestionar firma"
    )
    movimiento_final_informativo = (
        "Documento listo para revisión externa"
        if es_salida_notarial
        else "Documento listo para firma"
    )

    if aprobador_actual is not None:
        especificaciones.append(
            {
                "aprobador": aprobador_actual,
                "tipo_notificacion": "Acción requerida",
                "movimiento": (
                    movimiento_final_responsable
                    if ultimo_aprobador
                    else "Documento asignado para continuar la revisión"
                ),
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

    id_actual_notificado = (
        texto(aprobador_actual.get("ID_APROBACION_ACTUAL"))
        if aprobador_actual is not None
        else ""
    )
    for integrante in cadena:
        if (
            id_actual_notificado
            and texto(integrante.get("ID_APROBACION_ACTUAL")) == id_actual_notificado
        ):
            continue
        especificaciones.append(
            {
                "aprobador": integrante,
                "tipo_notificacion": "Informativa",
                "movimiento": (
                    movimiento_final_informativo
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
    estado_documento = texto(documento.get("ESTADO"))
    ultimo_aprobador = estado_documento in ESTADOS_APROBACION_INTERNA_COMPLETA

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
    id_aprobacion_actual = texto(documento.get("ID_APROBACION_ACTUAL"))
    if not id_aprobacion_actual:
        advertencias.append(
            "El documento no tiene encargado actual para reanudar las notificaciones."
        )
        return [], advertencias
    aprobador_actual = buscar_aprobacion_actual(id_aprobacion_actual)
    if ultimo_aprobador and not es_responsable_firmas(aprobador_actual):
        advertencias.append(
            f"El documento está {estado_documento}, pero su encargado actual no es "
            "Responsable de firmas."
        )
        return [], advertencias

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
            estado_final=(estado_documento if ultimo_aprobador else None),
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



# -----------------------------------------------------------------------------
# Fase definitiva Firma Simple: paquete externo y respuestas por documento
# -----------------------------------------------------------------------------

RESULTADOS_FIRMA_EXTERNA_VALIDOS = {
    "Pendiente",
    "Conforme",
    "Observado",
    "Firmado",
}
ESTADOS_PAQUETE_FIRMA_ACTIVOS = {
    "Enviado",
    "Con observaciones",
    "Esperando firma",
    "Listo para cierre",
}


def normalizar_resultado_firma_externa(valor: Any) -> str:
    valor_texto = texto(valor).strip()
    if not valor_texto:
        return ""
    equivalencias = {
        "pendiente": "Pendiente",
        "conforme": "Conforme",
        "observado": "Observado",
        "firmado": "Firmado",
    }
    clave = valor_texto.casefold()
    if clave not in equivalencias:
        raise ValueError(
            "RESULTADO_FIRMA_EXTERNA contiene un valor no reconocido: "
            f"{valor_texto!r}"
        )
    return equivalencias[clave]


def obtener_contexto_jerarquia_documental(
    id_documento: str,
) -> dict[str, Any]:
    """Devuelve raíz + jerarquía actual sin exigir que todos estén Listo para firma."""
    documentos = buscar_todos_documentos()
    indice = construir_indice_documentos(documentos)
    raiz, camino = obtener_documento_raiz_desde_indice(id_documento, indice)
    id_raiz = texto(raiz.get("ID_DOCUMENTO"))
    jerarquia = construir_jerarquia_desde_raiz(id_raiz, indice)

    proyecto_raiz = texto(raiz.get("ID_PROYECTO"))
    integrantes: list[dict[str, Any]] = []
    for documento, nivel in jerarquia:
        proyecto = texto(documento.get("ID_PROYECTO"))
        if proyecto_raiz and proyecto and proyecto != proyecto_raiz:
            raise ValueError(
                f"{texto(documento.get('TITULO')) or texto(documento.get('ID_DOCUMENTO'))} "
                f"pertenece al proyecto {proyecto}, distinto del proyecto raíz "
                f"{proyecto_raiz}"
            )
        integrantes.append(
            {
                "documento": documento,
                "id_documento": texto(documento.get("ID_DOCUMENTO")),
                "titulo": texto(documento.get("TITULO"))
                or texto(documento.get("ID_DOCUMENTO")),
                "nivel": nivel,
                "es_raiz": texto(documento.get("ID_DOCUMENTO")) == id_raiz,
            }
        )

    return {
        "raiz": raiz,
        "id_documento_raiz": id_raiz,
        "titulo_raiz": texto(raiz.get("TITULO")),
        "camino_a_raiz": camino,
        "solicitado_es_raiz": texto(id_documento) == id_raiz,
        "integrantes": integrantes,
        "tipo_firma_paquete": normalizar_tipo_firma(raiz.get("TIPO_FIRMA")),
    }


def buscar_integrantes_paquete_activo(id_documento_raiz: str) -> list[dict[str, Any]]:
    selector = (
        f"FILTER({TABLA_DOCUMENTOS}, "
        f"[ID_DOCUMENTO_RAIZ_FIRMA] = {literal_appsheet(id_documento_raiz)})"
    )
    filas = appsheet_find(TABLA_DOCUMENTOS, selector)
    filas.sort(
        key=lambda fila: (
            0 if texto(fila.get("ID_DOCUMENTO")) == id_documento_raiz else 1,
            texto(fila.get("TITULO")).casefold(),
            texto(fila.get("ID_DOCUMENTO")),
        )
    )
    return filas


def calcular_estado_paquete_firma(
    integrantes: list[dict[str, Any]],
) -> str:
    if not integrantes:
        raise ValueError("El paquete externo no tiene documentos asociados")

    resultados = [
        normalizar_resultado_firma_externa(
            fila.get("RESULTADO_FIRMA_EXTERNA")
        )
        for fila in integrantes
    ]

    if all(resultado == "Firmado" for resultado in resultados):
        return "Listo para cierre"
    if any(resultado == "Observado" for resultado in resultados):
        return "Con observaciones"
    if any(resultado == "Pendiente" for resultado in resultados):
        return "Enviado"
    if all(resultado in {"Conforme", "Firmado"} for resultado in resultados):
        return "Esperando firma"
    return "Enviado"


def actualizar_estado_paquete_desde_resultados(
    id_documento_raiz: str,
    *,
    fecha: str | None = None,
) -> str:
    integrantes = buscar_integrantes_paquete_activo(id_documento_raiz)
    if not integrantes:
        raise LookupError(
            f"No existen integrantes activos para el paquete {id_documento_raiz}"
        )
    estado = calcular_estado_paquete_firma(integrantes)
    cambios: dict[str, Any] = {
        "ID_DOCUMENTO": id_documento_raiz,
        "ESTADO_PAQUETE_FIRMA": estado,
        "FECHA_ULTIMA_ACTUALIZACION": fecha or ahora_iso(),
    }
    if estado == "Con observaciones":
        cambios["ESTADO_FIRMA"] = "Observado"
    elif estado in {"Enviado", "Esperando firma", "Listo para cierre"}:
        cambios["ESTADO_FIRMA"] = "Pendiente"

    appsheet_action(TABLA_DOCUMENTOS, "Edit", [cambios])
    return estado


def validar_membresia_paquete_activo(
    *,
    id_documento_raiz: str,
    integrantes_jerarquia: list[dict[str, Any]],
) -> None:
    ids_jerarquia = {
        texto(fila.get("id_documento"))
        for fila in integrantes_jerarquia
    }
    integrantes_activos = buscar_integrantes_paquete_activo(id_documento_raiz)
    ids_activos = {
        texto(fila.get("ID_DOCUMENTO"))
        for fila in integrantes_activos
    }
    if ids_activos != ids_jerarquia:
        faltan = sorted(ids_jerarquia - ids_activos)
        sobran = sorted(ids_activos - ids_jerarquia)
        detalle: list[str] = []
        if faltan:
            detalle.append("sin asociación activa: " + ", ".join(faltan))
        if sobran:
            detalle.append("asociados pero fuera de la jerarquía: " + ", ".join(sobran))
        raise ValueError(
            "La composición actual de la jerarquía no coincide con el paquete "
            "externo activo. " + "; ".join(detalle)
        )


def validar_documento_observado_listo_reenvio(documento: dict[str, Any]) -> None:
    titulo = texto(documento.get("TITULO")) or texto(documento.get("ID_DOCUMENTO"))
    motivos: list[str] = []
    if texto(documento.get("ESTADO")) != "Listo para firma":
        motivos.append(
            f"estado {texto(documento.get('ESTADO')) or '<vacío>'}; se requiere Listo para firma"
        )
    responsable_valido, motivo_responsable = validar_responsable_firmas_actual(
        documento
    )
    if not responsable_valido:
        motivos.append(motivo_responsable)
    if not texto(documento.get("ID_VERSION_ACTUAL")):
        motivos.append("falta ID_VERSION_ACTUAL")
    if not texto(documento.get("GOOGLE_DOC_ID")):
        motivos.append("falta GOOGLE_DOC_ID")
    if not texto(documento.get("PDF_PARA_FIRMA_ID")):
        motivos.append("falta PDF_PARA_FIRMA_ID")

    id_version_actual = texto(documento.get("ID_VERSION_ACTUAL"))
    id_version_respuesta = texto(documento.get("ID_VERSION_RESPUESTA_EXTERNA"))
    if id_version_respuesta and id_version_actual == id_version_respuesta:
        motivos.append("la versión corregida todavía coincide con la versión observada")

    if motivos:
        raise ValueError(f"{titulo}: " + "; ".join(motivos))


def validar_documento_conforme_reenvio(documento: dict[str, Any]) -> None:
    titulo = texto(documento.get("TITULO")) or texto(documento.get("ID_DOCUMENTO"))
    id_actual = texto(documento.get("ID_VERSION_ACTUAL"))
    id_respuesta = texto(documento.get("ID_VERSION_RESPUESTA_EXTERNA"))
    if not id_actual or not id_respuesta or id_actual != id_respuesta:
        raise ValueError(
            f"{titulo}: el documento marcado Conforme cambió de versión y debe "
            "volver a aprobación externa"
        )
    if not texto(documento.get("PDF_PARA_FIRMA_ID")):
        raise ValueError(f"{titulo}: falta PDF_PARA_FIRMA_ID")


def validar_documento_firmado_reenvio(documento: dict[str, Any]) -> None:
    titulo = texto(documento.get("TITULO")) or texto(documento.get("ID_DOCUMENTO"))
    id_actual = texto(documento.get("ID_VERSION_ACTUAL"))
    id_respuesta = texto(documento.get("ID_VERSION_RESPUESTA_EXTERNA"))
    if not id_actual or not id_respuesta or id_actual != id_respuesta:
        raise ValueError(
            f"{titulo}: la versión actual no coincide con la versión que fue firmada"
        )
    if not texto(documento.get("PDF_FIRMADO")):
        raise ValueError(f"{titulo}: RESULTADO_FIRMA_EXTERNA=Firmado pero falta PDF_FIRMADO")


@medir_operacion("firma.construir_adjuntos_paquete_estado")
def construir_adjuntos_paquete_firma_por_estado(
    *,
    drive_service: Any,
    integrantes: list[dict[str, Any]],
    primer_envio: bool,
) -> tuple[list[dict[str, Any]], dict[str, list[dict[str, Any]]]]:
    """Construye los adjuntos y los clasifica para el cuerpo del email."""
    if not integrantes:
        raise ValueError("El paquete documental no contiene documentos")

    adjuntos: list[dict[str, Any]] = []
    secciones: dict[str, list[dict[str, Any]]] = {
        "aprobados_no_firmados": [],
        "firmados": [],
        "por_aprobar": [],
    }

    for fila in integrantes:
        documento = fila.get("documento") or fila
        id_documento = texto(documento.get("ID_DOCUMENTO"))
        titulo = texto(documento.get("TITULO")) or id_documento
        nivel = fila.get("nivel", 0) if isinstance(fila, dict) else 0
        es_raiz = bool(fila.get("es_raiz")) if isinstance(fila, dict) else False
        resultado = (
            "Pendiente"
            if primer_envio
            else normalizar_resultado_firma_externa(
                documento.get("RESULTADO_FIRMA_EXTERNA")
            )
        )

        detalle_base = {
            "id_documento": id_documento,
            "titulo": titulo,
            "nivel": nivel,
            "es_raiz": es_raiz,
            "resultado_anterior": resultado,
            "id_version": texto(documento.get("ID_VERSION_ACTUAL")),
            "numero_version": texto(documento.get("VERSION_ACTUAL")),
            "numero_revision": texto(documento.get("REVISION_ACTUAL")),
            "archivos": [],
        }

        if primer_envio or resultado == "Observado":
            if not primer_envio:
                validar_documento_observado_listo_reenvio(documento)
            pdf_id = texto(documento.get("PDF_PARA_FIRMA_ID"))
            google_doc_id = texto(documento.get("GOOGLE_DOC_ID"))
            if not pdf_id or not google_doc_id:
                raise ValueError(
                    f"{titulo}: faltan archivos para revisión externa"
                )
            pdf_bytes, pdf_nombre = descargar_pdf_drive(
                drive_service=drive_service,
                file_id=pdf_id,
            )
            docx_bytes, docx_nombre = exportar_docx_drive(
                drive_service=drive_service,
                google_doc_id=google_doc_id,
                nombre_base=pdf_nombre,
            )
            adjuntos.extend(
                [
                    {
                        "contenido": pdf_bytes,
                        "nombre": pdf_nombre,
                        "maintype": "application",
                        "subtype": "pdf",
                    },
                    {
                        "contenido": docx_bytes,
                        "nombre": docx_nombre,
                        "maintype": "application",
                        "subtype": (
                            "vnd.openxmlformats-officedocument."
                            "wordprocessingml.document"
                        ),
                    },
                ]
            )
            detalle_base["archivos"] = [pdf_nombre, docx_nombre]
            detalle_base["pdf_id"] = pdf_id
            detalle_base["google_doc_id"] = google_doc_id
            secciones["por_aprobar"].append(detalle_base)
            continue

        if resultado == "Conforme":
            validar_documento_conforme_reenvio(documento)
            pdf_id = texto(documento.get("PDF_PARA_FIRMA_ID"))
            pdf_bytes, pdf_nombre = descargar_pdf_drive(
                drive_service=drive_service,
                file_id=pdf_id,
            )
            adjuntos.append(
                {
                    "contenido": pdf_bytes,
                    "nombre": pdf_nombre,
                    "maintype": "application",
                    "subtype": "pdf",
                }
            )
            detalle_base["archivos"] = [pdf_nombre]
            detalle_base["pdf_id"] = pdf_id
            secciones["aprobados_no_firmados"].append(detalle_base)
            continue

        if resultado == "Firmado":
            validar_documento_firmado_reenvio(documento)
            metadata = buscar_pdf_cargado_appsheet(
                drive_service,
                texto(documento.get("PDF_FIRMADO")),
            )
            pdf_bytes, pdf_nombre = descargar_pdf_drive(
                drive_service=drive_service,
                file_id=texto(metadata.get("id")),
            )
            adjuntos.append(
                {
                    "contenido": pdf_bytes,
                    "nombre": pdf_nombre,
                    "maintype": "application",
                    "subtype": "pdf",
                }
            )
            detalle_base["archivos"] = [pdf_nombre]
            detalle_base["pdf_firmado_origen_id"] = texto(metadata.get("id"))
            secciones["firmados"].append(detalle_base)
            continue

        raise ValueError(
            f"{titulo}: no se puede reenviar con RESULTADO_FIRMA_EXTERNA={resultado!r}"
        )

    return adjuntos, secciones


def _filas_texto_seccion(titulo: str, filas: list[dict[str, Any]]) -> str:
    if not filas:
        return ""
    lineas = [titulo.upper()]
    for fila in filas:
        lineas.append(f"- {fila['titulo']}")
        for archivo in fila.get("archivos") or []:
            lineas.append(f"  · {archivo}")
    return "\n".join(lineas)


def _bloque_html_seccion(
    titulo: str,
    descripcion: str,
    filas: list[dict[str, Any]],
) -> str:
    if not filas:
        return ""
    items = []
    for fila in filas:
        archivos = "".join(
            f"<div style='color:#6b7280;font-size:12px;margin-top:2px;'>"
            f"{html.escape(texto(nombre))}</div>"
            for nombre in (fila.get("archivos") or [])
        )
        items.append(
            "<li style='margin:0 0 10px 0;'>"
            f"<strong>{html.escape(texto(fila.get('titulo')))}</strong>"
            f"{archivos}</li>"
        )
    return f"""
    <div style="margin-top:18px;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;">
      <div style="padding:11px 14px;background:#f3f4f6;font-weight:700;color:#111827;">
        {html.escape(titulo)}
      </div>
      <div style="padding:12px 16px;color:#374151;font-size:13px;line-height:1.5;">
        <div style="margin-bottom:10px;">{html.escape(descripcion)}</div>
        <ul style="margin:0;padding-left:20px;">{''.join(items)}</ul>
      </div>
    </div>
    """


def construir_email_firma_paquete_por_estado(
    *,
    documento: dict[str, Any],
    plantilla: dict[str, Any],
    usuario: str,
    mensaje_adicional: str,
    fecha_envio: str,
    secciones: dict[str, list[dict[str, Any]]],
    primer_envio: bool,
) -> tuple[str, str, str]:
    """Correo externo claro: qué revisar, qué firmar y qué ya está firmado."""
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    titulo_raiz = texto(documento.get("TITULO")) or f"Documento {id_documento}"
    tipo_documento = texto(documento.get("TIPO_DOCUMENTO")) or "-"
    proyecto = (
        texto(documento.get("NOMBRE_PROYECTO"))
        or texto(documento.get("PROYECTO"))
        or texto(documento.get("ID_PROYECTO"))
    )

    mapa_nombres = construir_mapa_nombres_usuarios(id_documento=id_documento)
    enviado_por_nombre = obtener_nombre_usuario_evento(
        usuario_evento=usuario,
        mapa_nombres=mapa_nombres,
    )
    if enviado_por_nombre == "Usuario no identificado":
        enviado_por_nombre = usuario

    variables = {
        "TITULO": titulo_raiz,
        "TIPO_DOCUMENTO": tipo_documento,
        "ID_PROYECTO": proyecto,
        "PROYECTO": proyecto,
        "ENVIADO_POR": usuario,
        "ENVIADO_POR_NOMBRE": enviado_por_nombre,
        "FECHA_ENVIO": fecha_envio,
    }
    asunto_base = texto(plantilla.get("ASUNTO_EMAIL_FIRMA")) or "Solicitud de firma — {{TITULO}}"
    asunto = reemplazar_variables_email(asunto_base, variables).strip()
    if not primer_envio:
        asunto = f"Actualización — {asunto}"

    cuerpo_base = texto(plantilla.get("CUERPO_EMAIL_FIRMA"))
    mensaje_plantilla = reemplazar_variables_email(cuerpo_base, variables).strip()
    mensaje_normalizado = mensaje_plantilla.lower()
    if (
        "adjuntamos el documento" in mensaje_normalizado
        and "una vez firmado" in mensaje_normalizado
        and "respondiendo a este correo" in mensaje_normalizado
    ):
        mensaje_plantilla = ""

    if primer_envio:
        introduccion = (
            f'Se adjunta el paquete documental asociado a "{titulo_raiz}" para su revisión y firma.'
        )
    else:
        introduccion = (
            f'Se adjunta una actualización del paquete documental asociado a "{titulo_raiz}". '
            "Los archivos se agrupan según su estado para identificar rápidamente qué debe revisar, "
            "qué debe firmar y qué ya se encuentra firmado."
        )

    aprobados = secciones.get("aprobados_no_firmados") or []
    firmados = secciones.get("firmados") or []
    por_aprobar = secciones.get("por_aprobar") or []

    bloques_texto = [
        "REVISIÓN Y FIRMA DE PAQUETE DOCUMENTAL",
        "",
        "Estimado/a:",
        "",
        introduccion,
        "",
    ]
    for titulo_seccion, filas in (
        ("Documentos aprobados no firmados", aprobados),
        ("Documentos firmados", firmados),
        ("Documentos por aprobar", por_aprobar),
    ):
        bloque = _filas_texto_seccion(titulo_seccion, filas)
        if bloque:
            bloques_texto.extend([bloque, ""])

    bloques_texto.extend(
        [
            "GUÍA RÁPIDA",
            "- Documentos aprobados no firmados: ya están aprobados; corresponde firmarlos.",
            "- Documentos firmados: ya están resueltos y se adjuntan como referencia.",
            "- Documentos por aprobar: requieren revisión; se adjuntan PDF y archivo editable.",
        ]
    )
    if mensaje_plantilla:
        bloques_texto.extend(["", "MENSAJE", mensaje_plantilla])
    if mensaje_adicional:
        bloques_texto.extend(["", "INDICACIONES ADICIONALES", mensaje_adicional])
    bloques_texto.extend(
        [
            "",
            "Puede responder directamente a este correo adjuntando los PDF firmados que correspondan.",
            "",
            f"Este correo fue generado por {NOMBRE_APLICACION}.",
        ]
    )
    cuerpo_texto = "\n".join(bloques_texto)

    def esc(valor: Any) -> str:
        return html.escape(texto(valor))

    def saltos(valor: Any) -> str:
        return esc(valor).replace("\n", "<br>")

    resumen_filas = [("Documento principal", titulo_raiz)]
    if proyecto:
        resumen_filas.append(("Proyecto", proyecto))
    resumen_filas.extend(
        [
            ("Tipo de documento", tipo_documento),
            ("Enviado por", enviado_por_nombre),
        ]
    )
    resumen_html = "".join(
        f"<tr><td style='padding:9px 12px;border-top:1px solid #e5e7eb;color:#6b7280;width:34%;font-size:13px;'>{esc(k)}</td>"
        f"<td style='padding:9px 12px;border-top:1px solid #e5e7eb;color:#111827;font-size:13px;font-weight:600;'>{esc(v)}</td></tr>"
        for k, v in resumen_filas if texto(v)
    )

    bloques_secciones = (
        _bloque_html_seccion(
            "Documentos aprobados no firmados",
            "Ya fueron aprobados previamente. No requieren nueva revisión; corresponde firmarlos.",
            aprobados,
        )
        + _bloque_html_seccion(
            "Documentos firmados",
            "Ya cuentan con PDF firmado y se incluyen como referencia del paquete.",
            firmados,
        )
        + _bloque_html_seccion(
            "Documentos por aprobar",
            "Son documentos nuevos o corregidos que requieren revisión. Se adjuntan PDF y archivo editable.",
            por_aprobar,
        )
    )

    bloque_mensaje = ""
    if mensaje_plantilla:
        bloque_mensaje = (
            "<div style='margin-top:18px;padding:14px 16px;background:#f9fafb;border-left:4px solid #6b7280;border-radius:6px;'>"
            "<div style='font-size:12px;font-weight:700;color:#374151;text-transform:uppercase;margin-bottom:6px;'>Mensaje</div>"
            f"<div style='font-size:14px;line-height:1.6;color:#374151;'>{saltos(mensaje_plantilla)}</div></div>"
        )
    bloque_adicional = ""
    if mensaje_adicional:
        bloque_adicional = (
            "<div style='margin-top:18px;padding:14px 16px;background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;'>"
            "<div style='font-size:12px;font-weight:700;color:#9a3412;text-transform:uppercase;margin-bottom:6px;'>Indicaciones adicionales</div>"
            f"<div style='font-size:14px;line-height:1.6;color:#7c2d12;'>{saltos(mensaje_adicional)}</div></div>"
        )

    cuerpo_html = f"""
    <!doctype html>
    <html lang="es">
      <body style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,Helvetica,sans-serif;color:#111827;">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="width:100%;background:#f3f4f6;">
          <tr><td align="center" style="padding:24px 12px;">
            <table role="presentation" width="680" cellspacing="0" cellpadding="0"
                   style="width:100%;max-width:680px;background:#ffffff;border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;">
              <tr><td style="padding:24px 28px;background:#111827;color:#ffffff;">
                <div style="font-size:13px;color:#d1d5db;">{esc(NOMBRE_APLICACION)}</div>
                <div style="margin-top:5px;font-size:24px;font-weight:700;">Revisión y firma de paquete documental</div>
                <div style="margin-top:8px;font-size:14px;color:#d1d5db;line-height:1.5;">
                  {'Primer envío externo' if primer_envio else 'Actualización del paquete externo'}
                </div>
              </td></tr>
              <tr><td style="padding:28px;">
                <p style="margin:0;font-size:15px;line-height:1.65;">Estimado/a:</p>
                <p style="margin:14px 0 0 0;font-size:15px;line-height:1.65;color:#374151;">{esc(introduccion)}</p>

                <div style="margin-top:22px;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;">
                  <div style="padding:11px 14px;background:#f9fafb;font-weight:700;font-size:14px;">Resumen del paquete</div>
                  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;">{resumen_html}</table>
                </div>

                {bloques_secciones}

                <div style="margin-top:18px;padding:14px 16px;background:#eef2ff;border:1px solid #c7d2fe;border-radius:8px;color:#3730a3;font-size:13px;line-height:1.6;">
                  <strong>Guía rápida</strong><br>
                  <strong>Aprobados no firmados:</strong> firmar.<br>
                  <strong>Firmados:</strong> sin acción; se incluyen como referencia.<br>
                  <strong>Por aprobar:</strong> revisar nuevamente antes de firmar.
                </div>

                {bloque_mensaje}
                {bloque_adicional}

                <p style="margin:24px 0 0 0;font-size:14px;line-height:1.6;color:#374151;">
                  Puede responder directamente a este correo adjuntando los PDF firmados que correspondan.
                </p>
              </td></tr>
              <tr><td style="padding:15px 28px;background:#f9fafb;border-top:1px solid #e5e7eb;color:#6b7280;font-size:12px;line-height:1.5;">
                Correo generado automáticamente por <strong>{esc(NOMBRE_APLICACION)}</strong>.
              </td></tr>
            </table>
          </td></tr>
        </table>
      </body>
    </html>
    """
    return asunto, cuerpo_texto, cuerpo_html


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


def restaurar_documento_tras_error_envio_firma(
    id_documento: str,
    mensaje: str,
    estado_firma_anterior: str = "No iniciado",
) -> None:
    try:
        appsheet_action(
            TABLA_DOCUMENTOS,
            "Edit",
            [
                {
                    "ID_DOCUMENTO": id_documento,
                    "ESTADO_FIRMA": estado_firma_anterior or "No iniciado",
                    "ACCION_SOLICITADA": "",
                    "OBSERVACION_ACTUAL": mensaje[:1000],
                    "FECHA_ULTIMA_ACTUALIZACION": ahora_iso(),
                }
            ],
        )
    except Exception:
        traceback.print_exc()


def actualizar_paquete_tras_envio_exitoso(
    *,
    id_documento_raiz: str,
    integrantes: list[dict[str, Any]],
    primer_envio: bool,
    usuario: str,
    destinatarios: list[str],
    mensaje_adicional: str,
    message_id: str,
    fecha: str,
) -> None:
    filas: list[dict[str, Any]] = []
    for fila in integrantes:
        documento = fila.get("documento") or fila
        id_documento = texto(documento.get("ID_DOCUMENTO"))
        cambios: dict[str, Any] = {
            "ID_DOCUMENTO": id_documento,
            "ID_DOCUMENTO_RAIZ_FIRMA": id_documento_raiz,
            "FECHA_ULTIMA_ACTUALIZACION": fecha,
        }

        resultado_previo = normalizar_resultado_firma_externa(
            documento.get("RESULTADO_FIRMA_EXTERNA")
        )
        if primer_envio or resultado_previo == "Observado":
            cambios.update(
                {
                    "RESULTADO_FIRMA_EXTERNA": "Pendiente",
                    "ID_VERSION_RESPUESTA_EXTERNA": texto(
                        documento.get("ID_VERSION_ACTUAL")
                    ),
                    "RESPUESTA_FIRMA_EXTERNA_POR": "",
                    "FECHA_RESPUESTA_FIRMA_EXTERNA": "",
                    "OBSERVACION_ACTUAL": "",
                }
            )

        if id_documento == id_documento_raiz:
            cambios.update(
                {
                    "ESTADO": "En firma",
                    "ESTADO_FIRMA": "Pendiente",
                    "ESTADO_PAQUETE_FIRMA": "Enviado",
                    "DESTINATARIOS_FIRMA": ", ".join(destinatarios),
                    "MENSAJE_ADICIONAL_FIRMA": mensaje_adicional,
                    "ENVIADO_FIRMA_POR": usuario,
                    "EMAIL_FIRMA_MESSAGE_ID": message_id,
                    "FECHA_ENVIO_FIRMA": fecha,
                    "ULTIMO_ENVIADO_POR": usuario,
                    "FECHA_ULTIMO_ENVIO": fecha,
                    "ACCION_SOLICITADA": "",
                }
            )
        filas.append(cambios)

    appsheet_action(TABLA_DOCUMENTOS, "Edit", filas)


def crear_eventos_envio_paquete_por_estado(
    *,
    id_documento_raiz: str,
    usuario: str,
    fecha: str,
    destinatarios: list[str],
    secciones: dict[str, list[dict[str, Any]]],
    message_id: str,
    primer_envio: bool,
) -> None:
    eventos: list[dict[str, Any]] = []
    etiquetas = {
        "aprobados_no_firmados": "Aprobado no firmado",
        "firmados": "Firmado",
        "por_aprobar": "Por aprobar",
    }
    for clave, filas in secciones.items():
        for fila in filas:
            es_raiz = texto(fila.get("id_documento")) == id_documento_raiz
            archivos = ", ".join(fila.get("archivos") or [])
            if primer_envio:
                tipo_evento = "Enviado a firma" if es_raiz else "Incluido en envío a firma"
            else:
                tipo_evento = "Reenviado a firma" if es_raiz else "Incluido en reenvío a firma"
            eventos.append(
                {
                    "ID_EVENTO": nuevo_id(),
                    "ID_DOCUMENTO": texto(fila.get("id_documento")),
                    "ID_VERSION": texto(fila.get("id_version")),
                    "ID_APROBACION_ACTUAL": "",
                    "TIPO_EVENTO": tipo_evento,
                    "ESTADO_ANTERIOR": texto(fila.get("resultado_anterior")),
                    "ESTADO_NUEVO": "Pendiente" if clave == "por_aprobar" else etiquetas[clave],
                    "USUARIO": usuario,
                    "FECHA_EVENTO": fecha,
                    "COMENTARIO": (
                        f"Paquete {'inicial' if primer_envio else 'reenviado'} a "
                        f"{', '.join(destinatarios)}. Categoría: {etiquetas[clave]}. "
                        f"Archivos: {archivos}. Gmail message ID: {message_id}."
                    ),
                }
            )
    if eventos:
        appsheet_action(TABLA_EVENTOS, "Add", eventos)


@app.route("/enviar-firma", methods=["POST"])
def enviar_firma():
    """Primer envío o reenvío de un paquete de Firma Simple."""
    id_documento = ""
    correo_enviado = False
    message_id = ""
    estado_firma_anterior = "No iniciado"

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

        contexto = obtener_contexto_jerarquia_documental(id_documento)
        if not contexto.get("solicitado_es_raiz"):
            raise ValueError(
                "Solo el documento raíz puede iniciar el envío externo. "
                f"Raíz: {contexto.get('id_documento_raiz')}"
            )
        if contexto.get("tipo_firma_paquete") != "Simple":
            raise ValueError(
                "Este endpoint corresponde a Firma Simple; el paquete no es Simple"
            )

        raiz = contexto["raiz"]
        integrantes = contexto["integrantes"]
        id_raiz = contexto["id_documento_raiz"]
        estado_paquete = texto(raiz.get("ESTADO_PAQUETE_FIRMA")) or "No iniciado"
        estado_firma_anterior = texto(raiz.get("ESTADO_FIRMA")) or "No iniciado"
        message_id_existente = texto(raiz.get("EMAIL_FIRMA_MESSAGE_ID"))

        primer_envio = estado_paquete in {"", "No iniciado"} and not texto(
            raiz.get("ID_DOCUMENTO_RAIZ_FIRMA")
        )
        reenvio = estado_paquete == "Con observaciones"

        if not primer_envio and not reenvio:
            if estado_paquete in {"Enviado", "Esperando firma", "Listo para cierre"} and message_id_existente:
                return jsonify(
                    {
                        "ok": True,
                        "ya_procesado": True,
                        "id_documento": id_raiz,
                        "estado_paquete_firma": estado_paquete,
                        "message_id": message_id_existente,
                        "fecha_envio_firma": texto(raiz.get("FECHA_ENVIO_FIRMA")),
                    }
                )
            raise ValueError(
                f"ESTADO_PAQUETE_FIRMA={estado_paquete!r} no permite enviar el paquete"
            )

        if estado_firma_anterior == "Enviando":
            return {"error": "Ya existe un envío de firma en proceso"}, 409

        if primer_envio:
            paquete = validar_paquete_documental_datos(id_raiz)
            if not paquete.get("paquete_listo"):
                detalles = []
                for fila in paquete.get("documentos_pendientes") or []:
                    detalles.append(
                        f"{texto(fila.get('titulo')) or texto(fila.get('id_documento'))}: "
                        + "; ".join(fila.get("motivos") or [])
                    )
                raise ValueError(
                    "No se puede realizar el primer envío: " + " | ".join(detalles)
                )
            if texto(raiz.get("ESTADO")) != "Listo para firma":
                raise ValueError("La raíz debe estar Listo para firma en el primer envío")
        else:
            validar_membresia_paquete_activo(
                id_documento_raiz=id_raiz,
                integrantes_jerarquia=integrantes,
            )
            pendientes = []
            observados = []
            for fila in integrantes:
                documento = fila["documento"]
                resultado = normalizar_resultado_firma_externa(
                    documento.get("RESULTADO_FIRMA_EXTERNA")
                )
                if resultado == "Pendiente":
                    pendientes.append(fila["titulo"])
                if resultado == "Observado":
                    observados.append(fila["titulo"])
            if pendientes:
                raise ValueError(
                    "No se puede reenviar mientras existan respuestas pendientes: "
                    + ", ".join(pendientes)
                )
            if not observados:
                raise ValueError("No existen documentos Observados para reenviar")

        usuario_registrado = texto(raiz.get("ENVIADO_FIRMA_POR"))
        usuario = usuario or usuario_registrado
        if not usuario or not _EMAIL_RE.fullmatch(usuario.lower()):
            raise ValueError("El usuario que envía a firma no es válido")
        if reenvio and usuario_registrado and usuario_registrado.lower() != usuario.lower():
            raise PermissionError(
                "Solo el usuario que administra este paquete puede reenviarlo"
            )

        if destinatarios_entrada in (None, ""):
            destinatarios_entrada = raiz.get("DESTINATARIOS_FIRMA")
        destinatarios = normalizar_destinatarios(destinatarios_entrada)
        if not destinatarios:
            raise ValueError("No se indicaron destinatarios para la firma")
        if not mensaje_adicional:
            mensaje_adicional = texto(raiz.get("MENSAJE_ADICIONAL_FIRMA"))

        plantilla = buscar_plantilla(texto(raiz.get("ID_PLANTILLA")))
        fecha_envio = ahora_iso()
        marcar_envio_firma_en_proceso(id_raiz, fecha_envio)
        drive_service = obtener_drive_service()
        gmail_service = obtener_gmail_service()

        adjuntos, secciones = construir_adjuntos_paquete_firma_por_estado(
            drive_service=drive_service,
            integrantes=integrantes,
            primer_envio=primer_envio,
        )
        asunto, cuerpo, cuerpo_html = construir_email_firma_paquete_por_estado(
            documento=raiz,
            plantilla=plantilla,
            usuario=usuario,
            mensaje_adicional=mensaje_adicional,
            fecha_envio=fecha_envio,
            secciones=secciones,
            primer_envio=primer_envio,
        )
        respuesta_gmail = enviar_email_con_adjuntos(
            gmail_service=gmail_service,
            destinatarios=destinatarios,
            asunto=asunto,
            cuerpo=cuerpo,
            adjuntos=adjuntos,
            reply_to=usuario,
            cuerpo_html=cuerpo_html,
        )
        correo_enviado = True
        message_id = respuesta_gmail["message_id"]

        actualizar_paquete_tras_envio_exitoso(
            id_documento_raiz=id_raiz,
            integrantes=integrantes,
            primer_envio=primer_envio,
            usuario=usuario,
            destinatarios=destinatarios,
            mensaje_adicional=mensaje_adicional,
            message_id=message_id,
            fecha=fecha_envio,
        )

        advertencias: list[str] = []
        try:
            crear_eventos_envio_paquete_por_estado(
                id_documento_raiz=id_raiz,
                usuario=usuario,
                fecha=fecha_envio,
                destinatarios=destinatarios,
                secciones=secciones,
                message_id=message_id,
                primer_envio=primer_envio,
            )
        except Exception as exc_evento:
            traceback.print_exc()
            advertencias.append(
                "El correo se envió, pero falló parte de la trazabilidad: "
                f"{exc_evento}"
            )

        return jsonify(
            {
                "ok": True,
                "ya_procesado": False,
                "tipo_envio": "Primer envío" if primer_envio else "Reenvío",
                "id_documento_raiz": id_raiz,
                "estado_paquete_firma": "Enviado",
                "message_id": message_id,
                "destinatarios": destinatarios,
                "cantidad_documentos": len(integrantes),
                "cantidad_adjuntos": len(adjuntos),
                "secciones": {
                    clave: [
                        {
                            "id_documento": fila.get("id_documento"),
                            "titulo": fila.get("titulo"),
                            "archivos": fila.get("archivos"),
                        }
                        for fila in filas
                    ]
                    for clave, filas in secciones.items()
                },
                "advertencias": advertencias,
            }
        )

    except PermissionError as exc:
        if id_documento and not correo_enviado:
            restaurar_documento_tras_error_envio_firma(
                id_documento,
                str(exc),
                estado_firma_anterior,
            )
        return {"error": str(exc)}, 403
    except (ValueError, LookupError) as exc:
        if id_documento and not correo_enviado:
            restaurar_documento_tras_error_envio_firma(
                id_documento,
                str(exc),
                estado_firma_anterior,
            )
        return {"error": str(exc)}, 400
    except Exception as exc:
        traceback.print_exc()
        if id_documento and not correo_enviado:
            restaurar_documento_tras_error_envio_firma(
                id_documento,
                str(exc),
                estado_firma_anterior,
            )
        return {
            "error": str(exc),
            "correo_enviado": correo_enviado,
            "message_id": message_id,
        }, 500

# -----------------------------------------------------------------------------
# Flujo Notarial - primer envío a revisión externa
# -----------------------------------------------------------------------------


def buscar_revisiones_externas_raiz(id_documento_raiz: str) -> list[dict[str, Any]]:
    selector = (
        f"FILTER({TABLA_REVISIONES_EXTERNAS}, "
        f"[ID_DOCUMENTO_RAIZ] = {literal_appsheet(id_documento_raiz)})"
    )
    filas = appsheet_find(TABLA_REVISIONES_EXTERNAS, selector)
    filas.sort(
        key=lambda fila: entero(
            fila.get("NUMERO_REVISION_EXTERNA") or 0,
            "NUMERO_REVISION_EXTERNA",
        )
    )
    return filas


def buscar_detalles_revision_externa(id_revision_externa: str) -> list[dict[str, Any]]:
    selector = (
        f"FILTER({TABLA_REVISION_EXTERNA_DETALLE}, "
        f"[ID_REVISION_EXTERNA] = {literal_appsheet(id_revision_externa)})"
    )
    return appsheet_find(TABLA_REVISION_EXTERNA_DETALLE, selector)


def validar_paquete_para_revision_externa_notarial(
    *,
    contexto: dict[str, Any],
    permitir_reanudacion_post_gmail: bool = False,
) -> list[dict[str, Any]]:
    """Valida raíz + descendientes y devuelve su versión exacta para envío."""
    if contexto.get("tipo_firma_paquete") != "Notarial":
        raise ValueError("El paquete no corresponde a Firma Notarial")

    preparados: list[dict[str, Any]] = []
    errores: list[str] = []

    for item in contexto.get("integrantes") or []:
        documento = item.get("documento") or {}
        id_documento = texto(documento.get("ID_DOCUMENTO"))
        titulo = texto(documento.get("TITULO")) or id_documento
        motivos: list[str] = []

        estado_documento = texto(documento.get("ESTADO"))
        estados_admitidos = {ESTADO_LISTO_REVISION_EXTERNA}
        if permitir_reanudacion_post_gmail:
            estados_admitidos.add("En revisión externa")
        if estado_documento not in estados_admitidos:
            motivos.append(
                f"ESTADO={estado_documento!r}; se requiere "
                + " o ".join(repr(valor) for valor in sorted(estados_admitidos))
            )

        if estado_documento == "En revisión externa":
            raiz_notarial_fila = texto(documento.get("ID_DOCUMENTO_RAIZ_NOTARIAL"))
            resultado_externo = texto(documento.get("RESULTADO_REVISION_EXTERNA"))
            id_version_externa = texto(documento.get("ID_VERSION_REVISION_EXTERNA"))
            id_raiz = texto(contexto.get("id_documento_raiz"))
            if raiz_notarial_fila and raiz_notarial_fila != id_raiz:
                motivos.append("ID_DOCUMENTO_RAIZ_NOTARIAL apunta a otra raíz")
            if resultado_externo and resultado_externo != "Pendiente":
                motivos.append(
                    f"RESULTADO_REVISION_EXTERNA={resultado_externo!r}; "
                    "se esperaba 'Pendiente' durante la reanudación"
                )
            if id_version_externa and id_version_externa != texto(documento.get("ID_VERSION_ACTUAL")):
                motivos.append(
                    "ID_VERSION_REVISION_EXTERNA no coincide con ID_VERSION_ACTUAL"
                )

        responsable_ok, motivo_responsable = validar_responsable_firmas_actual(
            documento
        )
        if not responsable_ok:
            motivos.append(motivo_responsable)

        id_version = texto(documento.get("ID_VERSION_ACTUAL"))
        version: dict[str, Any] = {}
        if not id_version:
            motivos.append("falta ID_VERSION_ACTUAL")
        else:
            try:
                version = buscar_version_por_id(id_version)
            except Exception as exc:
                motivos.append(f"no se pudo resolver ID_VERSION_ACTUAL: {exc}")

        if version:
            if texto(version.get("ID_DOCUMENTO")) != id_documento:
                motivos.append("ID_VERSION_ACTUAL pertenece a otro documento")
            if texto(version.get("ETAPA")) != "Para revisión externa":
                motivos.append(
                    "la versión vigente no está en ETAPA='Para revisión externa'"
                )
            if texto(version.get("ESTADO_VERSION")) != "Activa":
                motivos.append(
                    f"ESTADO_VERSION={texto(version.get('ESTADO_VERSION'))!r}; "
                    "se requiere 'Activa'"
                )
            if not texto(version.get("GOOGLE_DOC_ID")):
                motivos.append("la versión no tiene GOOGLE_DOC_ID")
            if not texto(version.get("PDF_VERSION_ID")):
                motivos.append("la versión no tiene PDF_VERSION_ID")

        if motivos:
            errores.append(f"{titulo}: " + "; ".join(motivos))
            continue

        preparados.append(
            {
                **item,
                "documento": documento,
                "version": version,
                "id_version": id_version,
                "titulo": titulo,
            }
        )

    if errores:
        raise ValueError(
            "El paquete no está listo para revisión externa: " + " | ".join(errores)
        )
    if not preparados:
        raise ValueError("El paquete notarial no contiene documentos")

    return preparados


@medir_operacion("notarial.construir_adjuntos_revision_externa")
def construir_adjuntos_revision_externa_notarial(
    *,
    drive_service: Any,
    integrantes_preparados: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Adjunta PDF + DOCX de la versión exacta Para revisión externa."""
    adjuntos: list[dict[str, Any]] = []
    resumen: list[dict[str, Any]] = []

    for item in integrantes_preparados:
        documento = item["documento"]
        version = item["version"]
        titulo = item["titulo"]
        pdf_id = texto(version.get("PDF_VERSION_ID"))
        google_doc_id = texto(version.get("GOOGLE_DOC_ID"))
        nombre_base = (
            texto(version.get("NOMBRE_ARCHIVO"))
            or f"{titulo}_PARA_REVISION_EXTERNA"
        )

        pdf_bytes, pdf_nombre = descargar_pdf_drive(
            drive_service=drive_service,
            file_id=pdf_id,
        )
        docx_bytes, docx_nombre = exportar_docx_drive(
            drive_service=drive_service,
            google_doc_id=google_doc_id,
            nombre_base=nombre_base,
        )

        adjuntos.extend(
            [
                {
                    "contenido": pdf_bytes,
                    "nombre": pdf_nombre,
                    "maintype": "application",
                    "subtype": "pdf",
                },
                {
                    "contenido": docx_bytes,
                    "nombre": docx_nombre,
                    "maintype": "application",
                    "subtype": (
                        "vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    ),
                },
            ]
        )
        resumen.append(
            {
                "id_documento": texto(documento.get("ID_DOCUMENTO")),
                "titulo": titulo,
                "id_version": item["id_version"],
                "numero_version": texto(documento.get("VERSION_ACTUAL")),
                "numero_revision": texto(documento.get("REVISION_ACTUAL")),
                "nivel": item.get("nivel", 0),
                "es_raiz": bool(item.get("es_raiz")),
                "archivos": [pdf_nombre, docx_nombre],
            }
        )

    return adjuntos, resumen


def construir_email_revision_externa_notarial(
    *,
    raiz: dict[str, Any],
    usuario: str,
    mensaje_adicional: str,
    fecha_envio: str,
    documentos: list[dict[str, Any]],
) -> tuple[str, str, str]:
    titulo_raiz = texto(raiz.get("TITULO")) or texto(raiz.get("ID_DOCUMENTO"))
    proyecto = texto(raiz.get("ID_PROYECTO"))
    tipo_documento = texto(raiz.get("TIPO_DOCUMENTO"))
    cantidad = len(documentos)

    asunto = f"Revisión externa de paquete documental - {titulo_raiz}"

    lineas = []
    for indice, item in enumerate(documentos, start=1):
        detalle = item["titulo"]
        if item.get("numero_version"):
            detalle += f" — V{item['numero_version']}"
            if item.get("numero_revision"):
                detalle += f" / Rev. {item['numero_revision']}"
        lineas.append(f"{indice}. {detalle}")

    bloques = [
        "REVISIÓN EXTERNA DE PAQUETE DOCUMENTAL",
        "",
        "Estimado/a:",
        "",
        (
            f"Se envía el paquete documental asociado a \"{titulo_raiz}\" "
            "para revisión externa."
        ),
        "",
        "DOCUMENTOS INCLUIDOS",
        *lineas,
        "",
        "ACCIÓN REQUERIDA",
        (
            "Por favor revise los documentos adjuntos y responda este mismo correo "
            "indicando, para cada documento, si queda Aprobado u Observado."
        ),
        (
            "Cuando existan observaciones, incorpore el comentario correspondiente "
            "y el respaldo que estime necesario."
        ),
    ]
    if mensaje_adicional:
        bloques.extend(["", "INDICACIONES ADICIONALES", mensaje_adicional])
    bloques.extend(
        [
            "",
            f"Enviado por: {usuario}",
            f"Fecha de envío: {fecha_envio}",
            "",
            f"Correo generado automáticamente por {NOMBRE_APLICACION}.",
        ]
    )
    cuerpo_texto = "\n".join(bloques)

    def esc(valor: Any) -> str:
        return html.escape(texto(valor))

    filas_html = []
    for indice, item in enumerate(documentos, start=1):
        detalle_version = ""
        if item.get("numero_version"):
            detalle_version = f"V{esc(item['numero_version'])}"
            if item.get("numero_revision"):
                detalle_version += f" / Rev. {esc(item['numero_revision'])}"
        filas_html.append(
            "<tr>"
            f"<td style='padding:10px 12px;border-top:1px solid #e5e7eb;width:34px;"
            f"color:#4338ca;font-weight:700;vertical-align:top;'>{indice}.</td>"
            "<td style='padding:10px 12px;border-top:1px solid #e5e7eb;"
            "color:#111827;font-size:14px;line-height:1.45;'>"
            f"<strong>{esc(item['titulo'])}</strong>"
            + (
                f"<br><span style='color:#6b7280;font-size:12px;'>{detalle_version}</span>"
                if detalle_version
                else ""
            )
            + "</td></tr>"
        )

    resumen = [
        ("Documento principal", titulo_raiz),
        ("Proyecto", proyecto),
        ("Tipo de documento", tipo_documento),
        ("Enviado por", usuario),
        ("Fecha de envío", fecha_envio),
    ]
    resumen_html = "".join(
        "<tr>"
        f"<td style='padding:9px 12px;border-top:1px solid #e5e7eb;"
        f"color:#6b7280;width:34%;font-size:13px;'>{esc(clave)}</td>"
        f"<td style='padding:9px 12px;border-top:1px solid #e5e7eb;"
        f"color:#111827;font-size:13px;font-weight:600;'>{esc(valor)}</td>"
        "</tr>"
        for clave, valor in resumen
        if texto(valor)
    )

    bloque_adicional = ""
    if mensaje_adicional:
        bloque_adicional = (
            "<div style='margin-top:18px;padding:14px 16px;background:#fff7ed;"
            "border:1px solid #fed7aa;border-radius:8px;'>"
            "<div style='font-size:12px;font-weight:700;color:#9a3412;"
            "text-transform:uppercase;margin-bottom:6px;'>Indicaciones adicionales</div>"
            f"<div style='font-size:14px;line-height:1.6;color:#7c2d12;'>"
            f"{esc(mensaje_adicional).replace(chr(10), '<br>')}</div></div>"
        )

    cuerpo_html = f"""
    <!doctype html>
    <html lang="es">
      <body style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,Helvetica,sans-serif;color:#111827;">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="width:100%;background:#f3f4f6;">
          <tr><td align="center" style="padding:24px 12px;">
            <table role="presentation" width="680" cellspacing="0" cellpadding="0"
                   style="width:100%;max-width:680px;background:#ffffff;border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;">
              <tr><td style="padding:24px 28px;background:#111827;color:#ffffff;">
                <div style="font-size:13px;color:#d1d5db;">{esc(NOMBRE_APLICACION)}</div>
                <div style="margin-top:5px;font-size:24px;font-weight:700;">Revisión externa de paquete documental</div>
                <div style="margin-top:8px;font-size:14px;color:#d1d5db;line-height:1.5;">Primer envío de revisión externa</div>
              </td></tr>
              <tr><td style="padding:28px;">
                <p style="margin:0;font-size:15px;line-height:1.65;">Estimado/a:</p>
                <p style="margin:14px 0 0 0;font-size:15px;line-height:1.65;color:#374151;">
                  Se envía el paquete documental asociado a <strong>{esc(titulo_raiz)}</strong> para revisión externa.
                </p>

                <div style="margin-top:22px;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;">
                  <div style="padding:11px 14px;background:#f9fafb;font-weight:700;font-size:14px;">Resumen del paquete</div>
                  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;">{resumen_html}</table>
                </div>

                <div style="margin-top:22px;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;">
                  <div style="padding:12px 16px;background:#eef2ff;color:#3730a3;font-size:14px;font-weight:700;">
                    Documentos incluidos ({cantidad})
                  </div>
                  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;">
                    {''.join(filas_html)}
                  </table>
                </div>

                <div style="margin-top:18px;padding:14px 16px;background:#eef2ff;border:1px solid #c7d2fe;border-radius:8px;color:#3730a3;font-size:13px;line-height:1.65;">
                  <strong>Acción requerida</strong><br>
                  Revise los documentos adjuntos y responda este mismo correo indicando, para cada documento, si queda <strong>Aprobado</strong> u <strong>Observado</strong>.<br>
                  Si existen observaciones, incorpore el comentario correspondiente y el respaldo que estime necesario.
                </div>

                {bloque_adicional}

                <p style="margin:24px 0 0 0;font-size:14px;line-height:1.6;color:#374151;">
                  Puede responder directamente a este correo. La respuesta será registrada por el Responsable de firmas en el sistema.
                </p>
              </td></tr>
              <tr><td style="padding:15px 28px;background:#f9fafb;border-top:1px solid #e5e7eb;color:#6b7280;font-size:12px;line-height:1.5;">
                Correo generado automáticamente por <strong>{esc(NOMBRE_APLICACION)}</strong>.
              </td></tr>
            </table>
          </td></tr>
        </table>
      </body>
    </html>
    """

    return asunto, cuerpo_texto, cuerpo_html


def marcar_correo_revision_externa_enviado(
    *,
    id_documento_raiz: str,
    usuario: str,
    destinatarios: list[str],
    mensaje_adicional: str,
    message_id: str,
    thread_id: str,
    fecha: str,
) -> None:
    """Marca Gmail inmediatamente para poder reanudar sin reenviar el correo."""
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento_raiz,
                "DESTINATARIOS_REVISION_EXTERNA": ", ".join(destinatarios),
                "MENSAJE_ADICIONAL_REVISION_EXTERNA": mensaje_adicional,
                "ENVIADO_REVISION_EXTERNA_POR": usuario,
                "FECHA_ENVIO_REVISION_EXTERNA": fecha,
                "EMAIL_REVISION_EXTERNA_MESSAGE_ID": message_id,
                "GMAIL_THREAD_ID_REVISION_EXTERNA": thread_id,
                "ULTIMO_ENVIADO_POR": usuario,
                "FECHA_ULTIMO_ENVIO": fecha,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
            }
        ],
    )


def obtener_o_crear_revision_externa_cabecera(
    *,
    id_documento_raiz: str,
    usuario: str,
    destinatarios: list[str],
    fecha: str,
    message_id: str,
    thread_id: str,
    mensaje_enviado: str,
) -> tuple[str, int, bool]:
    revisiones = buscar_revisiones_externas_raiz(id_documento_raiz)
    for revision in revisiones:
        if texto(revision.get("GMAIL_MESSAGE_ID")) == message_id:
            return (
                texto(revision.get("ID_REVISION_EXTERNA")),
                entero(
                    revision.get("NUMERO_REVISION_EXTERNA") or 1,
                    "NUMERO_REVISION_EXTERNA",
                ),
                False,
            )

    numero = 1
    if revisiones:
        numero = max(
            entero(
                fila.get("NUMERO_REVISION_EXTERNA") or 0,
                "NUMERO_REVISION_EXTERNA",
            )
            for fila in revisiones
        ) + 1

    id_revision = nuevo_id()
    appsheet_action(
        TABLA_REVISIONES_EXTERNAS,
        "Add",
        [
            {
                "ID_REVISION_EXTERNA": id_revision,
                "ID_DOCUMENTO_RAIZ": id_documento_raiz,
                "NUMERO_REVISION_EXTERNA": numero,
                "ESTADO_REVISION": "Pendiente",
                "DESTINATARIOS": ", ".join(destinatarios),
                "FECHA_ENVIO": fecha,
                "ENVIADO_POR": usuario,
                "FECHA_RESPUESTA": "",
                "REGISTRADO_POR": "",
                "COMENTARIO_RESPUESTA": "",
                "ARCHIVO_RESPALDO_RESPUESTA": "",
                "NOMBRE_ARCHIVO_RESPALDO": "",
                "GMAIL_MESSAGE_ID": message_id,
                "MENSAJE_ENVIADO": mensaje_enviado,
                "GMAIL_THREAD_ID": thread_id,
            }
        ],
    )
    return id_revision, numero, True


def crear_detalles_revision_externa_faltantes(
    *,
    id_revision_externa: str,
    integrantes_preparados: list[dict[str, Any]],
) -> int:
    existentes = buscar_detalles_revision_externa(id_revision_externa)
    ids_existentes = {
        texto(fila.get("ID_DOCUMENTO"))
        for fila in existentes
        if texto(fila.get("ID_DOCUMENTO"))
    }

    nuevas: list[dict[str, Any]] = []
    for item in integrantes_preparados:
        documento = item["documento"]
        id_documento = texto(documento.get("ID_DOCUMENTO"))
        if id_documento in ids_existentes:
            continue
        nuevas.append(
            {
                "ID_REVISION_EXTERNA_DETALLE": nuevo_id(),
                "ID_REVISION_EXTERNA": id_revision_externa,
                "ID_DOCUMENTO": id_documento,
                "ID_VERSION_ENVIADA": item["id_version"],
                "RESULTADO": "Pendiente",
                "COMENTARIO": "",
                "ARCHIVO_RESPALDO_RESPUESTA": "",
                "NOMBRE_ARCHIVO_RESPALDO": "",
                "FECHA_RESPUESTA": "",
                "REGISTRADO_POR": "",
            }
        )

    if nuevas:
        appsheet_action(TABLA_REVISION_EXTERNA_DETALLE, "Add", nuevas)
    return len(nuevas)


def actualizar_documentos_tras_envio_revision_externa(
    *,
    id_documento_raiz: str,
    integrantes_preparados: list[dict[str, Any]],
    usuario: str,
    destinatarios: list[str],
    mensaje_adicional: str,
    message_id: str,
    thread_id: str,
    fecha: str,
) -> None:
    filas: list[dict[str, Any]] = []
    for item in integrantes_preparados:
        documento = item["documento"]
        id_documento = texto(documento.get("ID_DOCUMENTO"))
        cambios: dict[str, Any] = {
            "ID_DOCUMENTO": id_documento,
            "ESTADO": "En revisión externa",
            "ID_DOCUMENTO_RAIZ_NOTARIAL": id_documento_raiz,
            "RESULTADO_REVISION_EXTERNA": "Pendiente",
            "ID_VERSION_REVISION_EXTERNA": item["id_version"],
            "RESPUESTA_REVISION_EXTERNA_POR": "",
            "FECHA_RESPUESTA_REVISION_EXTERNA": "",
            "FECHA_ULTIMA_ACTUALIZACION": fecha,
            "OBSERVACION_ACTUAL": "",
        }

        if id_documento == id_documento_raiz:
            cambios.update(
                {
                    "ESTADO_PAQUETE_NOTARIAL": "En revisión externa",
                    "DESTINATARIOS_REVISION_EXTERNA": ", ".join(destinatarios),
                    "MENSAJE_ADICIONAL_REVISION_EXTERNA": mensaje_adicional,
                    "ENVIADO_REVISION_EXTERNA_POR": usuario,
                    "FECHA_ENVIO_REVISION_EXTERNA": fecha,
                    "EMAIL_REVISION_EXTERNA_MESSAGE_ID": message_id,
                    "GMAIL_THREAD_ID_REVISION_EXTERNA": thread_id,
                    "ULTIMO_ENVIADO_POR": usuario,
                    "FECHA_ULTIMO_ENVIO": fecha,
                    "ACCION_SOLICITADA": "",
                }
            )
        filas.append(cambios)

    appsheet_action(TABLA_DOCUMENTOS, "Edit", filas)


def crear_eventos_envio_revision_externa(
    *,
    integrantes_preparados: list[dict[str, Any]],
    id_revision_externa: str,
    numero_revision_externa: int,
    usuario: str,
    destinatarios: list[str],
    fecha: str,
) -> None:
    filas = []
    for item in integrantes_preparados:
        documento = item["documento"]
        filas.append(
            {
                "ID_EVENTO": nuevo_id(),
                "ID_DOCUMENTO": texto(documento.get("ID_DOCUMENTO")),
                "ID_VERSION": item["id_version"],
                "ID_APROBACION_ACTUAL": texto(
                    documento.get("ID_APROBACION_ACTUAL")
                ),
                "TIPO_EVENTO": "Envío a revisión externa",
                "ESTADO_ANTERIOR": ESTADO_LISTO_REVISION_EXTERNA,
                "ESTADO_NUEVO": "En revisión externa",
                "USUARIO": usuario,
                "FECHA_EVENTO": fecha,
                "COMENTARIO": (
                    f"Revisión externa N° {numero_revision_externa}. "
                    f"Destinatarios: {', '.join(destinatarios)}. "
                    f"ID revisión externa: {id_revision_externa}."
                ),
            }
        )
    if filas:
        appsheet_action(TABLA_EVENTOS, "Add", filas)


def registrar_error_envio_revision_externa(
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
                    "ACCION_SOLICITADA": "",
                    "OBSERVACION_ACTUAL": mensaje[:1000],
                    "FECHA_ULTIMA_ACTUALIZACION": ahora_iso(),
                }
            ],
        )
    except Exception:
        traceback.print_exc()


@app.route("/enviar-revision-externa", methods=["POST"])
def enviar_revision_externa():
    """Primer envío del paquete Notarial al revisor externo."""
    id_documento = ""
    correo_enviado = False
    message_id = ""
    thread_id = ""

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

        contexto = obtener_contexto_jerarquia_documental(id_documento)
        if not contexto.get("solicitado_es_raiz"):
            raise ValueError(
                "Solo la raíz puede iniciar la revisión externa. "
                f"Raíz: {contexto.get('id_documento_raiz')}"
            )
        if contexto.get("tipo_firma_paquete") != "Notarial":
            raise ValueError(
                "Este endpoint corresponde a Firma Notarial; el paquete no es Notarial"
            )

        raiz = contexto["raiz"]
        id_raiz = contexto["id_documento_raiz"]
        estado_paquete = texto(raiz.get("ESTADO_PAQUETE_NOTARIAL")) or "No iniciado"
        id_raiz_notarial = texto(raiz.get("ID_DOCUMENTO_RAIZ_NOTARIAL"))
        message_id_existente = texto(raiz.get("EMAIL_REVISION_EXTERNA_MESSAGE_ID"))
        thread_id_existente = texto(raiz.get("GMAIL_THREAD_ID_REVISION_EXTERNA"))

        if (
            estado_paquete == "En revisión externa"
            and id_raiz_notarial == id_raiz
            and message_id_existente
        ):
            return jsonify(
                {
                    "ok": True,
                    "ya_procesado": True,
                    "id_documento_raiz": id_raiz,
                    "estado_paquete_notarial": estado_paquete,
                    "message_id": message_id_existente,
                    "thread_id": thread_id_existente,
                    "fecha_envio_revision_externa": texto(
                        raiz.get("FECHA_ENVIO_REVISION_EXTERNA")
                    ),
                }
            )

        reanudando_post_gmail = bool(message_id_existente)
        estados_permitidos = {"", "No iniciado"}
        if reanudando_post_gmail:
            estados_permitidos.add("En revisión externa")
        if estado_paquete not in estados_permitidos:
            raise ValueError(
                f"ESTADO_PAQUETE_NOTARIAL={estado_paquete!r} no permite el primer envío"
            )
        if id_raiz_notarial and not (
            reanudando_post_gmail and id_raiz_notarial == id_raiz
        ):
            raise ValueError(
                "El paquete ya tiene ID_DOCUMENTO_RAIZ_NOTARIAL; "
                "para reenviar documentos corregidos use /reenviar-revision-externa"
            )

        preparados = validar_paquete_para_revision_externa_notarial(
            contexto=contexto,
            permitir_reanudacion_post_gmail=reanudando_post_gmail,
        )

        usuario_guardado = texto(raiz.get("ENVIADO_REVISION_EXTERNA_POR"))
        usuario = usuario or usuario_guardado
        if not usuario or not _EMAIL_RE.fullmatch(usuario.lower()):
            raise ValueError("El usuario que envía a revisión externa no es válido")

        if destinatarios_entrada in (None, ""):
            destinatarios_entrada = raiz.get("DESTINATARIOS_REVISION_EXTERNA")
        destinatarios = normalizar_destinatarios(destinatarios_entrada)
        if not destinatarios:
            raise ValueError("No se indicaron destinatarios para la revisión externa")

        if not mensaje_adicional:
            mensaje_adicional = texto(
                raiz.get("MENSAJE_ADICIONAL_REVISION_EXTERNA")
            )

        fecha_envio = (
            texto(raiz.get("FECHA_ENVIO_REVISION_EXTERNA"))
            if message_id_existente
            else ahora_iso()
        ) or ahora_iso()

        resumen_documentos = [
            {
                "id_documento": texto(item["documento"].get("ID_DOCUMENTO")),
                "titulo": item["titulo"],
                "id_version": item["id_version"],
                "numero_version": texto(item["documento"].get("VERSION_ACTUAL")),
                "numero_revision": texto(item["documento"].get("REVISION_ACTUAL")),
                "nivel": item.get("nivel", 0),
                "es_raiz": bool(item.get("es_raiz")),
            }
            for item in preparados
        ]
        asunto, cuerpo_texto, cuerpo_html = construir_email_revision_externa_notarial(
            raiz=raiz,
            usuario=usuario,
            mensaje_adicional=mensaje_adicional,
            fecha_envio=fecha_envio,
            documentos=resumen_documentos,
        )

        cantidad_adjuntos = 0

        if reanudando_post_gmail:
            correo_enviado = True
            message_id = message_id_existente
            thread_id = thread_id_existente
        else:
            drive_service = obtener_drive_service()
            gmail_service = obtener_gmail_service()
            adjuntos, resumen_documentos = construir_adjuntos_revision_externa_notarial(
                drive_service=drive_service,
                integrantes_preparados=preparados,
            )
            cantidad_adjuntos = len(adjuntos)
            asunto, cuerpo_texto, cuerpo_html = construir_email_revision_externa_notarial(
                raiz=raiz,
                usuario=usuario,
                mensaje_adicional=mensaje_adicional,
                fecha_envio=fecha_envio,
                documentos=resumen_documentos,
            )
            respuesta_gmail = enviar_email_con_adjuntos(
                gmail_service=gmail_service,
                destinatarios=destinatarios,
                asunto=asunto,
                cuerpo=cuerpo_texto,
                adjuntos=adjuntos,
                reply_to=usuario,
                cuerpo_html=cuerpo_html,
            )
            correo_enviado = True
            message_id = respuesta_gmail["message_id"]
            thread_id = respuesta_gmail["thread_id"]

            marcar_correo_revision_externa_enviado(
                id_documento_raiz=id_raiz,
                usuario=usuario,
                destinatarios=destinatarios,
                mensaje_adicional=mensaje_adicional,
                message_id=message_id,
                thread_id=thread_id,
                fecha=fecha_envio,
            )

        id_revision_externa, numero_revision_externa, cabecera_creada = (
            obtener_o_crear_revision_externa_cabecera(
                id_documento_raiz=id_raiz,
                usuario=usuario,
                destinatarios=destinatarios,
                fecha=fecha_envio,
                message_id=message_id,
                thread_id=thread_id,
                mensaje_enviado=cuerpo_texto,
            )
        )

        detalles_creados = crear_detalles_revision_externa_faltantes(
            id_revision_externa=id_revision_externa,
            integrantes_preparados=preparados,
        )

        actualizar_documentos_tras_envio_revision_externa(
            id_documento_raiz=id_raiz,
            integrantes_preparados=preparados,
            usuario=usuario,
            destinatarios=destinatarios,
            mensaje_adicional=mensaje_adicional,
            message_id=message_id,
            thread_id=thread_id,
            fecha=fecha_envio,
        )

        advertencias: list[str] = []
        try:
            crear_eventos_envio_revision_externa(
                integrantes_preparados=preparados,
                id_revision_externa=id_revision_externa,
                numero_revision_externa=numero_revision_externa,
                usuario=usuario,
                destinatarios=destinatarios,
                fecha=fecha_envio,
            )
        except Exception as exc_evento:
            traceback.print_exc()
            advertencias.append(
                "El envío quedó registrado, pero falló la creación de eventos: "
                f"{exc_evento}"
            )

        return jsonify(
            {
                "ok": True,
                "ya_procesado": False,
                "reanudado_post_gmail": reanudando_post_gmail,
                "id_documento_raiz": id_raiz,
                "id_revision_externa": id_revision_externa,
                "numero_revision_externa": numero_revision_externa,
                "estado_paquete_notarial": "En revisión externa",
                "message_id": message_id,
                "thread_id": thread_id,
                "destinatarios": destinatarios,
                "cantidad_documentos": len(preparados),
                "cantidad_adjuntos": cantidad_adjuntos,
                "cabecera_creada": cabecera_creada,
                "detalles_creados": detalles_creados,
                "advertencias": advertencias,
            }
        )

    except PermissionError as exc:
        if id_documento:
            registrar_error_envio_revision_externa(id_documento, str(exc))
        return {"error": str(exc)}, 403
    except (ValueError, LookupError) as exc:
        if id_documento:
            registrar_error_envio_revision_externa(id_documento, str(exc))
        return {"error": str(exc)}, 400
    except Exception as exc:
        traceback.print_exc()
        if id_documento:
            registrar_error_envio_revision_externa(id_documento, str(exc))
        return {
            "error": str(exc),
            "correo_enviado": correo_enviado,
            "message_id": message_id,
            "thread_id": thread_id,
        }, 500


# -----------------------------------------------------------------------------
# Flujo Notarial - reenvío individual de documento corregido
# -----------------------------------------------------------------------------


def buscar_revision_externa_pendiente_documento(
    *,
    id_documento: str,
    id_version: str = "",
) -> tuple[dict[str, Any], dict[str, Any]] | None:
    """Busca una ronda Pendiente que contenga al documento como detalle Pendiente."""
    selector = (
        f"FILTER({TABLA_REVISION_EXTERNA_DETALLE}, "
        f"AND([ID_DOCUMENTO] = {literal_appsheet(id_documento)}, "
        f"[RESULTADO] = \"Pendiente\"))"
    )
    detalles = appsheet_find(TABLA_REVISION_EXTERNA_DETALLE, selector)
    coincidencias: list[tuple[dict[str, Any], dict[str, Any]]] = []

    for detalle in detalles:
        if id_version and texto(detalle.get("ID_VERSION_ENVIADA")) != id_version:
            continue
        id_revision = texto(detalle.get("ID_REVISION_EXTERNA"))
        if not id_revision:
            continue
        try:
            revision = buscar_revision_externa_por_id(id_revision)
        except LookupError:
            continue
        if texto(revision.get("ESTADO_REVISION")) != "Pendiente":
            continue
        coincidencias.append((revision, detalle))

    if len(coincidencias) > 1:
        raise ValueError(
            "Existe más de una revisión externa Pendiente para el mismo documento"
        )
    return coincidencias[0] if coincidencias else None


def validar_documento_para_reenvio_revision_externa_notarial(
    *,
    id_documento: str,
    id_documento_raiz: str,
    id_version: str,
    permitir_pendiente_version_actual: bool = False,
) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any]]:
    """Valida un único documento observado que terminó nuevamente aprobación interna."""
    documento = buscar_documento(id_documento)
    id_raiz_fila = texto(documento.get("ID_DOCUMENTO_RAIZ_NOTARIAL"))
    if not id_raiz_fila:
        raise ValueError("El documento no pertenece a un paquete Notarial")
    if id_documento_raiz and id_documento_raiz != id_raiz_fila:
        raise ValueError("id_documento_raiz no coincide con el paquete Notarial")

    raiz = buscar_documento(id_raiz_fila)
    if normalizar_tipo_firma(raiz.get("TIPO_FIRMA")) != "Notarial":
        raise ValueError("La raíz del paquete no corresponde a Firma Notarial")
    if texto(raiz.get("ESTADO_PAQUETE_NOTARIAL")) != "Con observaciones":
        raise ValueError(
            "Solo un paquete con ESTADO_PAQUETE_NOTARIAL='Con observaciones' "
            "admite el reenvío de documentos corregidos"
        )

    if texto(documento.get("ESTADO")) != ESTADO_LISTO_REVISION_EXTERNA:
        raise ValueError(
            f"El documento debe estar en {ESTADO_LISTO_REVISION_EXTERNA!r} para reenviarse"
        )
    if normalizar_resultado_revision_externa(
        documento.get("RESULTADO_REVISION_EXTERNA")
    ) != "Observado":
        raise ValueError(
            "Solo un documento RESULTADO_REVISION_EXTERNA='Observado' puede reenviarse"
        )

    id_version_actual = texto(documento.get("ID_VERSION_ACTUAL"))
    id_version_anterior = texto(documento.get("ID_VERSION_REVISION_EXTERNA"))
    id_version_solicitada = id_version or id_version_actual
    if not id_version_solicitada:
        raise ValueError("Falta ID_VERSION_ACTUAL para el reenvío")
    if id_version_solicitada != id_version_actual:
        raise ValueError("id_version no coincide con ID_VERSION_ACTUAL")
    if id_version_anterior and id_version_solicitada == id_version_anterior:
        raise ValueError(
            "No se puede reenviar la misma versión que fue observada externamente"
        )

    pendiente = buscar_revision_externa_pendiente_documento(
        id_documento=id_documento,
    )
    if pendiente is not None:
        revision_pendiente, detalle_pendiente = pendiente
        pendiente_es_version_actual = (
            texto(detalle_pendiente.get("ID_VERSION_ENVIADA"))
            == id_version_solicitada
        )
        if not (
            permitir_pendiente_version_actual
            and pendiente_es_version_actual
        ):
            raise ValueError(
                "El documento ya participa en una revisión externa Pendiente: "
                f"{texto(revision_pendiente.get('ID_REVISION_EXTERNA'))}; "
                f"versión {texto(detalle_pendiente.get('ID_VERSION_ENVIADA'))}"
            )

    responsable_ok, motivo_responsable = validar_responsable_firmas_actual(documento)
    if not responsable_ok:
        raise ValueError(motivo_responsable)

    version = buscar_version_por_id(id_version_solicitada)
    if texto(version.get("ID_DOCUMENTO")) != id_documento:
        raise ValueError("ID_VERSION_ACTUAL pertenece a otro documento")
    if texto(version.get("ETAPA")) != "Para revisión externa":
        raise ValueError(
            "La versión vigente no está en ETAPA='Para revisión externa'"
        )
    if texto(version.get("ESTADO_VERSION")) != "Activa":
        raise ValueError(
            f"ESTADO_VERSION={texto(version.get('ESTADO_VERSION'))!r}; se requiere 'Activa'"
        )
    if not texto(version.get("GOOGLE_DOC_ID")):
        raise ValueError("La versión no tiene GOOGLE_DOC_ID")
    if not texto(version.get("PDF_VERSION_ID")):
        raise ValueError("La versión no tiene PDF_VERSION_ID")

    preparado = {
        "documento": documento,
        "version": version,
        "id_version": id_version_solicitada,
        "titulo": texto(documento.get("TITULO")) or id_documento,
        "nivel": 0,
        "es_raiz": id_documento == id_raiz_fila,
    }
    return documento, raiz, version, preparado


def siguiente_numero_revision_externa(id_documento_raiz: str) -> int:
    revisiones = buscar_revisiones_externas_raiz(id_documento_raiz)
    if not revisiones:
        return 1
    return max(
        entero(
            fila.get("NUMERO_REVISION_EXTERNA") or 0,
            "NUMERO_REVISION_EXTERNA",
        )
        for fila in revisiones
    ) + 1


def construir_email_reenvio_revision_externa_notarial(
    *,
    raiz: dict[str, Any],
    documento: dict[str, Any],
    usuario: str,
    mensaje_adicional: str,
    fecha_envio: str,
    numero_revision_externa: int,
    archivos: list[str],
) -> tuple[str, str, str]:
    """Correo externo para reenviar solamente la nueva versión corregida."""
    titulo_raiz = texto(raiz.get("TITULO")) or texto(raiz.get("ID_DOCUMENTO"))
    titulo = texto(documento.get("TITULO")) or texto(documento.get("ID_DOCUMENTO"))
    numero_version = texto(documento.get("VERSION_ACTUAL"))
    numero_revision = texto(documento.get("REVISION_ACTUAL"))
    version_texto = f"V{numero_version}" if numero_version else ""
    if numero_revision:
        version_texto += f" / Rev. {numero_revision}"

    asunto = (
        f"Reenvío revisión externa N° {numero_revision_externa} - {titulo}"
    )
    lineas_archivos = [f"- {nombre}" for nombre in archivos]
    bloques = [
        "REENVÍO DE DOCUMENTO CORREGIDO A REVISIÓN EXTERNA",
        "",
        "Estimado/a:",
        "",
        (
            f"Se reenvía el documento corregido \"{titulo}\" del paquete "
            f"\"{titulo_raiz}\" para una nueva revisión externa."
        ),
        f"Ronda externa: N° {numero_revision_externa}",
        *( [f"Versión enviada: {version_texto}"] if version_texto else [] ),
        "",
        "ARCHIVOS ADJUNTOS",
        *lineas_archivos,
        "",
        "ACCIÓN REQUERIDA",
        "Por favor revise esta nueva versión e indique si queda Aprobada u Observada.",
    ]
    if mensaje_adicional:
        bloques.extend(["", "INDICACIONES ADICIONALES", mensaje_adicional])
    bloques.extend(
        [
            "",
            f"Enviado por: {usuario}",
            f"Fecha de envío: {fecha_envio}",
            "",
            f"Correo generado automáticamente por {NOMBRE_APLICACION}.",
        ]
    )
    cuerpo_texto = "\n".join(bloques)

    def esc(valor: Any) -> str:
        return html.escape(texto(valor))

    archivos_html = "".join(
        f"<li style='margin:4px 0;'>{esc(nombre)}</li>" for nombre in archivos
    )
    adicional_html = ""
    if mensaje_adicional:
        adicional_html = (
            "<div style='margin-top:18px;padding:14px 16px;background:#fff7ed;"
            "border:1px solid #fed7aa;border-radius:8px;'>"
            "<div style='font-size:12px;font-weight:700;color:#9a3412;"
            "text-transform:uppercase;margin-bottom:6px;'>Indicaciones adicionales</div>"
            f"<div style='font-size:14px;line-height:1.6;color:#7c2d12;'>"
            f"{esc(mensaje_adicional).replace(chr(10), '<br>')}</div></div>"
        )

    cuerpo_html = f"""
    <!doctype html>
    <html lang="es">
      <body style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,Helvetica,sans-serif;color:#111827;">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="width:100%;background:#f3f4f6;">
          <tr><td align="center" style="padding:24px 12px;">
            <table role="presentation" width="680" cellspacing="0" cellpadding="0"
                   style="width:100%;max-width:680px;background:#ffffff;border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;">
              <tr><td style="padding:24px 28px;background:#111827;color:#ffffff;">
                <div style="font-size:13px;color:#d1d5db;">{esc(NOMBRE_APLICACION)}</div>
                <div style="margin-top:5px;font-size:24px;font-weight:700;">Reenvío a revisión externa</div>
                <div style="margin-top:8px;font-size:14px;color:#d1d5db;line-height:1.5;">Ronda externa N° {numero_revision_externa}</div>
              </td></tr>
              <tr><td style="padding:28px;">
                <p style="margin:0;font-size:15px;line-height:1.65;">Estimado/a:</p>
                <p style="margin:14px 0 0 0;font-size:15px;line-height:1.65;color:#374151;">
                  Se reenvía el documento corregido <strong>{esc(titulo)}</strong> del paquete
                  <strong>{esc(titulo_raiz)}</strong> para una nueva revisión externa.
                </p>
                <div style="margin-top:20px;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;">
                  <div style="padding:11px 14px;background:#f9fafb;font-weight:700;font-size:14px;">Documento reenviado</div>
                  <div style="padding:14px 16px;font-size:14px;line-height:1.6;">
                    <strong>{esc(titulo)}</strong><br>
                    <span style="color:#6b7280;">{esc(version_texto)}</span>
                    <ul style="margin:10px 0 0 18px;padding:0;color:#374151;">{archivos_html}</ul>
                  </div>
                </div>
                <div style="margin-top:18px;padding:14px 16px;background:#eef2ff;border:1px solid #c7d2fe;border-radius:8px;color:#3730a3;font-size:13px;line-height:1.65;">
                  <strong>Acción requerida</strong><br>
                  Revise esta nueva versión y responda indicando si queda <strong>Aprobada</strong> u <strong>Observada</strong>.
                </div>
                {adicional_html}
              </td></tr>
              <tr><td style="padding:15px 28px;background:#f9fafb;border-top:1px solid #e5e7eb;color:#6b7280;font-size:12px;line-height:1.5;">
                Correo generado automáticamente por <strong>{esc(NOMBRE_APLICACION)}</strong>.
              </td></tr>
            </table>
          </td></tr>
        </table>
      </body>
    </html>
    """
    return asunto, cuerpo_texto, cuerpo_html


def marcar_correo_reenvio_revision_externa_enviado(
    *,
    id_documento: str,
    usuario: str,
    destinatarios: list[str],
    mensaje_adicional: str,
    message_id: str,
    thread_id: str,
    fecha: str,
) -> None:
    """Persiste Gmail antes de crear cabecera/detalle para evitar correos duplicados."""
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "DESTINATARIOS_REVISION_EXTERNA": ", ".join(destinatarios),
                "MENSAJE_ADICIONAL_REVISION_EXTERNA": mensaje_adicional,
                "ENVIADO_REVISION_EXTERNA_POR": usuario,
                "FECHA_ENVIO_REVISION_EXTERNA": fecha,
                "EMAIL_REVISION_EXTERNA_MESSAGE_ID": message_id,
                "GMAIL_THREAD_ID_REVISION_EXTERNA": thread_id,
                "ULTIMO_ENVIADO_POR": usuario,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
            }
        ],
    )


def es_reanudacion_post_gmail_reenvio(documento: dict[str, Any]) -> bool:
    """Distingue el Gmail de este intento de un Gmail histórico de una ronda anterior."""
    message_id = texto(documento.get("EMAIL_REVISION_EXTERNA_MESSAGE_ID"))
    if not message_id:
        return False
    fecha_envio = parsear_fecha_appsheet(documento.get("FECHA_ENVIO_REVISION_EXTERNA"))
    fecha_solicitud = parsear_fecha_appsheet(documento.get("FECHA_ULTIMO_ENVIO"))
    if fecha_envio == datetime.min or fecha_solicitud == datetime.min:
        return False
    return fecha_envio >= fecha_solicitud


def actualizar_documento_tras_reenvio_revision_externa(
    *,
    id_documento: str,
    id_documento_raiz: str,
    id_version: str,
    usuario: str,
    destinatarios: list[str],
    mensaje_adicional: str,
    message_id: str,
    thread_id: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "ESTADO": "En revisión externa",
                "ID_DOCUMENTO_RAIZ_NOTARIAL": id_documento_raiz,
                "RESULTADO_REVISION_EXTERNA": "Pendiente",
                "ID_VERSION_REVISION_EXTERNA": id_version,
                "RESPUESTA_REVISION_EXTERNA_POR": "",
                "FECHA_RESPUESTA_REVISION_EXTERNA": "",
                "DESTINATARIOS_REVISION_EXTERNA": ", ".join(destinatarios),
                "MENSAJE_ADICIONAL_REVISION_EXTERNA": mensaje_adicional,
                "ENVIADO_REVISION_EXTERNA_POR": usuario,
                "FECHA_ENVIO_REVISION_EXTERNA": fecha,
                "EMAIL_REVISION_EXTERNA_MESSAGE_ID": message_id,
                "GMAIL_THREAD_ID_REVISION_EXTERNA": thread_id,
                "ULTIMO_ENVIADO_POR": usuario,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "OBSERVACION_ACTUAL": "",
                "ARCHIVO_RESPUESTA_REVISION_EXTERNA_TEMP": "",
                "ACCION_SOLICITADA": "",
            }
        ],
    )


@app.route("/reenviar-revision-externa", methods=["POST"])
def reenviar_revision_externa():
    """Reenvía únicamente la nueva versión de un documento observado del paquete."""
    id_documento = ""
    correo_enviado = False
    message_id = ""
    thread_id = ""

    try:
        validar_configuracion()
        validar_token()
        data = request.get_json(silent=True) or {}
        id_documento = texto(data.get("id_documento"))
        id_documento_raiz = texto(data.get("id_documento_raiz"))
        id_version = texto(data.get("id_version"))
        usuario = texto(data.get("usuario"))
        destinatarios_entrada = data.get("destinatarios")
        mensaje_adicional = texto(data.get("mensaje_adicional"))

        if not id_documento:
            return {"error": "Falta id_documento"}, 400
        if not id_documento_raiz:
            return {"error": "Falta id_documento_raiz"}, 400

        documento_inicial = buscar_documento(id_documento)
        id_version_actual = texto(documento_inicial.get("ID_VERSION_ACTUAL"))
        id_version_solicitada = id_version or id_version_actual

        # Idempotencia después de una ejecución completamente persistida.
        if (
            texto(documento_inicial.get("ESTADO")) == "En revisión externa"
            and normalizar_resultado_revision_externa(
                documento_inicial.get("RESULTADO_REVISION_EXTERNA")
            ) == "Pendiente"
            and texto(documento_inicial.get("ID_VERSION_REVISION_EXTERNA"))
            == id_version_solicitada
        ):
            pendiente_actual = buscar_revision_externa_pendiente_documento(
                id_documento=id_documento,
                id_version=id_version_solicitada,
            )
            if pendiente_actual is not None:
                revision_actual, _ = pendiente_actual
                raiz_actual = buscar_documento(id_documento_raiz)
                return jsonify(
                    {
                        "ok": True,
                        "ya_procesado": True,
                        "id_documento": id_documento,
                        "id_documento_raiz": id_documento_raiz,
                        "id_revision_externa": texto(
                            revision_actual.get("ID_REVISION_EXTERNA")
                        ),
                        "numero_revision_externa": entero(
                            revision_actual.get("NUMERO_REVISION_EXTERNA") or 0,
                            "NUMERO_REVISION_EXTERNA",
                        ),
                        "id_version_enviada": id_version_solicitada,
                        "estado_paquete_notarial": texto(
                            raiz_actual.get("ESTADO_PAQUETE_NOTARIAL")
                        ),
                        "message_id": texto(
                            documento_inicial.get("EMAIL_REVISION_EXTERNA_MESSAGE_ID")
                        ),
                        "thread_id": texto(
                            documento_inicial.get("GMAIL_THREAD_ID_REVISION_EXTERNA")
                        ),
                    }
                )

        # Recuperación adicional: si ya existe detalle Pendiente para la nueva
        # versión, Gmail/cabecera ya fueron persistidos y solo falta reconciliar
        # el snapshot de Documentos. No se vuelve a enviar correo.
        pendiente_preexistente = buscar_revision_externa_pendiente_documento(
            id_documento=id_documento,
            id_version=id_version_solicitada,
        )

        documento, raiz, version, preparado = (
            validar_documento_para_reenvio_revision_externa_notarial(
                id_documento=id_documento,
                id_documento_raiz=id_documento_raiz,
                id_version=id_version_solicitada,
                permitir_pendiente_version_actual=pendiente_preexistente is not None,
            )
        )
        id_raiz = texto(raiz.get("ID_DOCUMENTO"))

        usuario = usuario or texto(documento.get("ULTIMO_ENVIADO_POR"))
        if not usuario or not _EMAIL_RE.fullmatch(usuario.lower()):
            raise ValueError("El usuario que reenvía a revisión externa no es válido")

        if destinatarios_entrada in (None, ""):
            destinatarios_entrada = (
                documento.get("DESTINATARIOS_REVISION_EXTERNA")
                or raiz.get("DESTINATARIOS_REVISION_EXTERNA")
            )
        destinatarios = normalizar_destinatarios(destinatarios_entrada)
        if not destinatarios:
            raise ValueError("No se indicaron destinatarios para el reenvío externo")

        if not mensaje_adicional:
            mensaje_adicional = texto(
                documento.get("MENSAJE_ADICIONAL_REVISION_EXTERNA")
            ) or texto(raiz.get("MENSAJE_ADICIONAL_REVISION_EXTERNA"))

        recuperando_detalle = pendiente_preexistente is not None
        reanudando_post_gmail = (
            recuperando_detalle
            or es_reanudacion_post_gmail_reenvio(documento)
        )
        message_id_existente = texto(documento.get("EMAIL_REVISION_EXTERNA_MESSAGE_ID"))
        thread_id_existente = texto(documento.get("GMAIL_THREAD_ID_REVISION_EXTERNA"))

        revision_preexistente: dict[str, Any] | None = None
        if recuperando_detalle:
            revision_preexistente = pendiente_preexistente[0]
            message_id_existente = (
                texto(revision_preexistente.get("GMAIL_MESSAGE_ID"))
                or message_id_existente
            )
            thread_id_existente = (
                texto(revision_preexistente.get("GMAIL_THREAD_ID"))
                or thread_id_existente
            )

        fecha_envio = (
            texto((revision_preexistente or {}).get("FECHA_ENVIO"))
            or (
                texto(documento.get("FECHA_ENVIO_REVISION_EXTERNA"))
                if reanudando_post_gmail
                else ""
            )
            or ahora_iso()
        )

        numero_previsto = (
            entero(
                (revision_preexistente or {}).get("NUMERO_REVISION_EXTERNA") or 0,
                "NUMERO_REVISION_EXTERNA",
            )
            if revision_preexistente is not None
            else siguiente_numero_revision_externa(id_raiz)
        )
        cantidad_adjuntos = 0
        resumen_documentos: list[dict[str, Any]] = []

        if reanudando_post_gmail:
            if not message_id_existente:
                raise RuntimeError(
                    "Se detectó una reanudación de reenvío sin GMAIL_MESSAGE_ID"
                )
            correo_enviado = True
            message_id = message_id_existente
            thread_id = thread_id_existente
            # El texto se reconstruye solo para completar MENSAJE_ENVIADO si la
            # persistencia falló después de Gmail.
            archivos = []
            asunto, cuerpo_texto, cuerpo_html = construir_email_reenvio_revision_externa_notarial(
                raiz=raiz,
                documento=documento,
                usuario=usuario,
                mensaje_adicional=mensaje_adicional,
                fecha_envio=fecha_envio,
                numero_revision_externa=numero_previsto,
                archivos=archivos,
            )
        else:
            drive_service = obtener_drive_service()
            gmail_service = obtener_gmail_service()
            adjuntos, resumen_documentos = construir_adjuntos_revision_externa_notarial(
                drive_service=drive_service,
                integrantes_preparados=[preparado],
            )
            cantidad_adjuntos = len(adjuntos)
            archivos = list((resumen_documentos[0] if resumen_documentos else {}).get("archivos") or [])
            asunto, cuerpo_texto, cuerpo_html = construir_email_reenvio_revision_externa_notarial(
                raiz=raiz,
                documento=documento,
                usuario=usuario,
                mensaje_adicional=mensaje_adicional,
                fecha_envio=fecha_envio,
                numero_revision_externa=numero_previsto,
                archivos=archivos,
            )
            respuesta_gmail = enviar_email_con_adjuntos(
                gmail_service=gmail_service,
                destinatarios=destinatarios,
                asunto=asunto,
                cuerpo=cuerpo_texto,
                adjuntos=adjuntos,
                reply_to=usuario,
                cuerpo_html=cuerpo_html,
            )
            correo_enviado = True
            message_id = respuesta_gmail["message_id"]
            thread_id = respuesta_gmail["thread_id"]

            marcar_correo_reenvio_revision_externa_enviado(
                id_documento=id_documento,
                usuario=usuario,
                destinatarios=destinatarios,
                mensaje_adicional=mensaje_adicional,
                message_id=message_id,
                thread_id=thread_id,
                fecha=fecha_envio,
            )

        if revision_preexistente is not None:
            id_revision_externa = texto(
                revision_preexistente.get("ID_REVISION_EXTERNA")
            )
            numero_revision_externa = entero(
                revision_preexistente.get("NUMERO_REVISION_EXTERNA") or 0,
                "NUMERO_REVISION_EXTERNA",
            )
            cabecera_creada = False
        else:
            id_revision_externa, numero_revision_externa, cabecera_creada = (
                obtener_o_crear_revision_externa_cabecera(
                    id_documento_raiz=id_raiz,
                    usuario=usuario,
                    destinatarios=destinatarios,
                    fecha=fecha_envio,
                    message_id=message_id,
                    thread_id=thread_id,
                    mensaje_enviado=cuerpo_texto,
                )
            )

        # Si el número real cambió por concurrencia, la BD es la fuente de verdad;
        # el correo ya enviado conserva el número previsto, pero no se duplica.
        detalles_creados = crear_detalles_revision_externa_faltantes(
            id_revision_externa=id_revision_externa,
            integrantes_preparados=[preparado],
        )

        actualizar_documento_tras_reenvio_revision_externa(
            id_documento=id_documento,
            id_documento_raiz=id_raiz,
            id_version=id_version_solicitada,
            usuario=usuario,
            destinatarios=destinatarios,
            mensaje_adicional=mensaje_adicional,
            message_id=message_id,
            thread_id=thread_id,
            fecha=fecha_envio,
        )

        estados = recalcular_revision_y_paquete_notarial(
            id_revision_externa=id_revision_externa,
            id_documento_raiz=id_raiz,
            usuario=usuario,
            fecha=fecha_envio,
        )

        advertencias: list[str] = []
        try:
            crear_eventos_envio_revision_externa(
                integrantes_preparados=[preparado],
                id_revision_externa=id_revision_externa,
                numero_revision_externa=numero_revision_externa,
                usuario=usuario,
                destinatarios=destinatarios,
                fecha=fecha_envio,
            )
        except Exception as exc_evento:
            traceback.print_exc()
            advertencias.append(
                "El reenvío quedó registrado, pero falló la creación del evento: "
                f"{exc_evento}"
            )

        return jsonify(
            {
                "ok": True,
                "ya_procesado": False,
                "reanudado_post_gmail": reanudando_post_gmail,
                "id_documento": id_documento,
                "id_documento_raiz": id_raiz,
                "id_revision_externa": id_revision_externa,
                "numero_revision_externa": numero_revision_externa,
                "id_version_enviada": id_version_solicitada,
                **estados,
                "message_id": message_id,
                "thread_id": thread_id,
                "destinatarios": destinatarios,
                "cantidad_adjuntos": cantidad_adjuntos,
                "cabecera_creada": cabecera_creada,
                "detalles_creados": detalles_creados,
                "advertencias": advertencias,
            }
        )

    except PermissionError as exc:
        if id_documento:
            registrar_error_envio_revision_externa(id_documento, str(exc))
        return {"error": str(exc)}, 403
    except (ValueError, LookupError) as exc:
        if id_documento:
            registrar_error_envio_revision_externa(id_documento, str(exc))
        return {"error": str(exc)}, 400
    except Exception as exc:
        traceback.print_exc()
        if id_documento:
            registrar_error_envio_revision_externa(id_documento, str(exc))
        return {
            "error": str(exc),
            "correo_enviado": correo_enviado,
            "message_id": message_id,
            "thread_id": thread_id,
        }, 500


# -----------------------------------------------------------------------------
# Flujo Notarial - registrar respuesta de revisión externa
# -----------------------------------------------------------------------------


RESULTADOS_REVISION_EXTERNA_VALIDOS = {"Pendiente", "Aprobado", "Observado"}
ESTADOS_PAQUETE_NOTARIAL_ACTIVOS = {"En revisión externa", "Con observaciones", "Listo para notaría"}


def normalizar_resultado_revision_externa(valor: Any) -> str:
    valor_texto = texto(valor).strip()
    if not valor_texto:
        return ""
    equivalencias = {
        "pendiente": "Pendiente",
        "aprobado": "Aprobado",
        "observado": "Observado",
    }
    clave = valor_texto.casefold()
    if clave not in equivalencias:
        raise ValueError(
            "RESULTADO_REVISION_EXTERNA contiene un valor no reconocido: "
            f"{valor_texto!r}"
        )
    return equivalencias[clave]


def buscar_revision_externa_por_id(id_revision_externa: str) -> dict[str, Any]:
    selector = (
        f"FILTER({TABLA_REVISIONES_EXTERNAS}, "
        f"[ID_REVISION_EXTERNA] = {literal_appsheet(id_revision_externa)})"
    )
    filas = appsheet_find(TABLA_REVISIONES_EXTERNAS, selector)
    if not filas:
        raise LookupError(
            f"No se encontró ID_REVISION_EXTERNA={id_revision_externa}"
        )
    return filas[0]


def buscar_detalle_revision_externa_documento(
    *,
    id_revision_externa: str,
    id_documento: str,
) -> dict[str, Any]:
    selector = (
        f"FILTER({TABLA_REVISION_EXTERNA_DETALLE}, "
        f"AND([ID_REVISION_EXTERNA] = {literal_appsheet(id_revision_externa)}, "
        f"[ID_DOCUMENTO] = {literal_appsheet(id_documento)}))"
    )
    filas = appsheet_find(TABLA_REVISION_EXTERNA_DETALLE, selector)
    if not filas:
        raise LookupError(
            "No se encontró el detalle de revisión externa para "
            f"ID_REVISION_EXTERNA={id_revision_externa}, "
            f"ID_DOCUMENTO={id_documento}"
        )
    if len(filas) > 1:
        raise ValueError(
            "Existe más de un detalle para el mismo documento en la misma "
            "revisión externa"
        )
    return filas[0]


def buscar_integrantes_paquete_notarial(id_documento_raiz: str) -> list[dict[str, Any]]:
    selector = (
        f"FILTER({TABLA_DOCUMENTOS}, "
        f"[ID_DOCUMENTO_RAIZ_NOTARIAL] = {literal_appsheet(id_documento_raiz)})"
    )
    filas = appsheet_find(TABLA_DOCUMENTOS, selector)
    filas.sort(
        key=lambda fila: (
            0 if texto(fila.get("ID_DOCUMENTO")) == id_documento_raiz else 1,
            texto(fila.get("TITULO")).casefold(),
            texto(fila.get("ID_DOCUMENTO")),
        )
    )
    return filas


def nombre_respaldo_revision_externa(valor: Any) -> str:
    ruta = normalizar_archivo_observacion_appsheet(texto(valor))
    if not ruta:
        return ""
    return PurePosixPath(ruta).name.strip()


def validar_contexto_respuesta_revision_externa(
    *,
    id_documento: str,
    id_documento_raiz: str,
    id_revision_externa: str,
    id_version: str,
) -> tuple[
    dict[str, Any],
    dict[str, Any],
    dict[str, Any],
    dict[str, Any],
]:
    documento = buscar_documento(id_documento)
    id_raiz_documento = texto(documento.get("ID_DOCUMENTO_RAIZ_NOTARIAL"))
    if not id_raiz_documento:
        raise ValueError("El documento no pertenece a un paquete Notarial activo")
    if id_documento_raiz and id_documento_raiz != id_raiz_documento:
        raise ValueError("id_documento_raiz no coincide con el paquete Notarial activo")

    raiz = buscar_documento(id_raiz_documento)
    if normalizar_tipo_firma(raiz.get("TIPO_FIRMA")) != "Notarial":
        raise ValueError("La raíz del paquete no corresponde a Firma Notarial")

    estado_paquete = texto(raiz.get("ESTADO_PAQUETE_NOTARIAL"))
    if estado_paquete not in ESTADOS_PAQUETE_NOTARIAL_ACTIVOS:
        raise ValueError(
            "El paquete Notarial no admite respuestas externas en su estado actual: "
            f"{estado_paquete!r}"
        )

    revision = buscar_revision_externa_por_id(id_revision_externa)
    if texto(revision.get("ID_DOCUMENTO_RAIZ")) != id_raiz_documento:
        raise ValueError("La revisión externa indicada pertenece a otro paquete")

    detalle = buscar_detalle_revision_externa_documento(
        id_revision_externa=id_revision_externa,
        id_documento=id_documento,
    )

    id_version_enviada = texto(detalle.get("ID_VERSION_ENVIADA"))
    id_version_documento = texto(documento.get("ID_VERSION_REVISION_EXTERNA"))
    id_version_solicitada = id_version or id_version_documento or id_version_enviada
    if not id_version_solicitada:
        raise ValueError("No se pudo identificar la versión enviada a revisión externa")
    if id_version_enviada and id_version_solicitada != id_version_enviada:
        raise ValueError(
            "La respuesta corresponde a una versión distinta de la enviada "
            "en esta revisión externa"
        )
    if id_version_documento and id_version_documento != id_version_solicitada:
        raise ValueError(
            "ID_VERSION_REVISION_EXTERNA no coincide con la versión respondida"
        )

    # La evidencia de respuesta ya no es obligatoria en la cabecera de la ronda.
    # Cada documento conserva su propio archivo en Documento_Revision_Externa_Detalle.
    return documento, raiz, revision, detalle


def validar_archivo_respaldo_revision_externa(valor: Any) -> tuple[str, str]:
    """
    Normaliza el File de AppSheet usado como evidencia de una respuesta externa.

    El backend conserva la referencia histórica en el detalle de la ronda; no usa
    este archivo para crear el nuevo borrador. El reinicio siempre parte del
    GOOGLE_DOC_ID de la versión que efectivamente fue enviada a revisión externa.
    """
    ruta = normalizar_archivo_observacion_appsheet(texto(valor))
    if not ruta:
        return "", ""

    nombre = nombre_archivo_desde_valor_appsheet(ruta)
    if not nombre:
        raise ValueError(
            "No se pudo determinar el nombre del archivo de respaldo de revisión externa"
        )
    return ruta, nombre


def actualizar_detalle_revision_externa(
    *,
    detalle: dict[str, Any],
    resultado: str,
    comentario: str,
    usuario: str,
    fecha: str,
    archivo_respaldo: str = "",
) -> tuple[str, str]:
    # En reintentos, si AppSheet ya limpió el campo TEMP, conservamos el archivo
    # histórico que pudo haberse grabado en un intento anterior.
    archivo_existente = texto(detalle.get("ARCHIVO_RESPALDO_RESPUESTA"))
    valor_archivo = archivo_respaldo or archivo_existente
    ruta_archivo, nombre_archivo = validar_archivo_respaldo_revision_externa(
        valor_archivo
    )

    cambios: dict[str, Any] = {
        "ID_REVISION_EXTERNA_DETALLE": texto(
            detalle.get("ID_REVISION_EXTERNA_DETALLE")
        ),
        "RESULTADO": resultado,
        "COMENTARIO": comentario,
        "FECHA_RESPUESTA": fecha,
        "REGISTRADO_POR": usuario,
    }
    if ruta_archivo:
        cambios.update(
            {
                "ARCHIVO_RESPALDO_RESPUESTA": ruta_archivo,
                "NOMBRE_ARCHIVO_RESPALDO": nombre_archivo,
            }
        )

    appsheet_action(
        TABLA_REVISION_EXTERNA_DETALLE,
        "Edit",
        [cambios],
    )
    return ruta_archivo, nombre_archivo


def actualizar_documento_aprobado_revision_externa(
    *,
    id_documento: str,
    id_version: str,
    usuario: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "RESULTADO_REVISION_EXTERNA": "Aprobado",
                "ID_VERSION_REVISION_EXTERNA": id_version,
                "RESPUESTA_REVISION_EXTERNA_POR": usuario,
                "FECHA_RESPUESTA_REVISION_EXTERNA": fecha,
                "ULTIMO_ENVIADO_POR": usuario,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "OBSERVACION_ACTUAL": "",
                "ARCHIVO_RESPUESTA_REVISION_EXTERNA_TEMP": "",
                "ACCION_SOLICITADA": "",
            }
        ],
    )


def cerrar_gestion_revision_externa_anterior_por_observacion(
    *,
    cadena_anterior: list[dict[str, Any]],
    comentario: str,
    fecha: str,
) -> None:
    """Desactiva la cadena de la versión observada externamente."""
    filas: list[dict[str, Any]] = []
    for fila in cadena_anterior:
        if not es_verdadero(fila.get("CADENA_ACTIVA")):
            continue
        cambios: dict[str, Any] = {
            "ID_APROBACION_ACTUAL": texto(fila.get("ID_APROBACION_ACTUAL")),
            "CADENA_ACTIVA": False,
        }
        if es_responsable_firmas(fila):
            cambios.update(
                {
                    "ESTADO": "Cerrado",
                    "RESULTADO": "Observado en revisión externa",
                    "COMENTARIO": comentario,
                    "FECHA_RESPUESTA": fecha,
                }
            )
        filas.append(cambios)
    if filas:
        appsheet_action(TABLA_APROBADORES_ACTUAL, "Edit", filas)


def actualizar_documento_reinicio_revision_externa(
    *,
    id_documento: str,
    id_documento_raiz: str,
    numero_version: int,
    id_version_nueva: str,
    id_version_observada: str,
    copia: dict[str, str],
    primer_encargado: dict[str, Any],
    usuario: str,
    fecha: str,
) -> None:
    """Reinicia solo el documento observado y conserva el paquete Notarial."""
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "ESTADO": "Borrador",
                "VERSION_ACTUAL": numero_version,
                "REVISION_ACTUAL": 0,
                "ID_VERSION_ACTUAL": id_version_nueva,
                "GOOGLE_DOC_ID": copia["id"],
                "GOOGLE_DOC_URL": copia["url"],
                "ORDEN_ACTUAL": primer_encargado["ORDEN"],
                "ID_APROBACION_ACTUAL": primer_encargado[
                    "ID_APROBACION_ACTUAL"
                ],
                "ENCARGADO_ACTUAL_NOMBRE": primer_encargado.get("NOMBRE", ""),
                "ENCARGADO_ACTUAL_EMAIL": primer_encargado.get("APROBADOR", ""),
                "ID_DOCUMENTO_RAIZ_NOTARIAL": id_documento_raiz,
                "RESULTADO_REVISION_EXTERNA": "Observado",
                "ID_VERSION_REVISION_EXTERNA": id_version_observada,
                "RESPUESTA_REVISION_EXTERNA_POR": usuario,
                "FECHA_RESPUESTA_REVISION_EXTERNA": fecha,
                "FECHA_CIERRE": "",
                "ULTIMO_ENVIADO_POR": usuario,
                # FECHA_ULTIMO_ENVIO queda reservada para acciones iniciadas
                # desde AppSheet, evitando disparar nuevamente Bots internos.
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "OBSERVACION_ACTUAL": "",
                "ARCHIVO_RESPUESTA_REVISION_EXTERNA_TEMP": "",
                "ACCION_SOLICITADA": "",
            }
        ],
    )


def buscar_evento_respuesta_revision_externa(
    *,
    id_documento: str,
    id_version: str,
    resultado: str,
    id_revision_externa: str,
) -> dict[str, Any] | None:
    """Recupera el evento único de respuesta externa para reintentos idempotentes."""
    prefijo = f"Revisión externa {id_revision_externa}:"
    candidatos = [
        evento
        for evento in buscar_eventos_documento(id_documento)
        if texto(evento.get("TIPO_EVENTO")) == "Respuesta revisión externa"
        and texto(evento.get("ID_VERSION")) == id_version
        and texto(evento.get("ESTADO_NUEVO")) == resultado
        and texto(evento.get("COMENTARIO")).startswith(prefijo)
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


def crear_evento_respuesta_revision_externa(
    *,
    id_documento: str,
    id_version: str,
    usuario: str,
    fecha: str,
    resultado: str,
    comentario: str,
    id_revision_externa: str,
) -> dict[str, Any]:
    existente = buscar_evento_respuesta_revision_externa(
        id_documento=id_documento,
        id_version=id_version,
        resultado=resultado,
        id_revision_externa=id_revision_externa,
    )
    if existente is not None:
        return existente

    evento = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION": id_version,
        "ID_APROBACION_ACTUAL": "",
        "TIPO_EVENTO": "Respuesta revisión externa",
        "ESTADO_ANTERIOR": "Pendiente",
        "ESTADO_NUEVO": resultado,
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": (
            f"Revisión externa {id_revision_externa}: {resultado}."
            + (f" {comentario}" if comentario else "")
        ),
    }
    appsheet_action(TABLA_EVENTOS, "Add", [evento])
    return evento


def crear_eventos_reinicio_revision_externa(
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
) -> dict[str, Any] | None:
    evento_reinicio = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION": id_version_observada,
        "ID_APROBACION_ACTUAL": "",
        "TIPO_EVENTO": "Proceso reiniciado por revisión externa",
        "ESTADO_ANTERIOR": "En revisión externa",
        "ESTADO_NUEVO": "Borrador",
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": (
            "El documento recibió observaciones en revisión externa. "
            f"Motivo: {comentario}"
        ),
    }
    evento_version = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION": id_version_nueva,
        "ID_APROBACION_ACTUAL": id_aprobacion_nueva,
        "TIPO_EVENTO": "Nueva versión creada",
        "ESTADO_ANTERIOR": "Observado en revisión externa",
        "ESTADO_NUEVO": "Borrador",
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": (
            f"Se creó la versión {numero_version_nueva}: {nombre_archivo}. "
            "La aprobación interna se reinició desde el primer responsable."
        ),
    }
    appsheet_action(TABLA_EVENTOS, "Add", [evento_reinicio, evento_version])
    return evento_reinicio


def construir_especificaciones_reinicio_revision_externa(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
    cadena: list[dict[str, Any]],
    aprobador_destino: dict[str, Any],
    responsable_firmas: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    comentario_evento = texto(evento.get("COMENTARIO"))
    especificaciones: list[dict[str, Any]] = [
        {
            "aprobador": aprobador_destino,
            "tipo_notificacion": "Acción requerida",
            "movimiento": "Documento devuelto por revisión externa",
            "comentario_principal": comentario_evento,
            "link_documento": normalizar_url_appsheet(
                documento.get("GOOGLE_DOC_URL")
            ),
        }
    ]
    if responsable_firmas is not None:
        especificaciones.append(
            {
                "aprobador": responsable_firmas,
                "tipo_notificacion": "Confirmación",
                "movimiento": "Observación de revisión externa registrada",
                "comentario_principal": comentario_evento,
                "link_documento": "",
            }
        )
    ids_excluidos = {
        texto(aprobador_destino.get("ID_APROBACION_ACTUAL")),
        texto(responsable_firmas.get("ID_APROBACION_ACTUAL"))
        if responsable_firmas is not None
        else "",
    }
    for integrante in cadena:
        if texto(integrante.get("ID_APROBACION_ACTUAL")) in ids_excluidos:
            continue
        especificaciones.append(
            {
                "aprobador": integrante,
                "tipo_notificacion": "Informativa",
                "movimiento": "El documento volvió a aprobación interna",
                "comentario_principal": comentario_evento,
                "link_documento": "",
            }
        )
    return especificaciones


def obtener_responsable_firmas_documento_notarial(
    documento: dict[str, Any],
) -> dict[str, Any]:
    """Obtiene el Responsable de firmas de la versión vigente del documento."""
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    numero_version = entero(documento.get("VERSION_ACTUAL"), "VERSION_ACTUAL")
    cadena = buscar_cadena_documento_version(id_documento, numero_version)
    if not cadena:
        raise LookupError(
            f"No se encontró cadena para notificar el documento {id_documento} "
            f"versión {numero_version}"
        )
    return obtener_responsable_firmas_cadena(
        cadena,
        contexto=f"Documento {id_documento} versión {numero_version}",
    )


def notificar_aprobacion_revision_externa_documento(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
) -> list[dict[str, Any]]:
    responsable = obtener_responsable_firmas_documento_notarial(documento)
    comentario_evento = (
        texto(evento.get("COMENTARIO"))
        or "El documento fue aprobado en revisión externa."
    )
    return notificar_destinatarios_internos(
        documento=documento,
        evento=evento,
        destinatarios=[
            {
                "aprobador": responsable,
                "tipo_notificacion": "Informativa",
                "movimiento": "Documento aprobado en revisión externa",
                "comentario_principal": comentario_evento,
                "link_documento": "",
            }
        ],
    )


def buscar_evento_reinicio_revision_externa(
    *,
    id_documento: str,
    id_version_observada: str,
) -> dict[str, Any] | None:
    candidatos = [
        evento
        for evento in buscar_eventos_documento(id_documento)
        if texto(evento.get("TIPO_EVENTO"))
        == "Proceso reiniciado por revisión externa"
        and texto(evento.get("ID_VERSION")) == id_version_observada
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


def reanudar_notificaciones_observacion_revision_externa(
    *,
    id_documento: str,
    id_version_observada: str,
) -> list[dict[str, Any]]:
    """Reintenta las notificaciones del reinicio usando el mismo ID_EVENTO."""
    evento = buscar_evento_reinicio_revision_externa(
        id_documento=id_documento,
        id_version_observada=id_version_observada,
    )
    if evento is None:
        return []

    documento = buscar_documento(id_documento)
    numero_version = entero(documento.get("VERSION_ACTUAL"), "VERSION_ACTUAL")
    cadena = buscar_cadena_documento_version(id_documento, numero_version)
    if not cadena:
        raise LookupError(
            f"No se encontró la cadena reiniciada del documento {id_documento}"
        )
    primer_responsable = obtener_primer_responsable_aprobacion(
        cadena,
        contexto=f"Documento {id_documento} versión {numero_version}",
    )
    responsable_firmas = obtener_responsable_firmas_cadena(
        cadena,
        contexto=f"Documento {id_documento} versión {numero_version}",
    )
    especificaciones = construir_especificaciones_reinicio_revision_externa(
        documento=documento,
        evento=evento,
        cadena=cadena,
        aprobador_destino=primer_responsable,
        responsable_firmas=responsable_firmas,
    )
    return notificar_destinatarios_internos(
        documento=documento,
        evento=evento,
        destinatarios=especificaciones,
    )


def buscar_evento_estado_paquete_notarial(
    *,
    id_documento_raiz: str,
    id_revision_externa: str,
    estado_nuevo: str,
) -> dict[str, Any] | None:
    prefijo = f"Revisión externa {id_revision_externa}: paquete "
    candidatos = [
        evento
        for evento in buscar_eventos_documento(id_documento_raiz)
        if texto(evento.get("TIPO_EVENTO")) == "Estado paquete notarial"
        and texto(evento.get("ESTADO_NUEVO")) == estado_nuevo
        and texto(evento.get("COMENTARIO")).startswith(prefijo)
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


def asegurar_evento_estado_paquete_notarial(
    *,
    raiz: dict[str, Any],
    id_revision_externa: str,
    estado_anterior: str,
    estado_nuevo: str,
    usuario: str,
    fecha: str,
) -> dict[str, Any] | None:
    """Crea una sola evidencia por estado relevante de cada ronda notarial."""
    if estado_nuevo not in {"Con observaciones", "Listo para notaría"}:
        return None

    id_raiz = texto(raiz.get("ID_DOCUMENTO"))
    existente = buscar_evento_estado_paquete_notarial(
        id_documento_raiz=id_raiz,
        id_revision_externa=id_revision_externa,
        estado_nuevo=estado_nuevo,
    )
    if existente is not None:
        return existente

    if estado_nuevo == "Con observaciones":
        comentario = (
            f"Revisión externa {id_revision_externa}: paquete con observaciones. "
            "Existe al menos un documento observado que debe volver a aprobación interna."
        )
    else:
        comentario = (
            f"Revisión externa {id_revision_externa}: paquete listo para notaría. "
            "Todos los documentos del paquete fueron aprobados en revisión externa."
        )

    evento = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_raiz,
        "ID_VERSION": texto(raiz.get("ID_VERSION_ACTUAL")),
        "ID_APROBACION_ACTUAL": texto(raiz.get("ID_APROBACION_ACTUAL")),
        "TIPO_EVENTO": "Estado paquete notarial",
        "ESTADO_ANTERIOR": estado_anterior,
        "ESTADO_NUEVO": estado_nuevo,
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": comentario,
    }
    appsheet_action(TABLA_EVENTOS, "Add", [evento])
    return evento


def notificar_estado_paquete_notarial(
    *,
    raiz: dict[str, Any],
    evento: dict[str, Any],
) -> list[dict[str, Any]]:
    responsable = obtener_responsable_firmas_documento_notarial(raiz)
    estado_nuevo = texto(evento.get("ESTADO_NUEVO"))
    if estado_nuevo == "Listo para notaría":
        tipo = "Acción requerida"
        movimiento = "Paquete listo para enviar a notaría"
        link_documento = normalizar_url_appsheet(raiz.get("GOOGLE_DOC_URL"))
    else:
        tipo = "Informativa"
        movimiento = "Paquete notarial con observaciones"
        link_documento = ""

    return notificar_destinatarios_internos(
        documento=raiz,
        evento=evento,
        destinatarios=[
            {
                "aprobador": responsable,
                "tipo_notificacion": tipo,
                "movimiento": movimiento,
                "comentario_principal": texto(evento.get("COMENTARIO")),
                "link_documento": link_documento,
            }
        ],
    )


def ejecutar_notificaciones_post_respuesta_revision_externa(
    *,
    documento: dict[str, Any],
    raiz_antes: dict[str, Any],
    id_revision_externa: str,
    id_version_enviada: str,
    resultado: str,
    comentario: str,
    usuario: str,
    fecha: str,
    estado_paquete_nuevo: str,
) -> tuple[list[dict[str, Any]], list[str]]:
    """Notifica cambios externos sin revertir la transición si Gmail falla."""
    resultados: list[dict[str, Any]] = []
    advertencias: list[str] = []

    try:
        evento_respuesta = crear_evento_respuesta_revision_externa(
            id_documento=texto(documento.get("ID_DOCUMENTO")),
            id_version=id_version_enviada,
            usuario=usuario,
            fecha=fecha,
            resultado=resultado,
            comentario=comentario,
            id_revision_externa=id_revision_externa,
        )
        if resultado == "Aprobado":
            documento_actualizado = buscar_documento(
                texto(documento.get("ID_DOCUMENTO"))
            )
            notifs = notificar_aprobacion_revision_externa_documento(
                documento=documento_actualizado,
                evento=evento_respuesta,
            )
            resultados.extend(notifs)
            fallidas = [fila for fila in notifs if not fila.get("ok")]
            if fallidas:
                advertencias.append(
                    f"{len(fallidas)} notificación(es) de aprobación externa "
                    "quedaron omitidas o con error. Revisa Documento_Notificaciones."
                )
        elif resultado == "Observado":
            notifs = reanudar_notificaciones_observacion_revision_externa(
                id_documento=texto(documento.get("ID_DOCUMENTO")),
                id_version_observada=id_version_enviada,
            )
            resultados.extend(notifs)
            fallidas = [fila for fila in notifs if not fila.get("ok")]
            if fallidas:
                advertencias.append(
                    f"{len(fallidas)} notificación(es) del documento observado "
                    "quedaron omitidas o con error. Revisa Documento_Notificaciones."
                )
    except Exception as exc:
        traceback.print_exc()
        advertencias.append(
            "La respuesta externa quedó registrada, pero falló su notificación interna: "
            + str(exc)
        )

    try:
        raiz_actual = buscar_documento(texto(raiz_antes.get("ID_DOCUMENTO")))
        estado_paquete_anterior = texto(raiz_antes.get("ESTADO_PAQUETE_NOTARIAL"))
        evento_paquete = None
        if estado_paquete_nuevo != estado_paquete_anterior:
            evento_paquete = asegurar_evento_estado_paquete_notarial(
                raiz=raiz_actual,
                id_revision_externa=id_revision_externa,
                estado_anterior=estado_paquete_anterior,
                estado_nuevo=estado_paquete_nuevo,
                usuario=usuario,
                fecha=fecha,
            )
        if evento_paquete is not None:
            notifs = notificar_estado_paquete_notarial(
                raiz=raiz_actual,
                evento=evento_paquete,
            )
            resultados.extend(notifs)
            fallidas = [fila for fila in notifs if not fila.get("ok")]
            if fallidas:
                advertencias.append(
                    f"{len(fallidas)} notificación(es) del estado del paquete "
                    "quedaron omitidas o con error. Revisa Documento_Notificaciones."
                )
    except Exception as exc:
        traceback.print_exc()
        advertencias.append(
            "El estado del paquete quedó actualizado, pero falló su notificación interna: "
            + str(exc)
        )

    return resultados, advertencias


def recalcular_revision_y_paquete_notarial(
    *,
    id_revision_externa: str,
    id_documento_raiz: str,
    usuario: str,
    fecha: str,
) -> dict[str, str]:
    """
    Cierra/actualiza la ronda indicada usando sus detalles y calcula el estado
    global del paquete usando el snapshot operativo de TODOS sus documentos.

    Esto permite múltiples rondas parciales: una nueva ronda puede contener solo
    un documento corregido sin perder las aprobaciones externas de los demás.
    """
    detalles = buscar_detalles_revision_externa(id_revision_externa)
    if not detalles:
        raise LookupError("La revisión externa no contiene detalles")

    resultados_ronda = [
        normalizar_resultado_revision_externa(fila.get("RESULTADO"))
        for fila in detalles
    ]
    if any(not resultado for resultado in resultados_ronda):
        raise ValueError("Existe un detalle de revisión externa sin RESULTADO")

    hay_pendientes_ronda = any(
        resultado == "Pendiente" for resultado in resultados_ronda
    )
    hay_observados_ronda = any(
        resultado == "Observado" for resultado in resultados_ronda
    )

    if hay_pendientes_ronda:
        estado_revision = "Pendiente"
    elif hay_observados_ronda:
        estado_revision = "Observada"
    elif all(resultado == "Aprobado" for resultado in resultados_ronda):
        estado_revision = "Aprobada"
    else:
        raise ValueError("No fue posible determinar ESTADO_REVISION")

    cambios_revision: dict[str, Any] = {
        "ID_REVISION_EXTERNA": id_revision_externa,
        "ESTADO_REVISION": estado_revision,
    }
    if estado_revision != "Pendiente":
        cambios_revision.update(
            {
                "FECHA_RESPUESTA": fecha,
                "REGISTRADO_POR": usuario,
            }
        )
    appsheet_action(TABLA_REVISIONES_EXTERNAS, "Edit", [cambios_revision])

    integrantes = buscar_integrantes_paquete_notarial(id_documento_raiz)
    if not integrantes:
        raise LookupError("El paquete Notarial no contiene documentos asociados")

    resultados_globales = [
        normalizar_resultado_revision_externa(
            fila.get("RESULTADO_REVISION_EXTERNA")
        )
        for fila in integrantes
    ]
    if any(not resultado for resultado in resultados_globales):
        faltantes = [
            texto(fila.get("TITULO")) or texto(fila.get("ID_DOCUMENTO"))
            for fila, resultado in zip(integrantes, resultados_globales)
            if not resultado
        ]
        raise ValueError(
            "Existen documentos del paquete sin RESULTADO_REVISION_EXTERNA: "
            + ", ".join(faltantes)
        )

    hay_observados_global = any(
        resultado == "Observado" for resultado in resultados_globales
    )
    hay_pendientes_global = any(
        resultado == "Pendiente" for resultado in resultados_globales
    )

    if hay_observados_global:
        estado_paquete = "Con observaciones"
    elif hay_pendientes_global:
        estado_paquete = "En revisión externa"
    elif all(resultado == "Aprobado" for resultado in resultados_globales):
        estado_paquete = "Listo para notaría"
    else:
        raise ValueError("No fue posible determinar ESTADO_PAQUETE_NOTARIAL")

    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento_raiz,
                "ESTADO_PAQUETE_NOTARIAL": estado_paquete,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
            }
        ],
    )

    if estado_paquete == "Listo para notaría":
        filas_estado: list[dict[str, Any]] = []
        for fila in integrantes:
            filas_estado.append(
                {
                    "ID_DOCUMENTO": texto(fila.get("ID_DOCUMENTO")),
                    "ESTADO": "Listo para notaría",
                    "FECHA_ULTIMA_ACTUALIZACION": fecha,
                }
            )
        if filas_estado:
            appsheet_action(TABLA_DOCUMENTOS, "Edit", filas_estado)

    return {
        "estado_revision": estado_revision,
        "estado_paquete_notarial": estado_paquete,
    }

def procesar_observacion_revision_externa(
    *,
    documento: dict[str, Any],
    raiz: dict[str, Any],
    revision: dict[str, Any],
    detalle: dict[str, Any],
    id_version_observada: str,
    usuario: str,
    comentario: str,
    fecha: str,
    archivo_respaldo: str = "",
) -> dict[str, Any]:
    if not comentario:
        raise ValueError("El comentario es obligatorio al observar una revisión externa")

    archivo_existente = texto(detalle.get("ARCHIVO_RESPALDO_RESPUESTA"))
    archivo_efectivo = archivo_respaldo or archivo_existente
    if not normalizar_archivo_observacion_appsheet(archivo_efectivo):
        raise ValueError(
            "El archivo de respaldo es obligatorio al observar una revisión externa"
        )

    id_documento = texto(documento.get("ID_DOCUMENTO"))
    id_documento_raiz = texto(raiz.get("ID_DOCUMENTO"))
    version_observada = buscar_version_por_id(id_version_observada)
    if texto(version_observada.get("ID_DOCUMENTO")) != id_documento:
        raise ValueError("La versión observada pertenece a otro documento")
    if texto(version_observada.get("ETAPA")) != "Para revisión externa":
        raise ValueError(
            "La versión respondida no corresponde a ETAPA='Para revisión externa'"
        )

    numero_version_anterior = entero(
        version_observada.get("NUMERO_VERSION"),
        "NUMERO_VERSION",
    )
    numero_version_documento = entero(
        documento.get("VERSION_ACTUAL"),
        "VERSION_ACTUAL",
    )
    resultado_documento = normalizar_resultado_revision_externa(
        documento.get("RESULTADO_REVISION_EXTERNA")
    )

    # Reanudación idempotente: el documento ya fue reiniciado, pero pudo faltar
    # la actualización del detalle/cabecera por una falla posterior.
    if (
        numero_version_documento > numero_version_anterior
        and resultado_documento == "Observado"
        and texto(documento.get("ID_VERSION_REVISION_EXTERNA"))
        == id_version_observada
    ):
        actualizar_detalle_revision_externa(
            detalle=detalle,
            resultado="Observado",
            comentario=comentario,
            usuario=usuario,
            fecha=fecha,
            archivo_respaldo=archivo_efectivo,
        )
        return {
            "ya_reiniciado": True,
            "id_version_nueva": texto(documento.get("ID_VERSION_ACTUAL")),
            "numero_version_nueva": numero_version_documento,
            "notificaciones": [],
            "advertencias": [],
        }

    if texto(documento.get("ID_VERSION_ACTUAL")) != id_version_observada:
        raise ValueError(
            "La versión actual ya no coincide con la versión enviada a revisión externa"
        )

    google_doc_id_origen = texto(version_observada.get("GOOGLE_DOC_ID"))
    if not google_doc_id_origen:
        raise ValueError(
            "La versión observada no tiene GOOGLE_DOC_ID para crear el nuevo borrador"
        )

    cadena_anterior = buscar_cadena_documento_version(
        id_documento=id_documento,
        numero_version=numero_version_anterior,
    )
    if not cadena_anterior:
        raise ValueError(
            "No se encontró la cadena de aprobación de la versión enviada a revisión externa"
        )

    primer_anterior = obtener_primer_responsable_aprobacion(
        cadena_anterior,
        contexto=f"Documento {id_documento} versión {numero_version_anterior}",
    )
    responsable_firmas_anterior = obtener_responsable_firmas_cadena(
        cadena_anterior,
        contexto=f"Documento {id_documento} versión {numero_version_anterior}",
    )

    id_plantilla = texto(documento.get("ID_PLANTILLA"))
    plantilla = buscar_plantilla(id_plantilla)
    folder_id = texto(plantilla.get("CARPETA_DESTINO_ID"))
    if not folder_id:
        raise ValueError("La plantilla no tiene CARPETA_DESTINO_ID")

    numero_version_nueva = numero_version_anterior + 1
    nombre_archivo = limpiar_nombre_archivo(
        f"{texto(documento.get('TITULO')) or f'Documento_{id_documento}'}_"
        f"V{numero_version_nueva:02d}_BORRADOR"
    )
    drive_service = obtener_drive_service()

    version_existente = buscar_version_numero_revision(
        id_documento=id_documento,
        numero_version=numero_version_nueva,
        numero_revision=0,
    )
    if version_existente:
        id_version_nueva = texto(version_existente.get("ID_VERSION"))
        copia = {
            "id": texto(version_existente.get("GOOGLE_DOC_ID")),
            "url": normalizar_url_appsheet(version_existente.get("GOOGLE_DOC_URL")),
            "name": texto(version_existente.get("NOMBRE_ARCHIVO")) or nombre_archivo,
        }
        if not copia["id"]:
            raise RuntimeError("La nueva versión existente no tiene GOOGLE_DOC_ID")
    else:
        id_version_nueva = nuevo_id()
        copia = copiar_archivo_o_reutilizar(
            drive_service=drive_service,
            source_file_id=google_doc_id_origen,
            folder_id=folder_id,
            nombre_archivo=nombre_archivo,
        )

    # Congela la versión observada para todos los integrantes de la cadena.
    for fila in cadena_anterior:
        email = texto(fila.get("APROBADOR"))
        if email:
            asegurar_permiso_rol(
                drive_service=drive_service,
                file_id=google_doc_id_origen,
                email=email,
                role="reader",
            )

    email_primero = texto(primer_anterior.get("APROBADOR"))
    if not email_primero:
        raise ValueError("El primer responsable no tiene correo")
    permission_id_primero = asegurar_permiso_rol(
        drive_service=drive_service,
        file_id=copia["id"],
        email=email_primero,
        role="writer",
    )

    cerrar_gestion_revision_externa_anterior_por_observacion(
        cadena_anterior=cadena_anterior,
        comentario=comentario,
        fecha=fecha,
    )

    cadena_nueva_existente = buscar_cadena_documento_version(
        id_documento=id_documento,
        numero_version=numero_version_nueva,
    )
    if cadena_nueva_existente:
        orden_primero = entero(primer_anterior.get("ORDEN"), "ORDEN")
        primeros = [
            fila
            for fila in cadena_nueva_existente
            if entero(fila.get("ORDEN"), "ORDEN") == orden_primero
        ]
        if len(primeros) != 1:
            raise RuntimeError(
                "La cadena nueva existente no tiene un único primer responsable"
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
            indice_destino=cadena_anterior.index(primer_anterior),
            id_version_nueva=id_version_nueva,
            permission_id_destino=permission_id_primero,
            fecha=fecha,
        )
        appsheet_action(TABLA_APROBADORES_ACTUAL, "Add", filas_nuevas)

    if not version_existente:
        crear_registro_version_por_aprobacion(
            id_version=id_version_nueva,
            id_documento=id_documento,
            id_version_origen=id_version_observada,
            numero_version=numero_version_nueva,
            numero_revision=0,
            etapa="Borrador",
            nombre_archivo=copia["name"],
            google_doc_id=copia["id"],
            google_doc_url=copia["url"],
            id_aprobacion_responsable=primer_nuevo["ID_APROBACION_ACTUAL"],
            orden_responsable=entero(primer_nuevo.get("ORDEN"), "ORDEN"),
            motivo_creacion="Reinicio por observaciones de revisión externa",
            comentario=comentario,
            creado_por=usuario,
            fecha_creacion=fecha,
        )

    actualizar_estado_version(
        id_version=id_version_observada,
        estado_version="Observada",
        fecha_cierre=fecha,
    )

    # Archivamos la respuesta individual antes de limpiar el File temporal de
    # Documentos. Si una escritura posterior falla, el reintento puede recuperar
    # comentario y respaldo desde el detalle de esta misma ronda.
    actualizar_detalle_revision_externa(
        detalle=detalle,
        resultado="Observado",
        comentario=comentario,
        usuario=usuario,
        fecha=fecha,
        archivo_respaldo=archivo_efectivo,
    )

    actualizar_documento_reinicio_revision_externa(
        id_documento=id_documento,
        id_documento_raiz=id_documento_raiz,
        numero_version=numero_version_nueva,
        id_version_nueva=id_version_nueva,
        id_version_observada=id_version_observada,
        copia=copia,
        primer_encargado=primer_nuevo,
        usuario=usuario,
        fecha=fecha,
    )

    # El paquete queda con observaciones desde la primera observación, aunque
    # todavía existan otros documentos Pendientes en la misma ronda.
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento_raiz,
                "ESTADO_PAQUETE_NOTARIAL": "Con observaciones",
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
            }
        ],
    )

    advertencias: list[str] = []
    evento_reinicio: dict[str, Any] | None = None
    try:
        evento_reinicio = crear_eventos_reinicio_revision_externa(
            id_documento=id_documento,
            id_version_observada=id_version_observada,
            id_version_nueva=id_version_nueva,
            id_aprobacion_nueva=primer_nuevo["ID_APROBACION_ACTUAL"],
            usuario=usuario,
            fecha=fecha,
            comentario=comentario,
            numero_version_nueva=numero_version_nueva,
            nombre_archivo=copia["name"],
        )
    except Exception as exc:
        traceback.print_exc()
        advertencias.append(
            "El documento se reinició, pero falló la creación de eventos: "
            f"{exc}"
        )

    notificaciones: list[dict[str, Any]] = []
    if evento_reinicio is not None:
        try:
            documento_actualizado = buscar_documento(id_documento)
            cadena_nueva = buscar_cadena_documento_version(
                id_documento=id_documento,
                numero_version=numero_version_nueva,
            )
            responsable_nuevo = None
            for fila in cadena_nueva:
                if es_responsable_firmas(fila):
                    responsable_nuevo = fila
                    break
            especificaciones = construir_especificaciones_reinicio_revision_externa(
                documento=documento_actualizado,
                evento=evento_reinicio,
                cadena=cadena_nueva,
                aprobador_destino=primer_nuevo,
                responsable_firmas=responsable_nuevo or responsable_firmas_anterior,
            )
            notificaciones = notificar_destinatarios_internos(
                documento=documento_actualizado,
                evento=evento_reinicio,
                destinatarios=especificaciones,
            )
            fallidas = [r for r in notificaciones if not r.get("ok")]
            if fallidas:
                advertencias.append(
                    f"{len(fallidas)} notificación(es) quedaron omitidas o con error. "
                    "Revisa Documento_Notificaciones."
                )
        except Exception as exc:
            traceback.print_exc()
            advertencias.append(
                "El documento se reinició correctamente, pero fallaron las "
                f"notificaciones internas: {exc}"
            )

    return {
        "ya_reiniciado": False,
        "id_version_nueva": id_version_nueva,
        "numero_version_nueva": numero_version_nueva,
        "google_doc_id": copia["id"],
        "google_doc_url": copia["url"],
        "orden_actual": primer_nuevo["ORDEN"],
        "id_aprobacion_actual": primer_nuevo["ID_APROBACION_ACTUAL"],
        "encargado_actual": primer_nuevo.get("NOMBRE", ""),
        "encargado_email": email_primero,
        "notificaciones": notificaciones,
        "advertencias": advertencias,
    }


@app.route("/registrar-respuesta-revision-externa", methods=["POST"])
def registrar_respuesta_revision_externa():
    """Registra Aprobado u Observado para un documento de una ronda externa."""
    id_documento = ""
    try:
        validar_configuracion()
        validar_token()
        data = request.get_json(silent=True) or {}
        id_documento = texto(data.get("id_documento"))
        id_documento_raiz = texto(data.get("id_documento_raiz"))
        id_revision_externa = texto(data.get("id_revision_externa"))
        id_version = texto(data.get("id_version"))
        usuario = texto(data.get("usuario"))
        accion = texto(data.get("accion"))
        comentario = texto(data.get("comentario"))
        archivo_respaldo = texto(data.get("archivo_respaldo"))

        if not id_documento:
            return {"error": "Falta id_documento"}, 400
        if not id_documento_raiz:
            return {"error": "Falta id_documento_raiz"}, 400
        if not id_revision_externa:
            return {"error": "Falta id_revision_externa"}, 400
        if accion not in {
            "Aprobar revisión externa",
            "Observar revisión externa",
        }:
            raise ValueError(
                f"Acción de revisión externa no reconocida: {accion!r}"
            )
        if accion == "Observar revisión externa" and not comentario:
            raise ValueError(
                "El comentario es obligatorio para observar una revisión externa"
            )
        if not usuario or not _EMAIL_RE.fullmatch(usuario.lower()):
            raise ValueError("El usuario que registra la respuesta no es válido")

        documento, raiz, revision, detalle = validar_contexto_respuesta_revision_externa(
            id_documento=id_documento,
            id_documento_raiz=id_documento_raiz,
            id_revision_externa=id_revision_externa,
            id_version=id_version,
        )
        id_raiz = texto(raiz.get("ID_DOCUMENTO"))
        id_version_enviada = texto(detalle.get("ID_VERSION_ENVIADA"))
        resultado_detalle = normalizar_resultado_revision_externa(
            detalle.get("RESULTADO")
        )
        resultado_esperado = (
            "Aprobado"
            if accion == "Aprobar revisión externa"
            else "Observado"
        )
        archivo_efectivo = archivo_respaldo or texto(
            detalle.get("ARCHIVO_RESPALDO_RESPUESTA")
        )
        if resultado_esperado == "Observado" and not normalizar_archivo_observacion_appsheet(
            archivo_efectivo
        ):
            raise ValueError(
                "El archivo de respaldo es obligatorio al observar una revisión externa"
            )
        fecha = ahora_iso()

        advertencias: list[str] = []
        datos_reinicio: dict[str, Any] = {}

        if resultado_detalle == resultado_esperado:
            # Reintento idempotente. Reconciliamos Documentos antes de responder.
            if resultado_esperado == "Aprobado":
                actualizar_documento_aprobado_revision_externa(
                    id_documento=id_documento,
                    id_version=id_version_enviada,
                    usuario=usuario,
                    fecha=fecha,
                )
            else:
                datos_reinicio = procesar_observacion_revision_externa(
                    documento=documento,
                    raiz=raiz,
                    revision=revision,
                    detalle=detalle,
                    id_version_observada=id_version_enviada,
                    usuario=usuario,
                    comentario=comentario or texto(detalle.get("COMENTARIO")),
                    fecha=fecha,
                    archivo_respaldo=archivo_efectivo,
                )
                advertencias.extend(datos_reinicio.get("advertencias") or [])

            estados = recalcular_revision_y_paquete_notarial(
                id_revision_externa=id_revision_externa,
                id_documento_raiz=id_raiz,
                usuario=usuario,
                fecha=fecha,
            )
            notificaciones_post, advertencias_post = (
                ejecutar_notificaciones_post_respuesta_revision_externa(
                    documento=documento,
                    raiz_antes=raiz,
                    id_revision_externa=id_revision_externa,
                    id_version_enviada=id_version_enviada,
                    resultado=resultado_esperado,
                    comentario=comentario or texto(detalle.get("COMENTARIO")),
                    usuario=usuario,
                    fecha=fecha,
                    estado_paquete_nuevo=estados["estado_paquete_notarial"],
                )
            )
            advertencias.extend(advertencias_post)
            documento_final = buscar_documento(id_documento)
            return jsonify(
                {
                    "ok": True,
                    "ya_procesado": True,
                    "id_documento": id_documento,
                    "id_documento_raiz": id_raiz,
                    "id_revision_externa": id_revision_externa,
                    "resultado_revision_externa": resultado_esperado,
                    "estado": texto(documento_final.get("ESTADO")),
                    **estados,
                    "reinicio": datos_reinicio,
                    "notificaciones": notificaciones_post,
                    "advertencias": advertencias,
                }
            )

        if resultado_detalle not in {"", "Pendiente"}:
            raise ValueError(
                "El documento ya tiene una respuesta distinta en esta ronda externa: "
                f"{resultado_detalle!r}"
            )

        resultado_documento = normalizar_resultado_revision_externa(
            documento.get("RESULTADO_REVISION_EXTERNA")
        )
        if resultado_documento != "Pendiente":
            # Excepción de recuperación: el reinicio pudo completar Documentos
            # antes de alcanzar Documento_Revision_Externa_Detalle.
            if not (
                resultado_esperado == "Observado"
                and resultado_documento == "Observado"
                and texto(documento.get("ID_VERSION_REVISION_EXTERNA"))
                == id_version_enviada
            ):
                raise ValueError(
                    "Solo un documento Pendiente puede registrar respuesta externa. "
                    f"Resultado actual: {resultado_documento!r}"
                )

        if resultado_esperado == "Aprobado":
            if texto(documento.get("ID_VERSION_ACTUAL")) != id_version_enviada:
                raise ValueError(
                    "La versión actual ya no coincide con la versión enviada a revisión externa"
                )
            actualizar_detalle_revision_externa(
                detalle=detalle,
                resultado="Aprobado",
                comentario=comentario,
                usuario=usuario,
                fecha=fecha,
                archivo_respaldo=archivo_efectivo,
            )
            actualizar_documento_aprobado_revision_externa(
                id_documento=id_documento,
                id_version=id_version_enviada,
                usuario=usuario,
                fecha=fecha,
            )
        else:
            datos_reinicio = procesar_observacion_revision_externa(
                documento=documento,
                raiz=raiz,
                revision=revision,
                detalle=detalle,
                id_version_observada=id_version_enviada,
                usuario=usuario,
                comentario=comentario,
                fecha=fecha,
                archivo_respaldo=archivo_efectivo,
            )
            advertencias.extend(datos_reinicio.get("advertencias") or [])

        estados = recalcular_revision_y_paquete_notarial(
            id_revision_externa=id_revision_externa,
            id_documento_raiz=id_raiz,
            usuario=usuario,
            fecha=fecha,
        )
        notificaciones_post, advertencias_post = (
            ejecutar_notificaciones_post_respuesta_revision_externa(
                documento=documento,
                raiz_antes=raiz,
                id_revision_externa=id_revision_externa,
                id_version_enviada=id_version_enviada,
                resultado=resultado_esperado,
                comentario=comentario,
                usuario=usuario,
                fecha=fecha,
                estado_paquete_nuevo=estados["estado_paquete_notarial"],
            )
        )
        advertencias.extend(advertencias_post)
        documento_final = buscar_documento(id_documento)

        return jsonify(
            {
                "ok": True,
                "ya_procesado": False,
                "id_documento": id_documento,
                "id_documento_raiz": id_raiz,
                "id_revision_externa": id_revision_externa,
                "id_version_respuesta_externa": id_version_enviada,
                "resultado_revision_externa": resultado_esperado,
                "estado": texto(documento_final.get("ESTADO")),
                **estados,
                "reinicio": datos_reinicio,
                "notificaciones": notificaciones_post,
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


def normalizar_archivo_observacion_appsheet(valor: str) -> str:
    """
    Devuelve la ruta relativa real de un File de AppSheet.

    En automatizaciones AppSheet puede expandir una columna File como una URL
    gettablefileurl?...&fileName=Documentos_Files_/archivo.docx. En ese caso
    conservamos únicamente el parámetro fileName, porque Documento_Versiones
    también es una columna File y debe almacenar la ruta relativa del archivo.
    """
    valor_limpio = texto(valor).strip()
    if not valor_limpio:
        return ""

    valor_normalizado = valor_limpio.replace("\\", "/")

    if "://" in valor_normalizado:
        parsed = urlparse(valor_normalizado)
        consulta = parse_qs(parsed.query)
        archivo = texto((consulta.get("fileName") or [""])[0])
        if archivo:
            return unquote(archivo).replace("\\", "/")

        # Fallback para una URL directa que sí termina físicamente en archivo.
        return unquote(parsed.path).replace("\\", "/")

    return unquote(
        valor_normalizado.split("?", 1)[0].split("#", 1)[0]
    )


def nombre_archivo_desde_valor_appsheet(valor: str) -> str:
    """Obtiene el nombre final desde un File relativo o una URL AppSheet."""
    ruta = normalizar_archivo_observacion_appsheet(valor)
    if not ruta:
        return ""
    return PurePosixPath(ruta).name.strip()


EXTENSIONES_ARCHIVO_OBSERVACION_FIRMA = {".pdf", ".doc", ".docx"}


def validar_archivo_observacion_firma(valor: str) -> tuple[str, str]:
    """
    Valida el archivo opcional recibido como respaldo de observaciones de firma.

    El valor corresponde al File/path que AppSheet ya guardó. Cloud Run no
    descarga ni transforma este archivo: solo conserva su referencia histórica
    en Documento_Versiones.
    """
    ruta = normalizar_archivo_observacion_appsheet(valor)
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
    estado_anterior: str = "En firma",
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
        "ESTADO_ANTERIOR": estado_anterior,
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
        "ESTADO_ANTERIOR": estado_anterior,
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


def validar_documento_en_paquete_activo(
    *,
    documento: dict[str, Any],
    id_documento_raiz: str,
) -> dict[str, Any]:
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    raiz_asignada = texto(documento.get("ID_DOCUMENTO_RAIZ_FIRMA"))
    if not raiz_asignada:
        raise ValueError(f"{id_documento}: no pertenece a un paquete de firma activo")
    if raiz_asignada != id_documento_raiz:
        raise ValueError(
            f"{id_documento}: pertenece al paquete {raiz_asignada}, no a {id_documento_raiz}"
        )
    raiz = buscar_documento(id_documento_raiz)
    estado_paquete = texto(raiz.get("ESTADO_PAQUETE_FIRMA"))
    if estado_paquete not in ESTADOS_PAQUETE_FIRMA_ACTIVOS:
        raise ValueError(
            f"El paquete {id_documento_raiz} no está activo. Estado: {estado_paquete!r}"
        )
    return raiz


@app.route("/registrar-respuesta-firma", methods=["POST"])
def registrar_respuesta_firma():
    """Registra Conforme o PDF Firmado sin cerrar todavía el paquete."""
    id_documento = ""
    try:
        validar_configuracion()
        validar_token()
        data = request.get_json(silent=True) or {}
        id_documento = texto(data.get("id_documento"))
        id_documento_raiz = texto(data.get("id_documento_raiz"))
        id_version = texto(data.get("id_version"))
        usuario = texto(data.get("usuario"))
        accion = texto(data.get("accion"))
        pdf_firmado = texto(data.get("pdf_firmado"))

        if not id_documento:
            return {"error": "Falta id_documento"}, 400
        if not id_documento_raiz:
            return {"error": "Falta id_documento_raiz"}, 400
        if accion not in {"Conforme firma", "Registrar PDF firmado"}:
            raise ValueError(f"Acción de respuesta externa no reconocida: {accion!r}")
        if not usuario or not _EMAIL_RE.fullmatch(usuario.lower()):
            raise ValueError("El usuario que registra la respuesta no es válido")

        documento = buscar_documento(id_documento)
        validar_documento_en_paquete_activo(
            documento=documento,
            id_documento_raiz=id_documento_raiz,
        )
        id_version_actual = texto(documento.get("ID_VERSION_ACTUAL"))
        id_version_esperada = texto(documento.get("ID_VERSION_RESPUESTA_EXTERNA"))
        id_version = id_version or id_version_actual
        if not id_version or id_version != id_version_actual:
            raise ValueError("La versión indicada ya no coincide con la versión actual")
        if id_version_esperada and id_version_esperada != id_version:
            raise ValueError(
                "La respuesta corresponde a una versión distinta de la enviada externamente"
            )

        resultado_actual = normalizar_resultado_firma_externa(
            documento.get("RESULTADO_FIRMA_EXTERNA")
        )

        if accion == "Conforme firma":
            if resultado_actual == "Conforme":
                return jsonify(
                    {
                        "ok": True,
                        "ya_procesado": True,
                        "id_documento": id_documento,
                        "resultado_firma_externa": "Conforme",
                    }
                )
            if resultado_actual != "Pendiente":
                raise ValueError(
                    "Solo un documento Pendiente puede registrarse como Conforme. "
                    f"Resultado actual: {resultado_actual!r}"
                )
            nuevo_resultado = "Conforme"
            cambios_extra: dict[str, Any] = {}
        else:
            if resultado_actual == "Firmado" and texto(documento.get("PDF_FIRMADO")):
                return jsonify(
                    {
                        "ok": True,
                        "ya_procesado": True,
                        "id_documento": id_documento,
                        "resultado_firma_externa": "Firmado",
                    }
                )
            if resultado_actual not in {"Pendiente", "Conforme"}:
                raise ValueError(
                    "Solo un documento Pendiente o Conforme puede registrarse como Firmado. "
                    f"Resultado actual: {resultado_actual!r}"
                )
            pdf_firmado = pdf_firmado or texto(documento.get("PDF_FIRMADO"))
            if not pdf_firmado:
                raise ValueError("Debe cargar PDF_FIRMADO antes de registrar la firma")
            if not nombre_archivo_desde_valor_appsheet(pdf_firmado).lower().endswith(".pdf"):
                raise ValueError("PDF_FIRMADO debe corresponder a un archivo .pdf")
            nuevo_resultado = "Firmado"
            cambios_extra = {"PDF_FIRMADO": pdf_firmado, "CARGADO_POR": usuario}

        fecha = ahora_iso()
        cambios = {
            "ID_DOCUMENTO": id_documento,
            "RESULTADO_FIRMA_EXTERNA": nuevo_resultado,
            "ID_VERSION_RESPUESTA_EXTERNA": id_version,
            "RESPUESTA_FIRMA_EXTERNA_POR": usuario,
            "FECHA_RESPUESTA_FIRMA_EXTERNA": fecha,
            "ULTIMO_ENVIADO_POR": usuario,
            "FECHA_ULTIMO_ENVIO": fecha,
            "FECHA_ULTIMA_ACTUALIZACION": fecha,
            "ACCION_SOLICITADA": "",
            **cambios_extra,
        }
        appsheet_action(TABLA_DOCUMENTOS, "Edit", [cambios])

        try:
            appsheet_action(
                TABLA_EVENTOS,
                "Add",
                [
                    {
                        "ID_EVENTO": nuevo_id(),
                        "ID_DOCUMENTO": id_documento,
                        "ID_VERSION": id_version,
                        "ID_APROBACION_ACTUAL": "",
                        "TIPO_EVENTO": "Respuesta de firma externa",
                        "ESTADO_ANTERIOR": resultado_actual,
                        "ESTADO_NUEVO": nuevo_resultado,
                        "USUARIO": usuario,
                        "FECHA_EVENTO": fecha,
                        "COMENTARIO": (
                            "El usuario registró la respuesta externa como "
                            f"{nuevo_resultado}."
                        ),
                    }
                ],
            )
        except Exception:
            traceback.print_exc()

        estado_paquete = actualizar_estado_paquete_desde_resultados(
            id_documento_raiz,
            fecha=fecha,
        )
        return jsonify(
            {
                "ok": True,
                "ya_procesado": False,
                "id_documento": id_documento,
                "id_documento_raiz": id_documento_raiz,
                "resultado_firma_externa": nuevo_resultado,
                "id_version_respuesta_externa": id_version,
                "estado_paquete_firma": estado_paquete,
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


def cerrar_documento_firmado_del_paquete(
    *,
    documento: dict[str, Any],
    usuario: str,
    drive_service: Any,
) -> dict[str, Any]:
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    titulo = texto(documento.get("TITULO")) or id_documento
    resultado = normalizar_resultado_firma_externa(
        documento.get("RESULTADO_FIRMA_EXTERNA")
    )
    if resultado != "Firmado":
        raise ValueError(f"{titulo}: el resultado externo todavía no es Firmado")

    if texto(documento.get("ESTADO")) == "Proceso terminado" and texto(
        documento.get("PDF_FIRMADO_ID")
    ):
        return {
            "id_documento": id_documento,
            "titulo": titulo,
            "ya_procesado": True,
            "id_version_final": texto(documento.get("ID_VERSION_ACTUAL")),
            "pdf_firmado_id": texto(documento.get("PDF_FIRMADO_ID")),
            "pdf_firmado_url": texto(documento.get("PDF_FIRMADO_URL")),
            "advertencias": [],
        }

    id_version_origen = texto(documento.get("ID_VERSION_RESPUESTA_EXTERNA")) or texto(
        documento.get("ID_VERSION_ACTUAL")
    )
    if not id_version_origen:
        raise ValueError(f"{titulo}: falta la versión firmada externamente")
    if texto(documento.get("ID_VERSION_ACTUAL")) != id_version_origen:
        raise ValueError(
            f"{titulo}: la versión actual cambió después de recibir la firma"
        )
    valor_pdf_firmado = texto(documento.get("PDF_FIRMADO"))
    if not valor_pdf_firmado:
        raise ValueError(f"{titulo}: falta PDF_FIRMADO")

    version_origen = buscar_version_por_id(id_version_origen)
    numero_version = entero(documento.get("VERSION_ACTUAL"), "VERSION_ACTUAL")
    numero_revision = entero(documento.get("REVISION_ACTUAL") or 0, "REVISION_ACTUAL")
    plantilla = buscar_plantilla(texto(documento.get("ID_PLANTILLA")))
    folder_id = texto(plantilla.get("CARPETA_DESTINO_ID"))
    if not folder_id:
        raise ValueError(f"{titulo}: la plantilla no tiene CARPETA_DESTINO_ID")

    fecha = ahora_iso()
    pdf_subido = buscar_pdf_cargado_appsheet(drive_service, valor_pdf_firmado)
    titulo_archivo = limpiar_nombre_archivo(titulo)
    nombre_final = f"{titulo_archivo}_V{numero_version:02d}_FIRMADO.pdf"
    pdf_final = copiar_pdf_firmado_o_reutilizar(
        drive_service=drive_service,
        source_file_id=texto(pdf_subido.get("id")),
        folder_id=folder_id,
        nombre_final=nombre_final,
    )

    version_firmada_existente = buscar_version_firmada(id_documento, numero_version)
    if version_firmada_existente:
        id_version_final = texto(version_firmada_existente.get("ID_VERSION"))
    else:
        id_version_final = nuevo_id()
        actualizar_estado_version(
            id_version=id_version_origen,
            estado_version="Cerrada",
            fecha_cierre=fecha,
        )
        crear_version_firmada(
            id_version=id_version_final,
            id_documento=id_documento,
            id_version_origen=id_version_origen,
            numero_version=numero_version,
            numero_revision=numero_revision,
            nombre_archivo=nombre_final,
            google_doc_id=texto(version_origen.get("GOOGLE_DOC_ID")),
            google_doc_url=normalizar_url_appsheet(version_origen.get("GOOGLE_DOC_URL")),
            pdf_id=pdf_final["id"],
            pdf_url=pdf_final["url"],
            comentario="Cierre de paquete de firma simple",
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
            comentario="Cierre de paquete de firma simple",
            estado_anterior=texto(documento.get("ESTADO")) or "En firma externa",
        )
    except Exception as exc_evento:
        traceback.print_exc()
        advertencias.append(f"No se pudieron crear todos los eventos: {exc_evento}")

    if evento_cierre is not None:
        try:
            documento_actualizado = buscar_documento(id_documento)
            cadena = buscar_cadena_documento_version(
                id_documento=id_documento,
                numero_version=numero_version,
            )
            resultados_notificacion = ejecutar_notificaciones_cierre_proceso(
                documento=documento_actualizado,
                evento=evento_cierre,
                cadena=cadena,
            )
            fallidas = [fila for fila in resultados_notificacion if not fila.get("ok")]
            if fallidas:
                advertencias.append(
                    f"{len(fallidas)} notificación(es) internas quedaron con error"
                )
        except Exception as exc_notificacion:
            traceback.print_exc()
            advertencias.append(
                f"Falló el proceso de notificaciones internas: {exc_notificacion}"
            )

    return {
        "id_documento": id_documento,
        "titulo": titulo,
        "ya_procesado": False,
        "id_version_final": id_version_final,
        "pdf_firmado_id": pdf_final["id"],
        "pdf_firmado_url": pdf_final["url"],
        "pdf_firmado_nombre": pdf_final["name"],
        "advertencias": advertencias,
    }


@app.route("/registrar-firma", methods=["POST"])
def registrar_firma():
    """Cierra definitivamente un paquete cuando todos sus documentos están Firmados."""
    id_documento = ""
    try:
        validar_configuracion()
        validar_token()
        data = request.get_json(silent=True) or {}
        id_documento = texto(data.get("id_documento"))
        usuario = texto(data.get("usuario"))
        if not id_documento:
            return {"error": "Falta id_documento"}, 400

        contexto = obtener_contexto_jerarquia_documental(id_documento)
        if not contexto.get("solicitado_es_raiz"):
            raise ValueError("Solo la raíz puede cerrar el paquete de firma")
        id_raiz = contexto["id_documento_raiz"]
        raiz = contexto["raiz"]
        if texto(raiz.get("ESTADO_PAQUETE_FIRMA")) == "Cerrado":
            return jsonify(
                {
                    "ok": True,
                    "ya_procesado": True,
                    "id_documento_raiz": id_raiz,
                    "estado_paquete_firma": "Cerrado",
                }
            )

        validar_membresia_paquete_activo(
            id_documento_raiz=id_raiz,
            integrantes_jerarquia=contexto["integrantes"],
        )
        activos = buscar_integrantes_paquete_activo(id_raiz)
        faltantes: list[str] = []
        for fila in activos:
            titulo = texto(fila.get("TITULO")) or texto(fila.get("ID_DOCUMENTO"))
            resultado = normalizar_resultado_firma_externa(
                fila.get("RESULTADO_FIRMA_EXTERNA")
            )
            if resultado != "Firmado":
                faltantes.append(f"{titulo}: resultado {resultado or '<vacío>'}")
            elif not texto(fila.get("PDF_FIRMADO")) and not texto(fila.get("PDF_FIRMADO_ID")):
                faltantes.append(f"{titulo}: falta PDF_FIRMADO")
        if faltantes:
            raise ValueError(
                "No se puede cerrar el paquete. " + " | ".join(faltantes)
            )

        usuario = usuario or texto(raiz.get("ULTIMO_ENVIADO_POR")) or texto(
            raiz.get("ENVIADO_FIRMA_POR")
        )
        if not usuario or not _EMAIL_RE.fullmatch(usuario.lower()):
            raise ValueError("El usuario que cierra la firma no es válido")

        # Procesamos descendientes primero y la raíz al final. En caso de
        # error parcial, un reintento omite los documentos ya cerrados.
        orden = sorted(
            contexto["integrantes"],
            key=lambda fila: (fila.get("nivel", 0), fila.get("titulo", "")),
            reverse=True,
        )
        drive_service = obtener_drive_service()
        resultados: list[dict[str, Any]] = []
        for item in orden:
            documento_actual = buscar_documento(item["id_documento"])
            resultados.append(
                cerrar_documento_firmado_del_paquete(
                    documento=documento_actual,
                    usuario=usuario,
                    drive_service=drive_service,
                )
            )

        fecha = ahora_iso()
        appsheet_action(
            TABLA_DOCUMENTOS,
            "Edit",
            [
                {
                    "ID_DOCUMENTO": id_raiz,
                    "ESTADO_PAQUETE_FIRMA": "Cerrado",
                    "ESTADO_FIRMA": "Firmado",
                    "FECHA_ULTIMA_ACTUALIZACION": fecha,
                    "ACCION_SOLICITADA": "",
                    "OBSERVACION_ACTUAL": "",
                }
            ],
        )
        try:
            appsheet_action(
                TABLA_EVENTOS,
                "Add",
                [
                    {
                        "ID_EVENTO": nuevo_id(),
                        "ID_DOCUMENTO": id_raiz,
                        "ID_VERSION": texto(buscar_documento(id_raiz).get("ID_VERSION_ACTUAL")),
                        "ID_APROBACION_ACTUAL": "",
                        "TIPO_EVENTO": "Paquete de firma cerrado",
                        "ESTADO_ANTERIOR": "Listo para cierre",
                        "ESTADO_NUEVO": "Cerrado",
                        "USUARIO": usuario,
                        "FECHA_EVENTO": fecha,
                        "COMENTARIO": (
                            f"Se cerró el paquete con {len(resultados)} documento(s) firmados."
                        ),
                    }
                ],
            )
        except Exception:
            traceback.print_exc()

        advertencias = [
            adv
            for resultado in resultados
            for adv in (resultado.get("advertencias") or [])
        ]
        return jsonify(
            {
                "ok": True,
                "ya_procesado": False,
                "id_documento_raiz": id_raiz,
                "estado_paquete_firma": "Cerrado",
                "cantidad_documentos": len(resultados),
                "documentos": resultados,
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
# Flujo: observaciones/rechazo después del envío a firma
# -----------------------------------------------------------------------------


def actualizar_documento_reinicio_firma(
    *,
    id_documento: str,
    numero_version: int,
    id_version: str,
    id_version_observada: str,
    copia: dict[str, str],
    primer_encargado: dict[str, Any],
    usuario: str,
    fecha: str,
) -> None:
    """Reinicia solo el documento observado y conserva su paquete externo activo."""
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
                "ID_APROBACION_ACTUAL": primer_encargado["ID_APROBACION_ACTUAL"],
                "ENCARGADO_ACTUAL_NOMBRE": primer_encargado.get("NOMBRE", ""),
                "ENCARGADO_ACTUAL_EMAIL": primer_encargado.get("APROBADOR", ""),
                "PDF_PARA_FIRMA_ID": "",
                "PDF_PARA_FIRMA_URL": "",
                "ESTADO_FIRMA": "Observado",
                "RESULTADO_FIRMA_EXTERNA": "Observado",
                "ID_VERSION_RESPUESTA_EXTERNA": id_version_observada,
                "RESPUESTA_FIRMA_EXTERNA_POR": usuario,
                "FECHA_RESPUESTA_FIRMA_EXTERNA": fecha,
                "PDF_FIRMADO": "",
                "PDF_FIRMADO_ID": "",
                "PDF_FIRMADO_URL": "",
                "FECHA_FIRMA_COMPLETA": "",
                "CARGADO_POR": "",
                "ARCHIVO_OBSERVACION_FIRMA_TEMP": "",
                "FECHA_CIERRE": "",
                "ULTIMO_ENVIADO_POR": usuario,
                # FECHA_ULTIMO_ENVIO se reserva para acciones iniciadas desde
                # AppSheet. No debe cambiar durante este reinicio interno, porque
                # podría disparar Bots de envío a revisión inmediatamente.
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "OBSERVACION_ACTUAL": "",
                "ACCION_SOLICITADA": "",
            }
        ],
    )


def marcar_paquete_con_observaciones(
    id_documento_raiz: str,
    *,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento_raiz,
                "ESTADO_PAQUETE_FIRMA": "Con observaciones",
                "ESTADO_FIRMA": "Observado",
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
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
    estado_anterior: str = "En firma",
) -> tuple[list[str], dict[str, Any] | None]:
    """Crea la bitácora del reinicio y devuelve el evento principal."""
    advertencias: list[str] = []
    evento_reinicio: dict[str, Any] = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION": id_version_observada,
        "ID_APROBACION_ACTUAL": "",
        "TIPO_EVENTO": "Proceso reiniciado",
        "ESTADO_ANTERIOR": estado_anterior,
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
        id_documento_raiz_solicitud = texto(data.get("id_documento_raiz"))
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
        id_documento_raiz = texto(documento.get("ID_DOCUMENTO_RAIZ_FIRMA"))
        if not id_documento_raiz:
            raise ValueError("El documento no pertenece a un paquete de firma activo")
        if id_documento_raiz_solicitud and id_documento_raiz_solicitud != id_documento_raiz:
            raise ValueError("id_documento_raiz no coincide con el paquete activo del documento")
        raiz_paquete = buscar_documento(id_documento_raiz)
        if texto(raiz_paquete.get("ESTADO_PAQUETE_FIRMA")) not in ESTADOS_PAQUETE_FIRMA_ACTIVOS:
            raise ValueError("El paquete de firma ya no está activo")
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
                        usuario=usuario or texto(raiz_paquete.get("ENVIADO_FIRMA_POR")),
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

            try:
                marcar_paquete_con_observaciones(
                    id_documento_raiz,
                    fecha=ahora_iso(),
                )
            except Exception:
                traceback.print_exc()

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

        resultado_firma_externa = normalizar_resultado_firma_externa(
            documento.get("RESULTADO_FIRMA_EXTERNA")
        )
        if resultado_firma_externa != "Pendiente":
            raise ValueError(
                "Solo puede rechazarse un documento con respuesta externa Pendiente. "
                f"Resultado actual: {resultado_firma_externa!r}"
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

        usuario_envio = texto(raiz_paquete.get("ENVIADO_FIRMA_POR"))
        usuario = usuario or usuario_envio
        if not usuario or not _EMAIL_RE.fullmatch(usuario.lower()):
            raise ValueError("El usuario que reinicia la aprobación no es válido")
        if usuario_envio and usuario_envio.lower() != usuario.lower():
            raise PermissionError(
                "Solo el usuario que administra el paquete de firma puede "
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

        primer_anterior = obtener_primer_responsable_aprobacion(
            cadena_anterior,
            contexto=f"Documento {id_documento} versión {numero_version_anterior}",
        )
        email_primero = texto(primer_anterior.get("APROBADOR"))
        if not email_primero:
            raise ValueError("El primer responsable no tiene correo")

        permission_id_primero = asegurar_permiso_rol(
            drive_service=drive_service,
            file_id=copia["id"],
            email=email_primero,
            role="writer",
        )

        cerrar_gestion_firma_anterior_por_observacion(
            cadena_anterior=cadena_anterior,
            comentario=comentario,
            fecha=fecha,
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
                indice_destino=cadena_anterior.index(primer_anterior),
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
            id_version_observada=id_version_origen,
            copia=copia,
            primer_encargado=primer_nuevo,
            usuario=usuario,
            fecha=fecha,
        )
        marcar_paquete_con_observaciones(
            id_documento_raiz,
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
            estado_anterior=estado_documento or "En firma externa",
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
                "id_documento_raiz": id_documento_raiz,
                "estado_paquete_firma": "Con observaciones",
                "resultado_firma_externa": "Observado",
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


# -----------------------------------------------------------------------------
# Flujo Notarial - envío del paquete aprobado a notaría
# -----------------------------------------------------------------------------

ESTADO_PAQUETE_NOTARIAL_EN_NOTARIA = "En notaría"
ESTADO_DOCUMENTO_EN_NOTARIA = "En notaría"
ESTADO_DOCUMENTO_PROCESO_TERMINADO = "Proceso terminado"
ESTADO_PAQUETE_NOTARIAL_CERRADO = "Cerrado"
EXTENSIONES_ANTECEDENTES_NOTARIA = {".pdf", ".png", ".jpg", ".jpeg", ".doc", ".docx"}

# Tamaño máximo de archivos reales por correo a notaría.
# Gmail codifica los adjuntos en base64, lo que aumenta el tamaño del mensaje
# aproximadamente un 33 %. 17 MiB deja margen bajo el límite de Gmail.
MAX_ADJUNTOS_NOTARIA_MB = float(os.environ.get("MAX_ADJUNTOS_NOTARIA_MB", "17"))
MAX_ADJUNTOS_NOTARIA_BYTES = int(MAX_ADJUNTOS_NOTARIA_MB * 1024 * 1024)


def primer_valor(fila: dict[str, Any], *columnas: str) -> str:
    """Devuelve el primer valor no vacío entre posibles nombres de columna."""
    for columna in columnas:
        valor = texto(fila.get(columna))
        if valor:
            return valor
    return ""


def buscar_notaria_por_id(id_notaria: str) -> dict[str, Any]:
    selector = (
        f"FILTER({TABLA_NOTARIAS}, "
        f"[ID_NOTARIA] = {literal_appsheet(id_notaria)})"
    )
    filas = appsheet_find(TABLA_NOTARIAS, selector)
    if not filas:
        raise LookupError(f"No se encontró ID_NOTARIA={id_notaria}")
    if len(filas) > 1:
        raise ValueError(f"ID_NOTARIA duplicado: {id_notaria}")
    return filas[0]


def normalizar_datos_notaria(notaria: dict[str, Any]) -> dict[str, str]:
    datos = {
        "id_notaria": primer_valor(notaria, "ID_NOTARIA"),
        "nombre": primer_valor(notaria, "NOMBRE_NOTARIA", "NOMBRE"),
        "direccion": primer_valor(notaria, "DIRECCION_NOTARIA", "DIRECCION"),
        "contacto": primer_valor(
            notaria,
            "PERSONA_CONTACTO_NOTARIA",
            "PERSONA_CONTACTO",
            "ENCARGADO_NOTARIA",
            "ENCARGADO",
        ),
        "email": primer_valor(notaria, "EMAIL_NOTARIA", "EMAIL").lower(),
        "horario": primer_valor(
            notaria,
            "HORARIO_ATENCION_NOTARIA",
            "HORARIO_ATENCION",
            "HORARIO",
        ),
    }
    if not datos["id_notaria"]:
        raise ValueError("La notaría no tiene ID_NOTARIA")
    if not datos["nombre"]:
        raise ValueError("La notaría no tiene NOMBRE_NOTARIA")
    if not datos["email"] or not _EMAIL_RE.fullmatch(datos["email"]):
        raise ValueError("La notaría no tiene un EMAIL_NOTARIA válido")
    if not datos["direccion"]:
        raise ValueError("La notaría no tiene DIRECCION_NOTARIA")
    if not datos["contacto"]:
        raise ValueError("La notaría no tiene PERSONA_CONTACTO_NOTARIA")
    if not datos["horario"]:
        raise ValueError("La notaría no tiene HORARIO_ATENCION_NOTARIA")
    return datos


def buscar_documentos_prime_propiedad(id_propiedad: str) -> list[dict[str, Any]]:
    selector = (
        f"FILTER({TABLA_DOCUMENTOS_PRIME}, "
        f"[id_Propiedades] = {literal_appsheet(id_propiedad)})"
    )
    return appsheet_find(TABLA_DOCUMENTOS_PRIME, selector)


def mime_adjunto_desde_nombre(nombre: str) -> tuple[str, str]:
    extension = PurePosixPath(nombre.lower()).suffix
    mapa = {
        ".pdf": ("application", "pdf"),
        ".png": ("image", "png"),
        ".jpg": ("image", "jpeg"),
        ".jpeg": ("image", "jpeg"),
        ".doc": ("application", "msword"),
        ".docx": (
            "application",
            "vnd.openxmlformats-officedocument.wordprocessingml.document",
        ),
    }
    if extension not in mapa:
        raise ValueError(
            f"Formato no permitido para antecedente notarial: {nombre!r}"
        )
    return mapa[extension]


def descargar_file_appsheet(
    *,
    table_name: str,
    valor_file: Any,
) -> tuple[bytes, str, str, str]:
    """Descarga una columna File de AppSheet y devuelve bytes/nombre/MIME."""
    valor_original = texto(valor_file).strip()
    ruta = normalizar_archivo_observacion_appsheet(valor_original)
    if not ruta:
        raise ValueError(f"{table_name}: se recibió un File vacío")

    nombre = nombre_archivo_desde_valor_appsheet(ruta)
    if not nombre:
        raise ValueError(f"{table_name}: no se pudo determinar el nombre del archivo")

    extension = PurePosixPath(nombre.lower()).suffix
    if extension not in EXTENSIONES_ANTECEDENTES_NOTARIA:
        raise ValueError(
            f"{table_name}: el archivo {nombre!r} no tiene un formato admitido"
        )

    if valor_original.lower().startswith(("http://", "https://")) and "gettablefileurl" not in valor_original:
        url = valor_original
    else:
        prepared = requests.Request(
            "GET",
            "https://www.appsheet.com/template/gettablefileurl",
            params={
                "appName": APPSHEET_APP_ID,
                "tableName": table_name,
                "fileName": ruta,
            },
        ).prepare()
        url = prepared.url

    headers = {"Accept": "*/*"}
    if APPSHEET_ACCESS_KEY:
        headers["ApplicationAccessKey"] = APPSHEET_ACCESS_KEY

    respuesta = requests.get(url, headers=headers, timeout=90, allow_redirects=True)
    if respuesta.status_code != 200:
        raise RuntimeError(
            f"No se pudo descargar {nombre} desde AppSheet: "
            f"HTTP {respuesta.status_code} - {respuesta.text[:300]}"
        )
    if not respuesta.content:
        raise RuntimeError(f"AppSheet devolvió vacío el archivo {nombre}")

    maintype, subtype = mime_adjunto_desde_nombre(nombre)
    return respuesta.content, nombre, maintype, subtype


def guardar_docx_enviado_notaria(
    *,
    drive_service: Any,
    google_doc_id: str,
    contenido: bytes,
    nombre_docx: str,
    id_envio_notaria: str,
) -> dict[str, str]:
    """Guarda en Drive exactamente los mismos bytes DOCX que se adjuntarán."""
    metadata_origen = (
        drive_service.files()
        .get(
            fileId=google_doc_id,
            fields="id,parents",
            supportsAllDrives=True,
        )
        .execute()
    )
    parents = metadata_origen.get("parents") or []
    if not parents:
        raise RuntimeError(
            f"El Google Doc {google_doc_id} no tiene una carpeta padre para guardar el respaldo"
        )
    folder_id = texto(parents[0])

    base, _ = os.path.splitext(nombre_docx)
    nombre_respaldo = limpiar_nombre_archivo(
        f"{base}_ENVIADO_NOTARIA_{id_envio_notaria[:8]}"
    ) + ".docx"

    existente = buscar_archivo_en_carpeta(
        drive_service=drive_service,
        folder_id=folder_id,
        nombre_archivo=nombre_respaldo,
    )
    if existente:
        file_id = texto(existente.get("id"))
        return {
            "id": file_id,
            "name": texto(existente.get("name")) or nombre_respaldo,
            "url": texto(existente.get("url")) or f"https://drive.google.com/file/d/{file_id}/view",
        }

    media = MediaInMemoryUpload(
        contenido,
        mimetype=DOCX_MIME_TYPE,
        resumable=False,
    )
    archivo = (
        drive_service.files()
        .create(
            body={
                "name": nombre_respaldo,
                "parents": [folder_id],
                "mimeType": DOCX_MIME_TYPE,
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
        "name": texto(archivo.get("name")) or nombre_respaldo,
        "url": (
            texto(archivo.get("webViewLink"))
            or texto(archivo.get("webContentLink"))
            or f"https://drive.google.com/file/d/{file_id}/view"
        ),
    }


def buscar_envios_notaria_raiz(id_documento_raiz: str) -> list[dict[str, Any]]:
    selector = (
        f"FILTER({TABLA_ENVIOS_NOTARIA}, "
        f"[ID_DOCUMENTO_RAIZ] = {literal_appsheet(id_documento_raiz)})"
    )
    filas = appsheet_find(TABLA_ENVIOS_NOTARIA, selector)
    return sorted(
        filas,
        key=lambda fila: (
            parsear_fecha_appsheet(fila.get("FECHA_ENVIO")),
            texto(fila.get("ID_ENVIO_NOTARIA")),
        ),
    )


def buscar_envio_notaria_preparando(id_documento_raiz: str) -> dict[str, Any] | None:
    filas = buscar_envios_notaria_raiz(id_documento_raiz)
    candidatas = [
        fila for fila in filas if texto(fila.get("ESTADO_ENVIO")) == "Preparando"
    ]
    return candidatas[-1] if candidatas else None


def buscar_ultimo_envio_notaria(id_documento_raiz: str) -> dict[str, Any] | None:
    filas = buscar_envios_notaria_raiz(id_documento_raiz)
    return filas[-1] if filas else None


def obtener_o_crear_envio_notaria(
    *,
    raiz: dict[str, Any],
    id_notaria: str,
    id_propiedad_prime: str,
    firmantes: list[str],
    usuario: str,
    mensaje_adicional: str,
    datos_notaria: dict[str, str],
    fecha: str,
    reutilizar_ultimo: bool = False,
) -> tuple[dict[str, Any], bool]:
    id_raiz = texto(raiz.get("ID_DOCUMENTO"))

    if reutilizar_ultimo:
        existente = buscar_ultimo_envio_notaria(id_raiz)
        if existente:
            # En un reenvío controlado reutilizamos la misma cabecera y, por ende,
            # los mismos respaldos DOCX/detalles. No se crea una segunda copia.
            return existente, False

    existente = buscar_envio_notaria_preparando(id_raiz)
    if existente:
        return existente, False

    fila = {
        "ID_ENVIO_NOTARIA": nuevo_id(),
        "ID_DOCUMENTO_RAIZ": id_raiz,
        "ID_NOTARIA": id_notaria,
        "ID_PROPIEDAD_PRIME": id_propiedad_prime,
        "DESTINATARIOS_FIRMANTES": ", ".join(firmantes),
        "FECHA_ENVIO": fecha,
        "ENVIADO_POR": usuario,
        "MENSAJE_ADICIONAL": mensaje_adicional,
        "GMAIL_MESSAGE_ID": "",
        "GMAIL_THREAD_ID": "",
        "ESTADO_ENVIO": "Preparando",
        "NOMBRE_NOTARIA": datos_notaria["nombre"],
        "DIRECCION_NOTARIA": datos_notaria["direccion"],
        "PERSONA_CONTACTO_NOTARIA": datos_notaria["contacto"],
        "EMAIL_NOTARIA": datos_notaria["email"],
        "HORARIO_ATENCION_NOTARIA": datos_notaria["horario"],
    }
    appsheet_action(TABLA_ENVIOS_NOTARIA, "Add", [fila])
    return fila, True

def asegurar_detalle_envio_notaria(
    *,
    id_envio_notaria: str,
    tipo_adjunto: str,
    id_documento: str = "",
    id_version: str = "",
    id_documento_prime: str = "",
    nombre_archivo: str,
    drive_file_id: str = "",
    drive_file_url: str = "",
    fecha: str,
) -> None:
    condiciones = [
        f"[ID_ENVIO_NOTARIA] = {literal_appsheet(id_envio_notaria)}",
        f"[TIPO_ADJUNTO] = {literal_appsheet(tipo_adjunto)}",
        f"[NOMBRE_ARCHIVO] = {literal_appsheet(nombre_archivo)}",
    ]
    if id_documento:
        condiciones.append(f"[ID_DOCUMENTO] = {literal_appsheet(id_documento)}")
    if id_documento_prime:
        condiciones.append(
            f"[ID_DOCUMENTO_PRIME] = {literal_appsheet(id_documento_prime)}"
        )
    selector = (
        f"FILTER({TABLA_ENVIOS_NOTARIA_DETALLE}, AND("
        + ", ".join(condiciones)
        + "))"
    )
    if appsheet_find(TABLA_ENVIOS_NOTARIA_DETALLE, selector):
        return

    appsheet_action(
        TABLA_ENVIOS_NOTARIA_DETALLE,
        "Add",
        [
            {
                "ID_DETALLE_ENVIO_NOTARIA": nuevo_id(),
                "ID_ENVIO_NOTARIA": id_envio_notaria,
                "TIPO_ADJUNTO": tipo_adjunto,
                "ID_DOCUMENTO": id_documento,
                "ID_VERSION": id_version,
                "ID_DOCUMENTO_PRIME": id_documento_prime,
                "NOMBRE_ARCHIVO": nombre_archivo,
                "DRIVE_FILE_ID": drive_file_id,
                "DRIVE_FILE_URL": drive_file_url,
                "FECHA_CREACION": fecha,
            }
        ],
    )


def eliminar_detalle_envio_notaria_omitido(
    *,
    id_envio_notaria: str,
    tipo_adjunto: str,
    nombre_archivo: str,
    id_documento: str = "",
    id_documento_prime: str = "",
) -> None:
    """Elimina del detalle un archivo que finalmente NO será enviado.

    Es especialmente útil para limpiar detalles creados por versiones anteriores
    del backend antes de incorporar la exclusión automática por tamaño.
    """
    condiciones = [
        f"[ID_ENVIO_NOTARIA] = {literal_appsheet(id_envio_notaria)}",
        f"[TIPO_ADJUNTO] = {literal_appsheet(tipo_adjunto)}",
        f"[NOMBRE_ARCHIVO] = {literal_appsheet(nombre_archivo)}",
    ]
    if id_documento:
        condiciones.append(f"[ID_DOCUMENTO] = {literal_appsheet(id_documento)}")
    if id_documento_prime:
        condiciones.append(
            f"[ID_DOCUMENTO_PRIME] = {literal_appsheet(id_documento_prime)}"
        )
    selector = (
        f"FILTER({TABLA_ENVIOS_NOTARIA_DETALLE}, AND("
        + ", ".join(condiciones)
        + "))"
    )
    filas = appsheet_find(TABLA_ENVIOS_NOTARIA_DETALLE, selector)
    claves = [
        texto(fila.get("ID_DETALLE_ENVIO_NOTARIA"))
        for fila in filas
        if texto(fila.get("ID_DETALLE_ENVIO_NOTARIA"))
    ]
    if not claves:
        return
    appsheet_action(
        TABLA_ENVIOS_NOTARIA_DETALLE,
        "Delete",
        [{"ID_DETALLE_ENVIO_NOTARIA": clave} for clave in claves],
    )


def validar_paquete_para_envio_notaria(
    *,
    contexto: dict[str, Any],
    permitir_en_notaria: bool = False,
) -> list[dict[str, Any]]:
    if not contexto.get("solicitado_es_raiz"):
        raise ValueError(
            "Solo el documento raíz puede iniciar el envío a notaría. "
            f"Raíz: {contexto.get('id_documento_raiz')}"
        )
    if contexto.get("tipo_firma_paquete") != "Notarial":
        raise ValueError("El paquete no corresponde a Firma Notarial")

    raiz = contexto["raiz"]
    id_raiz = texto(contexto.get("id_documento_raiz"))
    estado_paquete = texto(raiz.get("ESTADO_PAQUETE_NOTARIAL"))
    estados_paquete_permitidos = {"Listo para notaría"}
    if permitir_en_notaria:
        estados_paquete_permitidos.add(ESTADO_PAQUETE_NOTARIAL_EN_NOTARIA)
    if estado_paquete not in estados_paquete_permitidos:
        raise ValueError(
            "El estado del paquete no permite el envío/reenvío a notaría. "
            f"Estado actual: {estado_paquete!r}; permitidos: "
            + ", ".join(sorted(estados_paquete_permitidos))
        )

    integrantes_notariales = buscar_integrantes_paquete_notarial(id_raiz)
    ids_notariales = {
        texto(fila.get("ID_DOCUMENTO")) for fila in integrantes_notariales
    }
    ids_jerarquia = {
        texto(fila.get("id_documento"))
        for fila in (contexto.get("integrantes") or [])
    }
    if ids_notariales != ids_jerarquia:
        raise ValueError(
            "La jerarquía actual no coincide con el paquete notarial activo"
        )

    preparados: list[dict[str, Any]] = []
    errores: list[str] = []
    for item in contexto.get("integrantes") or []:
        documento = item.get("documento") or {}
        id_documento = texto(documento.get("ID_DOCUMENTO"))
        titulo = texto(documento.get("TITULO")) or id_documento
        motivos: list[str] = []

        if texto(documento.get("ID_DOCUMENTO_RAIZ_NOTARIAL")) != id_raiz:
            motivos.append("no pertenece al paquete notarial activo")
        estado_documento = texto(documento.get("ESTADO"))
        estados_documento_permitidos = {"Listo para notaría"}
        if permitir_en_notaria:
            estados_documento_permitidos.add(ESTADO_DOCUMENTO_EN_NOTARIA)
        if estado_documento not in estados_documento_permitidos:
            motivos.append(
                f"ESTADO={estado_documento!r}; se requiere uno de "
                f"{sorted(estados_documento_permitidos)!r}"
            )
        resultado = normalizar_resultado_revision_externa(
            documento.get("RESULTADO_REVISION_EXTERNA")
        )
        if resultado != "Aprobado":
            motivos.append(
                f"RESULTADO_REVISION_EXTERNA={resultado!r}; se requiere 'Aprobado'"
            )

        id_version_actual = texto(documento.get("ID_VERSION_ACTUAL"))
        id_version_aprobada = texto(documento.get("ID_VERSION_REVISION_EXTERNA"))
        if not id_version_actual:
            motivos.append("falta ID_VERSION_ACTUAL")
        if not id_version_aprobada:
            motivos.append("falta ID_VERSION_REVISION_EXTERNA")
        if id_version_actual and id_version_aprobada and id_version_actual != id_version_aprobada:
            motivos.append("la versión actual no coincide con la versión aprobada externamente")

        version = None
        if id_version_aprobada:
            try:
                version = buscar_version_por_id(id_version_aprobada)
            except Exception as exc:
                motivos.append(f"no se pudo resolver la versión aprobada: {exc}")

        if version is not None:
            if texto(version.get("ID_DOCUMENTO")) != id_documento:
                motivos.append("ID_VERSION_REVISION_EXTERNA pertenece a otro documento")
            if texto(version.get("ESTADO_VERSION")) != "Activa":
                motivos.append("la versión aprobada no tiene ESTADO_VERSION='Activa'")
            if not texto(version.get("GOOGLE_DOC_ID")):
                motivos.append("la versión aprobada no tiene GOOGLE_DOC_ID")

        if motivos:
            errores.append(f"{titulo}: " + "; ".join(motivos))
            continue

        preparados.append(
            {
                "documento": documento,
                "version": version,
                "id_documento": id_documento,
                "titulo": titulo,
                "es_raiz": bool(item.get("es_raiz")),
                "id_version": id_version_aprobada,
            }
        )

    if errores:
        raise ValueError(
            "El paquete no está listo para envío a notaría: " + " | ".join(errores)
        )
    if not preparados:
        raise ValueError("El paquete notarial no contiene documentos para enviar")
    return preparados


@medir_operacion("notarial.construir_adjuntos_envio_notaria")
def construir_adjuntos_envio_notaria(
    *,
    drive_service: Any,
    preparados: list[dict[str, Any]],
    antecedentes_prime: list[dict[str, Any]],
    id_envio_notaria: str,
    fecha: str,
    limite_bytes: int = MAX_ADJUNTOS_NOTARIA_BYTES,
) -> tuple[
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
]:
    """Construye los adjuntos enviables y separa los archivos demasiado grandes.

    Los archivos cuyo tamaño individual supera ``limite_bytes`` NO se adjuntan al
    correo de notaría. Se devuelven en ``archivos_omitidos`` para dejar trazabilidad
    y notificar al encargado del proceso que debe remitirlos por otro medio.
    """
    if limite_bytes <= 0:
        raise ValueError("MAX_ADJUNTOS_NOTARIA_MB debe ser mayor que cero")

    adjuntos: list[dict[str, Any]] = []
    detalle_archivos: list[dict[str, Any]] = []
    archivos_omitidos: list[dict[str, Any]] = []

    def registrar_omitido(
        *,
        tipo: str,
        nombre: str,
        contenido: bytes,
        id_documento: str = "",
        id_version: str = "",
        id_documento_prime: str = "",
        titulo: str = "",
        descripcion: str = "",
    ) -> None:
        tamano = len(contenido)
        archivos_omitidos.append(
            {
                "tipo": tipo,
                "nombre": nombre,
                "tamano_bytes": tamano,
                "tamano_mb": round(tamano / (1024 * 1024), 2),
                "id_documento": id_documento,
                "id_version": id_version,
                "id_documento_prime": id_documento_prime,
                "titulo": titulo,
                "descripcion": descripcion,
                "motivo": (
                    f"Supera el máximo seguro de "
                    f"{limite_bytes / (1024 * 1024):.2f} MB por archivo."
                ),
            }
        )

    # 1) Documentos aprobados para firma: siempre editables DOCX.
    for item in preparados:
        version = item["version"]
        google_doc_id = texto(version.get("GOOGLE_DOC_ID"))
        nombre_base = texto(version.get("NOMBRE_ARCHIVO")) or item["titulo"]
        docx_bytes, docx_nombre = exportar_docx_drive(
            drive_service=drive_service,
            google_doc_id=google_doc_id,
            nombre_base=nombre_base,
        )

        if len(docx_bytes) > limite_bytes:
            registrar_omitido(
                tipo="Documento para firma",
                nombre=docx_nombre,
                contenido=docx_bytes,
                id_documento=item["id_documento"],
                id_version=item["id_version"],
                titulo=item["titulo"],
            )
            try:
                eliminar_detalle_envio_notaria_omitido(
                    id_envio_notaria=id_envio_notaria,
                    tipo_adjunto="Documento para firma",
                    nombre_archivo=docx_nombre,
                    id_documento=item["id_documento"],
                )
            except Exception:
                traceback.print_exc()
            continue

        respaldo = guardar_docx_enviado_notaria(
            drive_service=drive_service,
            google_doc_id=google_doc_id,
            contenido=docx_bytes,
            nombre_docx=docx_nombre,
            id_envio_notaria=id_envio_notaria,
        )
        adjuntos.append(
            {
                "contenido": docx_bytes,
                "nombre": docx_nombre,
                "maintype": "application",
                "subtype": "vnd.openxmlformats-officedocument.wordprocessingml.document",
            }
        )
        detalle = {
            "tipo": "Documento para firma",
            "id_documento": item["id_documento"],
            "id_version": item["id_version"],
            "titulo": item["titulo"],
            "nombre": docx_nombre,
            "drive_file_id": respaldo["id"],
            "drive_file_url": respaldo["url"],
        }
        detalle_archivos.append(detalle)
        asegurar_detalle_envio_notaria(
            id_envio_notaria=id_envio_notaria,
            tipo_adjunto="Documento para firma",
            id_documento=item["id_documento"],
            id_version=item["id_version"],
            nombre_archivo=docx_nombre,
            drive_file_id=respaldo["id"],
            drive_file_url=respaldo["url"],
            fecha=fecha,
        )

    # 2) Antecedentes complementarios de la propiedad.
    for fila in antecedentes_prime:
        valor_file = fila.get("file1")
        if not texto(valor_file):
            continue
        contenido, nombre, maintype, subtype = descargar_file_appsheet(
            table_name=TABLA_DOCUMENTOS_PRIME,
            valor_file=valor_file,
        )
        id_documento_prime = primer_valor(fila, "id_documento", "ID_DOCUMENTO")
        descripcion = primer_valor(fila, "Descripcion", "DESCRIPCION")

        if len(contenido) > limite_bytes:
            registrar_omitido(
                tipo="Antecedente",
                nombre=nombre,
                contenido=contenido,
                id_documento_prime=id_documento_prime,
                descripcion=descripcion,
            )
            try:
                eliminar_detalle_envio_notaria_omitido(
                    id_envio_notaria=id_envio_notaria,
                    tipo_adjunto="Antecedente",
                    nombre_archivo=nombre,
                    id_documento_prime=id_documento_prime,
                )
            except Exception:
                traceback.print_exc()
            continue

        adjuntos.append(
            {
                "contenido": contenido,
                "nombre": nombre,
                "maintype": maintype,
                "subtype": subtype,
            }
        )
        detalle_archivos.append(
            {
                "tipo": "Antecedente",
                "id_documento_prime": id_documento_prime,
                "nombre": nombre,
                "descripcion": descripcion,
            }
        )
        asegurar_detalle_envio_notaria(
            id_envio_notaria=id_envio_notaria,
            tipo_adjunto="Antecedente",
            id_documento_prime=id_documento_prime,
            nombre_archivo=nombre,
            fecha=fecha,
        )

    if not adjuntos:
        if archivos_omitidos:
            nombres = ", ".join(x["nombre"] for x in archivos_omitidos)
            raise ValueError(
                "Todos los archivos del paquete superan el límite individual de "
                f"{limite_bytes / (1024 * 1024):.2f} MB y fueron excluidos: {nombres}. "
                "No existe ningún archivo enviable a la notaría."
            )
        raise ValueError("No se reunieron archivos para enviar a notaría")

    return adjuntos, detalle_archivos, archivos_omitidos


def dividir_adjuntos_notaria(
    adjuntos: list[dict[str, Any]],
    *,
    limite_bytes: int = MAX_ADJUNTOS_NOTARIA_BYTES,
) -> list[dict[str, Any]]:
    """Agrupa adjuntos respetando un máximo de bytes reales por correo.

    El límite se aplica antes de la codificación MIME/base64. Mantenerlo en torno
    a 17 MiB deja margen suficiente para que el mensaje final no se acerque al
    límite habitual de Gmail.
    """
    if limite_bytes <= 0:
        raise ValueError("MAX_ADJUNTOS_NOTARIA_MB debe ser mayor que cero")
    if not adjuntos:
        raise ValueError("No existen adjuntos para dividir")

    lotes: list[dict[str, Any]] = []
    lote_actual: list[dict[str, Any]] = []
    indices_actuales: list[int] = []
    bytes_actuales = 0

    for indice, adjunto in enumerate(adjuntos):
        contenido = adjunto.get("contenido")
        nombre = texto(adjunto.get("nombre")) or f"adjunto_{indice + 1}"
        if not isinstance(contenido, bytes) or not contenido:
            raise ValueError(f"El adjunto {nombre} no contiene bytes válidos")

        tamano = len(contenido)
        if tamano > limite_bytes:
            raise ValueError(
                f"El archivo {nombre} pesa {tamano / (1024 * 1024):.2f} MB y supera "
                f"el máximo seguro de {limite_bytes / (1024 * 1024):.2f} MB para un "
                "único correo. Debe reducirse o enviarse mediante otro mecanismo."
            )

        if lote_actual and bytes_actuales + tamano > limite_bytes:
            lotes.append(
                {
                    "adjuntos": lote_actual,
                    "indices": indices_actuales,
                    "total_bytes": bytes_actuales,
                }
            )
            lote_actual = []
            indices_actuales = []
            bytes_actuales = 0

        lote_actual.append(adjunto)
        indices_actuales.append(indice)
        bytes_actuales += tamano

    if lote_actual:
        lotes.append(
            {
                "adjuntos": lote_actual,
                "indices": indices_actuales,
                "total_bytes": bytes_actuales,
            }
        )

    return lotes


def normalizar_lista_ids_gmail(valor: Any) -> list[str]:
    """Convierte uno o varios IDs Gmail guardados en texto a una lista limpia."""
    bruto = texto(valor)
    if not bruto:
        return []
    partes = re.split(r"[,;\n]+", bruto)
    return [x.strip() for x in partes if x.strip()]


def serializar_lista_ids_gmail(valores: list[str]) -> str:
    return ", ".join([texto(x) for x in valores if texto(x)])


def construir_email_envio_notaria(
    *,
    raiz: dict[str, Any],
    usuario: str,
    mensaje_adicional: str,
    datos_notaria: dict[str, str],
    detalle_archivos: list[dict[str, Any]],
    numero_parte: int = 1,
    total_partes: int = 1,
) -> tuple[str, str, str]:
    titulo = texto(raiz.get("TITULO")) or texto(raiz.get("ID_DOCUMENTO"))
    contacto = datos_notaria["contacto"]
    docs = [
        f"- {fila['nombre']}"
        for fila in detalle_archivos
        if fila.get("tipo") == "Documento para firma"
    ]
    antecedentes = [
        f"- {fila['nombre']}"
        for fila in detalle_archivos
        if fila.get("tipo") == "Antecedente"
    ]

    sufijo_parte = f" — Parte {numero_parte} de {total_partes}" if total_partes > 1 else ""
    asunto = f"Documentos para firma notarial — {titulo}{sufijo_parte}"

    bloques = [
        f"Estimado/a {contacto}:",
        "",
        f"Se remite el paquete documental asociado a {titulo} para continuar el proceso de firma notarial.",
    ]
    if total_partes > 1:
        bloques.extend(
            [
                "",
                f"IMPORTANTE: este correo corresponde a la PARTE {numero_parte} DE {total_partes} del envío.",
                "Revise todas las partes para contar con el paquete documental completo.",
            ]
        )

    bloques.extend(
        [
            "",
            "DOCUMENTOS PARA FIRMA / EDICIÓN NOTARIAL (DOCX)",
            *(docs or ["- En esta parte no se incluyen documentos para firma"]),
            "",
            "ANTECEDENTES COMPLEMENTARIOS (NO REQUIEREN FIRMA)",
            *(antecedentes or ["- En esta parte no se incluyen antecedentes complementarios"]),
        ]
    )
    if mensaje_adicional:
        bloques.extend(["", "INDICACIONES ADICIONALES", mensaje_adicional])
    bloques.extend(
        [
            "",
            "Los antecedentes complementarios se adjuntan únicamente como respaldo y no requieren firma.",
            "",
            "Agradeceremos devolver los documentos resultantes del proceso notarial respondiendo a este mismo correo.",
        ]
    )
    cuerpo = "\n".join(bloques)

    def esc(v: Any) -> str:
        return html.escape(texto(v))

    lista_docs = "".join(
        f"<li>{esc(fila['nombre'])}</li>"
        for fila in detalle_archivos
        if fila.get("tipo") == "Documento para firma"
    ) or "<li>En esta parte no se incluyen documentos para firma</li>"
    lista_ant = "".join(
        f"<li>{esc(fila['nombre'])}</li>"
        for fila in detalle_archivos
        if fila.get("tipo") == "Antecedente"
    ) or "<li>En esta parte no se incluyen antecedentes complementarios</li>"
    adicional_html = ""
    if mensaje_adicional:
        adicional_html = (
            f"<p><strong>Indicaciones adicionales:</strong><br>"
            f"{esc(mensaje_adicional).replace(chr(10), '<br>')}</p>"
        )
    bloque_parte_html = ""
    if total_partes > 1:
        bloque_parte_html = f"""
        <div style="margin:16px 0;padding:12px 14px;background:#fff7ed;border:1px solid #fdba74;border-radius:8px;">
          <strong>Parte {numero_parte} de {total_partes}</strong><br>
          Este envío fue dividido por tamaño. Revise todas las partes para contar con el paquete completo.
        </div>
        """

    cuerpo_html = f"""
    <html><body style="font-family:Arial,Helvetica,sans-serif;color:#111827;line-height:1.55;">
      <h2>Envío de documentos a notaría</h2>
      <p>Estimado/a <strong>{esc(contacto)}</strong>:</p>
      <p>Se remite el paquete documental asociado a <strong>{esc(titulo)}</strong> para continuar el proceso de firma notarial.</p>
      {bloque_parte_html}
      <h3>Documentos para firma / edición notarial (DOCX)</h3>
      <ul>{lista_docs}</ul>
      <h3>Antecedentes complementarios (no requieren firma)</h3>
      <ul>{lista_ant}</ul>
      {adicional_html}
      <p>Los antecedentes complementarios se adjuntan únicamente como respaldo y no requieren firma.</p>
      <p>Agradeceremos devolver los documentos resultantes del proceso notarial respondiendo a este mismo correo.</p>
    </body></html>
    """
    return asunto, cuerpo, cuerpo_html


def construir_email_firmante_notaria(
    *,
    raiz: dict[str, Any],
    datos_notaria: dict[str, str],
) -> tuple[str, str, str]:
    titulo = texto(raiz.get("TITULO")) or texto(raiz.get("ID_DOCUMENTO"))
    asunto = f"Documentación disponible para firma en notaría — {titulo}"
    cuerpo = (
        "Estimado/a:\n\n"
        "El proceso de revisión y aprobación de la documentación concluyó exitosamente. "
        "La documentación fue enviada a notaría para continuar con el proceso de firma.\n\n"
        f"Notaría: {datos_notaria['nombre']}\n"
        f"Dirección: {datos_notaria['direccion']}\n"
        f"Contacto: {datos_notaria['contacto']}\n"
        f"Correo: {datos_notaria['email']}\n"
        f"Horario de atención: {datos_notaria['horario']}\n\n"
        "Para coordinar su firma, favor considerar la información indicada.\n\n"
        "Esta notificación no contiene documentos adjuntos."
    )
    cuerpo_html = f"""
    <html><body style="font-family:Arial,Helvetica,sans-serif;color:#111827;line-height:1.6;">
      <h2>Documentación disponible para firma en notaría</h2>
      <p>Estimado/a:</p>
      <p>El proceso de revisión y aprobación de la documentación concluyó exitosamente. La documentación fue enviada a notaría para continuar con el proceso de firma.</p>
      <table style="border-collapse:collapse;">
        <tr><td style="padding:6px 12px 6px 0;font-weight:bold;">Notaría</td><td>{html.escape(datos_notaria['nombre'])}</td></tr>
        <tr><td style="padding:6px 12px 6px 0;font-weight:bold;">Dirección</td><td>{html.escape(datos_notaria['direccion'])}</td></tr>
        <tr><td style="padding:6px 12px 6px 0;font-weight:bold;">Contacto</td><td>{html.escape(datos_notaria['contacto'])}</td></tr>
        <tr><td style="padding:6px 12px 6px 0;font-weight:bold;">Correo</td><td>{html.escape(datos_notaria['email'])}</td></tr>
        <tr><td style="padding:6px 12px 6px 0;font-weight:bold;">Horario</td><td>{html.escape(datos_notaria['horario'])}</td></tr>
      </table>
      <p>Para coordinar su firma, favor considerar la información indicada.</p>
      <p><strong>Esta notificación no contiene documentos adjuntos.</strong></p>
    </body></html>
    """
    return asunto, cuerpo, cuerpo_html


def enviar_notificaciones_firmantes_notaria(
    *,
    gmail_service: Any,
    raiz: dict[str, Any],
    firmantes: list[str],
    datos_notaria: dict[str, str],
    id_envio_notaria: str,
) -> list[dict[str, Any]]:
    asunto, cuerpo, cuerpo_html = construir_email_firmante_notaria(
        raiz=raiz,
        datos_notaria=datos_notaria,
    )
    resultados: list[dict[str, Any]] = []
    for email in firmantes:
        try:
            rfc_message_id = construir_rfc_message_id_notificacion(
                f"notaria-{id_envio_notaria}-{email}"
            )
            respuesta = enviar_email_notificacion(
                gmail_service=gmail_service,
                destinatario=email,
                asunto=asunto,
                cuerpo_texto=cuerpo,
                cuerpo_html=cuerpo_html,
                rfc_message_id=rfc_message_id,
            )
            resultados.append(
                {
                    "ok": True,
                    "destinatario": email,
                    "message_id": respuesta["message_id"],
                    "thread_id": respuesta["thread_id"],
                }
            )
        except Exception as exc:
            traceback.print_exc()
            resultados.append(
                {
                    "ok": False,
                    "destinatario": email,
                    "error": str(exc),
                }
            )
    return resultados


def registrar_error_envio_notaria(id_documento: str, mensaje: str) -> None:
    try:
        appsheet_action(
            TABLA_DOCUMENTOS,
            "Edit",
            [
                {
                    "ID_DOCUMENTO": id_documento,
                    "ACCION_SOLICITADA": "",
                    "OBSERVACION_ACTUAL": texto(mensaje)[:1000],
                    "FECHA_ULTIMA_ACTUALIZACION": ahora_iso(),
                }
            ],
        )
    except Exception:
        traceback.print_exc()


def log_envio_notaria(evento: str, **campos: Any) -> None:
    payload = {
        "severity": "INFO",
        "message": "ENVIO_NOTARIA",
        "marca": "ENVIO_NOTARIA",
        "evento": evento,
        "fecha_chile": ahora_iso(),
        **campos,
    }
    try:
        payload.setdefault("solicitud_id", getattr(g, "solicitud_id", ""))
    except RuntimeError:
        pass
    print(json.dumps(payload, ensure_ascii=False, default=str), flush=True)


def registrar_checkpoint_gmail_notaria(
    *,
    id_documento_raiz: str,
    usuario: str,
    destinatario_notaria: str,
    mensaje_adicional: str,
    message_id: str,
    thread_id: str,
    fecha: str,
) -> None:
    """Persiste inmediatamente la aceptación de Gmail y limpia el trigger.

    Este checkpoint se ejecuta antes de cualquier trazabilidad secundaria. Si una
    escritura posterior falla, un retry automático del Bot no vuelve a enviar el
    correo porque ACCION_SOLICITADA ya quedó limpia y el message_id quedó guardado.
    """
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento_raiz,
                "DESTINATARIOS_NOTARIA": destinatario_notaria,
                "MENSAJE_ADICIONAL_NOTARIA": mensaje_adicional,
                "ENVIADO_NOTARIA_POR": usuario,
                "FECHA_ENVIO_NOTARIA": fecha,
                "EMAIL_NOTARIA_MESSAGE_ID": message_id,
                "GMAIL_THREAD_ID_NOTARIA": thread_id,
                "ULTIMO_ENVIADO_POR": usuario,
                "FECHA_ULTIMO_ENVIO": fecha,
                "ACCION_SOLICITADA": "",
                "OBSERVACION_ACTUAL": "",
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
            }
        ],
    )


def actualizar_envio_notaria_enviado(
    *,
    id_envio_notaria: str,
    message_id: str,
    thread_id: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_ENVIOS_NOTARIA,
        "Edit",
        [
            {
                "ID_ENVIO_NOTARIA": id_envio_notaria,
                "GMAIL_MESSAGE_ID": message_id,
                "GMAIL_THREAD_ID": thread_id,
                "FECHA_ENVIO": fecha,
                "ESTADO_ENVIO": "Enviado",
            }
        ],
    )


def actualizar_progreso_envio_notaria(
    *,
    id_envio_notaria: str,
    message_ids: list[str],
    thread_ids: list[str],
    fecha: str,
) -> None:
    """Guarda el progreso de un envío multipartes sin marcarlo todavía Enviado."""
    appsheet_action(
        TABLA_ENVIOS_NOTARIA,
        "Edit",
        [
            {
                "ID_ENVIO_NOTARIA": id_envio_notaria,
                "GMAIL_MESSAGE_ID": serializar_lista_ids_gmail(message_ids),
                "GMAIL_THREAD_ID": serializar_lista_ids_gmail(thread_ids),
                "FECHA_ENVIO": fecha,
                "ESTADO_ENVIO": "Preparando",
            }
        ],
    )


def reiniciar_progreso_reenvio_notaria(
    *,
    id_envio_notaria: str,
    fecha: str,
) -> None:
    """Inicia un nuevo intento explícito sin crear otra cabecera ni respaldos DOCX."""
    appsheet_action(
        TABLA_ENVIOS_NOTARIA,
        "Edit",
        [
            {
                "ID_ENVIO_NOTARIA": id_envio_notaria,
                "GMAIL_MESSAGE_ID": "",
                "GMAIL_THREAD_ID": "",
                "FECHA_ENVIO": fecha,
                "ESTADO_ENVIO": "Preparando",
            }
        ],
    )


def actualizar_paquete_tras_envio_notaria(
    *,
    id_documento_raiz: str,
    preparados: list[dict[str, Any]],
    usuario: str,
    datos_notaria: dict[str, str],
    mensaje_adicional: str,
    message_id: str,
    thread_id: str,
    fecha: str,
) -> None:
    filas: list[dict[str, Any]] = []
    for item in preparados:
        id_documento = item["id_documento"]
        cambios: dict[str, Any] = {
            "ID_DOCUMENTO": id_documento,
            "ESTADO": ESTADO_DOCUMENTO_EN_NOTARIA,
            "FECHA_ULTIMA_ACTUALIZACION": fecha,
            "OBSERVACION_ACTUAL": "",
        }
        if id_documento == id_documento_raiz:
            cambios.update(
                {
                    "ESTADO_PAQUETE_NOTARIAL": ESTADO_PAQUETE_NOTARIAL_EN_NOTARIA,
                    "DESTINATARIOS_NOTARIA": datos_notaria["email"],
                    "MENSAJE_ADICIONAL_NOTARIA": mensaje_adicional,
                    "ENVIADO_NOTARIA_POR": usuario,
                    "FECHA_ENVIO_NOTARIA": fecha,
                    "EMAIL_NOTARIA_MESSAGE_ID": message_id,
                    "GMAIL_THREAD_ID_NOTARIA": thread_id,
                    "ULTIMO_ENVIADO_POR": usuario,
                    "FECHA_ULTIMO_ENVIO": fecha,
                    "ACCION_SOLICITADA": "",
                }
            )
        filas.append(cambios)
    appsheet_action(TABLA_DOCUMENTOS, "Edit", filas)


def crear_eventos_envio_notaria(
    *,
    preparados: list[dict[str, Any]],
    id_documento_raiz: str,
    usuario: str,
    fecha: str,
    datos_notaria: dict[str, str],
    id_envio_notaria: str,
    message_id: str,
) -> dict[str, Any]:
    evento_raiz: dict[str, Any] | None = None
    eventos: list[dict[str, Any]] = []
    for item in preparados:
        es_raiz = item["id_documento"] == id_documento_raiz
        evento = {
            "ID_EVENTO": nuevo_id(),
            "ID_DOCUMENTO": item["id_documento"],
            "ID_VERSION": item["id_version"],
            "ID_APROBACION_ACTUAL": texto(item["documento"].get("ID_APROBACION_ACTUAL")),
            "TIPO_EVENTO": "Enviado a notaría" if es_raiz else "Incluido en envío a notaría",
            "ESTADO_ANTERIOR": "Listo para notaría",
            "ESTADO_NUEVO": ESTADO_DOCUMENTO_EN_NOTARIA,
            "USUARIO": usuario,
            "FECHA_EVENTO": fecha,
            "COMENTARIO": (
                f"Envío {id_envio_notaria} a {datos_notaria['nombre']} "
                f"({datos_notaria['email']}). Gmail message ID: {message_id}."
            ),
        }
        existentes = [
            e for e in buscar_eventos_documento(item["id_documento"])
            if id_envio_notaria in texto(e.get("COMENTARIO"))
            and texto(e.get("TIPO_EVENTO")) == evento["TIPO_EVENTO"]
        ]
        if existentes:
            if es_raiz:
                evento_raiz = existentes[-1]
            continue
        eventos.append(evento)
        if es_raiz:
            evento_raiz = evento
    if eventos:
        appsheet_action(TABLA_EVENTOS, "Add", eventos)
    return evento_raiz or (eventos[0] if eventos else {})


def notificar_encargado_archivos_omitidos_notaria(
    *,
    gmail_service: Any,
    raiz: dict[str, Any],
    usuario: str,
    datos_notaria: dict[str, str],
    id_envio_notaria: str,
    archivos_omitidos: list[dict[str, Any]],
) -> dict[str, Any]:
    """Avisa al encargado del proceso sobre archivos excluidos por tamaño.

    Prioridad del destinatario:
    1) Responsable de firmas de la versión vigente del documento raíz.
    2) ENCARGADO_ACTUAL_EMAIL del documento raíz.
    3) Usuario que ejecutó el envío a notaría.

    Esta notificación es informativa y no bloquea el envío de los demás archivos.
    """
    if not archivos_omitidos:
        return {"ok": True, "omitida": True, "motivo": "Sin archivos omitidos"}

    destinatario = ""
    nombre_destinatario = ""
    fuente_destinatario = ""

    try:
        responsable = obtener_responsable_firmas_documento_notarial(raiz)
        email_responsable = obtener_email_notificacion(responsable)
        if email_responsable and _EMAIL_RE.fullmatch(email_responsable):
            destinatario = email_responsable
            nombre_destinatario = texto(responsable.get("NOMBRE")) or "Responsable de firmas"
            fuente_destinatario = "Responsable de firmas"
    except Exception:
        traceback.print_exc()

    if not destinatario:
        email_encargado = texto(raiz.get("ENCARGADO_ACTUAL_EMAIL")).strip().lower()
        if email_encargado and _EMAIL_RE.fullmatch(email_encargado):
            destinatario = email_encargado
            nombre_destinatario = texto(raiz.get("ENCARGADO_ACTUAL_NOMBRE")) or "Encargado del proceso"
            fuente_destinatario = "ENCARGADO_ACTUAL_EMAIL"

    if not destinatario:
        email_usuario = texto(usuario).strip().lower()
        if email_usuario and _EMAIL_RE.fullmatch(email_usuario):
            destinatario = email_usuario
            nombre_destinatario = "Encargado del proceso"
            fuente_destinatario = "Usuario que ejecutó el envío"

    if not destinatario:
        return {
            "ok": False,
            "omitida": True,
            "error": "No se encontró un correo válido para notificar al encargado del proceso.",
        }

    titulo = texto(raiz.get("TITULO")) or texto(raiz.get("ID_DOCUMENTO"))
    cantidad = len(archivos_omitidos)
    limite = f"{MAX_ADJUNTOS_NOTARIA_MB:.2f} MB"

    lineas = []
    filas_html = []
    for archivo in archivos_omitidos:
        nombre = texto(archivo.get("nombre"))
        tamano_mb = float(archivo.get("tamano_mb") or 0)
        lineas.append(f"- {nombre} ({tamano_mb:.2f} MB)")
        filas_html.append(
            "<li><strong>" + html.escape(nombre) + "</strong> "
            + f"({tamano_mb:.2f} MB)</li>"
        )

    if cantidad == 1:
        asunto = f"Archivo no enviado a notaría por tamaño — {titulo}"
        frase_exclusion = (
            f"se excluyó 1 archivo porque supera el límite seguro de {limite} por archivo."
        )
        instruccion = (
            "El archivo indicado no fue enviado a la notaría y debe ser remitido "
            "a través de otro medio."
        )
    else:
        asunto = f"Archivos no enviados a notaría por tamaño — {titulo}"
        frase_exclusion = (
            f"se excluyeron {cantidad} archivos porque superan el límite seguro "
            f"de {limite} por archivo."
        )
        instruccion = (
            "Los archivos indicados no fueron enviados a la notaría y deben ser "
            "remitidos a través de otro medio."
        )

    cuerpo_texto = (
        f"Estimado/a {nombre_destinatario}:\n\n"
        f"Durante el envío del paquete '{titulo}' a {datos_notaria['nombre']}, "
        f"{frase_exclusion}\n\n"
        + "\n".join(lineas)
        + f"\n\n{instruccion}"
        + "\n\nLos demás archivos que cumplen el límite continúan su envío normal."
    )

    cuerpo_html = f"""
    <html><body style="font-family:Arial,Helvetica,sans-serif;color:#111827;line-height:1.6;">
      <h2>{html.escape(asunto)}</h2>
      <p>Estimado/a {html.escape(nombre_destinatario)}:</p>
      <p>Durante el envío del paquete <strong>{html.escape(titulo)}</strong> a
      <strong>{html.escape(datos_notaria['nombre'])}</strong>, {html.escape(frase_exclusion)}</p>
      <ul>{''.join(filas_html)}</ul>
      <p><strong>{html.escape(instruccion)}</strong></p>
      <p>Los demás archivos que cumplen el límite continúan su envío normal.</p>
    </body></html>
    """

    respuesta = enviar_email_notificacion(
        gmail_service=gmail_service,
        destinatario=destinatario,
        asunto=asunto,
        cuerpo_texto=cuerpo_texto,
        cuerpo_html=cuerpo_html,
        rfc_message_id=construir_rfc_message_id_notificacion(
            f"notaria-omitidos-{id_envio_notaria}-{nuevo_id()}"
        ),
    )
    return {
        "ok": True,
        "omitida": False,
        "destinatario": destinatario,
        "fuente_destinatario": fuente_destinatario,
        "cantidad_archivos": cantidad,
        "message_id": respuesta.get("message_id", ""),
        "thread_id": respuesta.get("thread_id", ""),
    }


@app.route("/enviar-notaria", methods=["POST"])
@app.route("/reenviar-notaria", methods=["POST"])
def enviar_notaria():
    """Envía/reenvía el paquete notarial dividiéndolo automáticamente por tamaño.

    - /enviar-notaria: envío inicial e idempotente.
    - /reenviar-notaria: reenvío manual explícito cuando el paquete ya está En notaría.

    El progreso se persiste por cada correo aceptado por Gmail en la cabecera
    Documento_Envios_Notaria. Si el flujo se interrumpe, el siguiente intento
    continúa desde la parte faltante mientras la cabecera permanezca Preparando.
    """
    id_documento_raiz = ""
    correo_notaria_aceptado_gmail = False
    message_id = ""
    thread_id = ""
    try:
        validar_configuracion()
        validar_token()
        data = request.get_json(silent=True) or {}

        id_documento_raiz = texto(
            data.get("id_documento_raiz") or data.get("id_documento")
        )
        id_notaria = texto(data.get("id_notaria"))
        id_propiedad_prime = texto(data.get("id_propiedad_prime"))
        usuario = texto(data.get("usuario"))
        firmantes_entrada = data.get("firmantes")
        mensaje_adicional = texto(data.get("mensaje_adicional"))
        reenvio_explicito = (
            request.path.rstrip("/").endswith("/reenviar-notaria")
            or es_verdadero(data.get("forzar_reenvio"))
        )

        if not id_documento_raiz:
            return {"error": "Falta id_documento_raiz"}, 400

        contexto = obtener_contexto_jerarquia_documental(id_documento_raiz)
        raiz = contexto["raiz"]
        id_raiz = texto(contexto.get("id_documento_raiz"))
        estado_paquete_actual = texto(raiz.get("ESTADO_PAQUETE_NOTARIAL"))
        message_id_existente = texto(raiz.get("EMAIL_NOTARIA_MESSAGE_ID"))

        if (
            not reenvio_explicito
            and estado_paquete_actual == ESTADO_PAQUETE_NOTARIAL_EN_NOTARIA
            and message_id_existente
        ):
            return jsonify(
                {
                    "ok": True,
                    "ya_procesado": True,
                    "id_documento_raiz": id_raiz,
                    "estado_paquete_notarial": estado_paquete_actual,
                    "message_id": message_id_existente,
                    "message_ids": normalizar_lista_ids_gmail(message_id_existente),
                    "thread_id": texto(raiz.get("GMAIL_THREAD_ID_NOTARIA")),
                    "fecha_envio_notaria": texto(raiz.get("FECHA_ENVIO_NOTARIA")),
                    "mensaje": "El envío ya estaba registrado; no se duplicó el correo.",
                }
            )

        if reenvio_explicito and estado_paquete_actual != ESTADO_PAQUETE_NOTARIAL_EN_NOTARIA:
            raise ValueError(
                "El reenvío controlado solo se permite cuando el paquete ya está 'En notaría'. "
                f"Estado actual: {estado_paquete_actual!r}"
            )

        preparados = validar_paquete_para_envio_notaria(
            contexto=contexto,
            permitir_en_notaria=reenvio_explicito,
        )

        id_notaria = id_notaria or texto(raiz.get("ID_NOTARIA"))
        id_propiedad_prime = id_propiedad_prime or texto(
            raiz.get("ID_PROPIEDAD_PRIME_NOTARIA")
        )
        if not id_notaria:
            raise ValueError("No se seleccionó ID_NOTARIA")
        if not id_propiedad_prime:
            raise ValueError("No se seleccionó ID_PROPIEDAD_PRIME_NOTARIA")
        if texto(raiz.get("ID_NOTARIA")) and texto(raiz.get("ID_NOTARIA")) != id_notaria:
            raise ValueError("ID_NOTARIA del webhook no coincide con Documentos")
        if (
            texto(raiz.get("ID_PROPIEDAD_PRIME_NOTARIA"))
            and texto(raiz.get("ID_PROPIEDAD_PRIME_NOTARIA")) != id_propiedad_prime
        ):
            raise ValueError("ID_PROPIEDAD_PRIME_NOTARIA del webhook no coincide con Documentos")

        notaria = buscar_notaria_por_id(id_notaria)
        datos_notaria = normalizar_datos_notaria(notaria)

        if firmantes_entrada in (None, ""):
            firmantes_entrada = raiz.get("DESTINATARIOS_FIRMANTES_NOTARIA")
        firmantes = normalizar_destinatarios(firmantes_entrada)
        if not firmantes:
            raise ValueError("No se indicaron firmantes a notificar")

        usuario_registrado = texto(raiz.get("ULTIMO_ENVIADO_POR"))
        usuario = usuario or usuario_registrado
        if not usuario or not _EMAIL_RE.fullmatch(usuario.lower()):
            raise ValueError("El usuario que envía a notaría no es válido")

        if not mensaje_adicional:
            mensaje_adicional = texto(raiz.get("MENSAJE_ADICIONAL_NOTARIA"))

        antecedentes_prime = buscar_documentos_prime_propiedad(id_propiedad_prime)
        if not antecedentes_prime:
            raise ValueError(
                "La propiedad seleccionada no tiene registros en Documentos_prime"
            )

        fecha_envio = ahora_iso()
        envio, cabecera_creada = obtener_o_crear_envio_notaria(
            raiz=raiz,
            id_notaria=id_notaria,
            id_propiedad_prime=id_propiedad_prime,
            firmantes=firmantes,
            usuario=usuario,
            mensaje_adicional=mensaje_adicional,
            datos_notaria=datos_notaria,
            fecha=fecha_envio,
            reutilizar_ultimo=reenvio_explicito,
        )
        id_envio_notaria = texto(envio.get("ID_ENVIO_NOTARIA"))
        if not id_envio_notaria:
            raise RuntimeError("No se pudo obtener ID_ENVIO_NOTARIA")

        # Un reenvío explícito inicia una nueva serie de partes, salvo que una
        # serie anterior haya quedado Preparando y deba continuarse.
        if reenvio_explicito and texto(envio.get("ESTADO_ENVIO")) != "Preparando":
            reiniciar_progreso_reenvio_notaria(
                id_envio_notaria=id_envio_notaria,
                fecha=fecha_envio,
            )
            envio["GMAIL_MESSAGE_ID"] = ""
            envio["GMAIL_THREAD_ID"] = ""
            envio["ESTADO_ENVIO"] = "Preparando"

        drive_service = obtener_drive_service()
        adjuntos, detalle_archivos, archivos_omitidos = construir_adjuntos_envio_notaria(
            drive_service=drive_service,
            preparados=preparados,
            antecedentes_prime=antecedentes_prime,
            id_envio_notaria=id_envio_notaria,
            fecha=fecha_envio,
            limite_bytes=MAX_ADJUNTOS_NOTARIA_BYTES,
        )

        lotes = dividir_adjuntos_notaria(adjuntos)
        total_bytes = sum(len(a.get("contenido") or b"") for a in adjuntos)
        total_omitidos_bytes = sum(int(x.get("tamano_bytes") or 0) for x in archivos_omitidos)
        nombres_adjuntos = [texto(a.get("nombre")) for a in adjuntos]
        resumen_partes = [
            {
                "parte": i + 1,
                "cantidad_adjuntos": len(lote["adjuntos"]),
                "total_mb": round(lote["total_bytes"] / (1024 * 1024), 2),
                "archivos": [texto(x.get("nombre")) for x in lote["adjuntos"]],
            }
            for i, lote in enumerate(lotes)
        ]

        ids_enviados = normalizar_lista_ids_gmail(envio.get("GMAIL_MESSAGE_ID"))
        threads_enviados = normalizar_lista_ids_gmail(envio.get("GMAIL_THREAD_ID"))
        partes_ya_enviadas = len(ids_enviados)
        if partes_ya_enviadas > len(lotes):
            raise ValueError(
                "La cabecera de envío registra más correos que las partes calculadas; "
                "revisa Documento_Envios_Notaria antes de continuar."
            )

        log_envio_notaria(
            "antes_gmail",
            id_documento_raiz=id_raiz,
            id_envio_notaria=id_envio_notaria,
            reenvio_explicito=reenvio_explicito,
            destinatario=datos_notaria["email"],
            cantidad_adjuntos=len(adjuntos),
            total_adjuntos_bytes=total_bytes,
            total_adjuntos_mb=round(total_bytes / (1024 * 1024), 2),
            cantidad_adjuntos_omitidos=len(archivos_omitidos),
            total_omitidos_mb=round(total_omitidos_bytes / (1024 * 1024), 2),
            archivos_omitidos=[
                {
                    "nombre": texto(x.get("nombre")),
                    "tamano_mb": x.get("tamano_mb"),
                    "tipo": texto(x.get("tipo")),
                }
                for x in archivos_omitidos
            ],
            limite_por_correo_mb=MAX_ADJUNTOS_NOTARIA_MB,
            cantidad_partes=len(lotes),
            partes_ya_enviadas=partes_ya_enviadas,
            nombres_adjuntos=nombres_adjuntos,
            resumen_partes=resumen_partes,
        )

        gmail_service = obtener_gmail_service()
        respuestas_partes: list[dict[str, Any]] = []

        # Conserva en la respuesta las partes que ya estaban persistidas.
        for i in range(partes_ya_enviadas):
            respuestas_partes.append(
                {
                    "parte": i + 1,
                    "reanudada": True,
                    "message_id": ids_enviados[i],
                    "thread_id": threads_enviados[i] if i < len(threads_enviados) else "",
                }
            )

        for indice in range(partes_ya_enviadas, len(lotes)):
            lote = lotes[indice]
            numero_parte = indice + 1
            detalles_lote = [detalle_archivos[i] for i in lote["indices"]]
            asunto, cuerpo, cuerpo_html = construir_email_envio_notaria(
                raiz=raiz,
                usuario=usuario,
                mensaje_adicional=mensaje_adicional,
                datos_notaria=datos_notaria,
                detalle_archivos=detalles_lote,
                numero_parte=numero_parte,
                total_partes=len(lotes),
            )

            log_envio_notaria(
                "antes_gmail_parte",
                id_documento_raiz=id_raiz,
                id_envio_notaria=id_envio_notaria,
                parte=numero_parte,
                total_partes=len(lotes),
                destinatario=datos_notaria["email"],
                asunto=asunto,
                cantidad_adjuntos=len(lote["adjuntos"]),
                total_mb=round(lote["total_bytes"] / (1024 * 1024), 2),
                archivos=[texto(x.get("nombre")) for x in lote["adjuntos"]],
            )

            try:
                respuesta = enviar_email_con_adjuntos(
                    gmail_service=gmail_service,
                    destinatarios=[datos_notaria["email"]],
                    asunto=asunto,
                    cuerpo=cuerpo,
                    adjuntos=lote["adjuntos"],
                    reply_to=usuario,
                    cuerpo_html=cuerpo_html,
                )
            except Exception as exc_parte:
                traceback.print_exc()
                mensaje_error = (
                    f"Falló el envío de la parte {numero_parte} de {len(lotes)} a notaría: "
                    f"{exc_parte}"
                )
                registrar_error_envio_notaria(id_raiz, mensaje_error)
                log_envio_notaria(
                    "error_gmail_parte",
                    id_documento_raiz=id_raiz,
                    id_envio_notaria=id_envio_notaria,
                    parte=numero_parte,
                    total_partes=len(lotes),
                    partes_enviadas=len(ids_enviados),
                    error=str(exc_parte),
                )
                if ids_enviados:
                    return jsonify(
                        {
                            "ok": False,
                            "envio_parcial": True,
                            "id_documento_raiz": id_raiz,
                            "id_envio_notaria": id_envio_notaria,
                            "parte_fallida": numero_parte,
                            "total_partes": len(lotes),
                            "partes_enviadas": len(ids_enviados),
                            "message_ids": ids_enviados,
                            "error": str(exc_parte),
                            "mensaje": (
                                "El envío quedó parcial. Al ejecutar nuevamente la misma acción, "
                                "el backend continuará desde la parte faltante."
                            ),
                        }
                    ), 200
                raise

            correo_notaria_aceptado_gmail = True
            message_id = respuesta["message_id"]
            thread_id = respuesta.get("thread_id", "")
            label_ids = respuesta.get("label_ids") or []
            ids_enviados.append(message_id)
            threads_enviados.append(thread_id)

            try:
                actualizar_progreso_envio_notaria(
                    id_envio_notaria=id_envio_notaria,
                    message_ids=ids_enviados,
                    thread_ids=threads_enviados,
                    fecha=fecha_envio,
                )
            except Exception as exc_progreso:
                traceback.print_exc()
                mensaje_error = (
                    f"Gmail aceptó la parte {numero_parte}, pero no fue posible persistir "
                    f"el progreso del envío: {exc_progreso}"
                )
                registrar_error_envio_notaria(id_raiz, mensaje_error)
                log_envio_notaria(
                    "error_persistencia_parte",
                    id_documento_raiz=id_raiz,
                    id_envio_notaria=id_envio_notaria,
                    parte=numero_parte,
                    message_id=message_id,
                    error=str(exc_progreso),
                )
                return jsonify(
                    {
                        "ok": False,
                        "envio_parcial": True,
                        "gmail_acepto_correo": True,
                        "parte": numero_parte,
                        "message_id": message_id,
                        "error": str(exc_progreso),
                        "mensaje": (
                            "Gmail aceptó esta parte, pero el progreso no pudo guardarse. "
                            "No reintentes automáticamente hasta revisar Documento_Envios_Notaria."
                        ),
                    }
                ), 200

            respuestas_partes.append(
                {
                    "parte": numero_parte,
                    "reanudada": False,
                    "message_id": message_id,
                    "thread_id": thread_id,
                    "label_ids": label_ids,
                    "cantidad_adjuntos": len(lote["adjuntos"]),
                    "total_mb": round(lote["total_bytes"] / (1024 * 1024), 2),
                }
            )
            log_envio_notaria(
                "gmail_aceptado_parte",
                id_documento_raiz=id_raiz,
                id_envio_notaria=id_envio_notaria,
                parte=numero_parte,
                total_partes=len(lotes),
                destinatario=datos_notaria["email"],
                message_id=message_id,
                thread_id=thread_id,
                label_ids=label_ids,
            )

        if len(ids_enviados) != len(lotes):
            raise RuntimeError(
                f"Se enviaron {len(ids_enviados)} de {len(lotes)} partes; el paquete no puede cerrarse."
            )

        message_ids_texto = serializar_lista_ids_gmail(ids_enviados)
        thread_ids_texto = serializar_lista_ids_gmail(threads_enviados)
        message_id = ids_enviados[0] if ids_enviados else ""
        thread_id = threads_enviados[0] if threads_enviados else ""
        correo_notaria_aceptado_gmail = bool(ids_enviados)
        advertencias: list[str] = []

        # Solo después de completar TODAS las partes se marca el paquete En notaría.
        try:
            registrar_checkpoint_gmail_notaria(
                id_documento_raiz=id_raiz,
                usuario=usuario,
                destinatario_notaria=datos_notaria["email"],
                mensaje_adicional=mensaje_adicional,
                message_id=message_ids_texto,
                thread_id=thread_ids_texto,
                fecha=fecha_envio,
            )
        except Exception as exc_checkpoint:
            traceback.print_exc()
            advertencias.append(
                "Todas las partes fueron aceptadas por Gmail, pero falló el checkpoint en Documentos: "
                + str(exc_checkpoint)
            )

        try:
            actualizar_envio_notaria_enviado(
                id_envio_notaria=id_envio_notaria,
                message_id=message_ids_texto,
                thread_id=thread_ids_texto,
                fecha=fecha_envio,
            )
        except Exception as exc_envio:
            traceback.print_exc()
            advertencias.append(
                "Todas las partes fueron aceptadas por Gmail, pero no fue posible cerrar Documento_Envios_Notaria: "
                + str(exc_envio)
            )

        paquete_actualizado = False
        try:
            actualizar_paquete_tras_envio_notaria(
                id_documento_raiz=id_raiz,
                preparados=preparados,
                usuario=usuario,
                datos_notaria=datos_notaria,
                mensaje_adicional=mensaje_adicional,
                message_id=message_ids_texto,
                thread_id=thread_ids_texto,
                fecha=fecha_envio,
            )
            paquete_actualizado = True
        except Exception as exc_paquete:
            traceback.print_exc()
            advertencias.append(
                "Todas las partes fueron aceptadas por Gmail, pero falló la actualización del estado del paquete: "
                + str(exc_paquete)
            )

        evento_raiz: dict[str, Any] = {}
        try:
            evento_raiz = crear_eventos_envio_notaria(
                preparados=preparados,
                id_documento_raiz=id_raiz,
                usuario=usuario,
                fecha=fecha_envio,
                datos_notaria=datos_notaria,
                id_envio_notaria=id_envio_notaria,
                message_id=(
                    f"{message_id} (+{len(ids_enviados)-1} parte(s) adicional(es))"
                    if len(ids_enviados) > 1
                    else message_id
                ),
            )
        except Exception as exc_evento:
            traceback.print_exc()
            advertencias.append(
                "El correo fue aceptado por Gmail, pero no fue posible crear uno o más "
                "eventos de trazabilidad: " + str(exc_evento)
            )

        # Si algún archivo individual supera el límite, el envío continúa sin él
        # y se avisa al encargado del proceso para que lo remita por otro medio.
        notificacion_archivos_omitidos: dict[str, Any] = {}
        if archivos_omitidos:
            resumen_omitidos = ", ".join(
                f"{texto(x.get('nombre'))} ({float(x.get('tamano_mb') or 0):.2f} MB)"
                for x in archivos_omitidos
            )
            advertencias.append(
                f"Se omitieron {len(archivos_omitidos)} archivo(s) por superar "
                f"{MAX_ADJUNTOS_NOTARIA_MB:.2f} MB: {resumen_omitidos}. "
                "Deben enviarse a la notaría por otro medio."
            )
            try:
                notificacion_archivos_omitidos = notificar_encargado_archivos_omitidos_notaria(
                    gmail_service=gmail_service,
                    raiz=raiz,
                    usuario=usuario,
                    datos_notaria=datos_notaria,
                    id_envio_notaria=id_envio_notaria,
                    archivos_omitidos=archivos_omitidos,
                )
                if not notificacion_archivos_omitidos.get("ok"):
                    advertencias.append(
                        "No fue posible notificar al encargado sobre los archivos omitidos: "
                        + texto(notificacion_archivos_omitidos.get("error"))
                    )
                log_envio_notaria(
                    "notificacion_archivos_omitidos",
                    id_documento_raiz=id_raiz,
                    id_envio_notaria=id_envio_notaria,
                    cantidad_archivos=len(archivos_omitidos),
                    archivos=[
                        {
                            "nombre": texto(x.get("nombre")),
                            "tamano_mb": x.get("tamano_mb"),
                        }
                        for x in archivos_omitidos
                    ],
                    resultado=notificacion_archivos_omitidos,
                )
            except Exception as exc_omitidos:
                traceback.print_exc()
                notificacion_archivos_omitidos = {
                    "ok": False,
                    "error": str(exc_omitidos),
                }
                advertencias.append(
                    "El envío a notaría continuó, pero falló la notificación al encargado "
                    f"sobre los archivos omitidos: {exc_omitidos}"
                )

        # El reenvío explícito NO vuelve a notificar a los firmantes.
        notificaciones_firmantes: list[dict[str, Any]] = []
        if paquete_actualizado and not reenvio_explicito:
            notificaciones_firmantes = enviar_notificaciones_firmantes_notaria(
                gmail_service=gmail_service,
                raiz=raiz,
                firmantes=firmantes,
                datos_notaria=datos_notaria,
                id_envio_notaria=id_envio_notaria,
            )
            fallidas = [x for x in notificaciones_firmantes if not x.get("ok")]
            if fallidas:
                advertencias.append(
                    f"{len(fallidas)} notificación(es) a firmantes fallaron; "
                    "el envío a notaría ya fue completado."
                )
        elif not paquete_actualizado:
            advertencias.append(
                "No se notificó a los firmantes porque no fue posible confirmar "
                "la actualización del estado del paquete."
            )

        log_envio_notaria(
            "fin_endpoint",
            id_documento_raiz=id_raiz,
            id_envio_notaria=id_envio_notaria,
            message_id=message_id,
            message_ids=ids_enviados,
            cantidad_partes=len(lotes),
            paquete_actualizado=paquete_actualizado,
            reenvio_explicito=reenvio_explicito,
            advertencias=advertencias,
        )

        return jsonify(
            {
                "ok": True,
                "gmail_acepto_correo": True,
                "entrega_destinatario_confirmada": False,
                "ya_procesado": False,
                "reenvio_explicito": reenvio_explicito,
                "cabecera_creada": cabecera_creada,
                "id_documento_raiz": id_raiz,
                "id_envio_notaria": id_envio_notaria,
                "id_notaria": id_notaria,
                "notaria": datos_notaria["nombre"],
                "destinatario_notaria": datos_notaria["email"],
                "estado_paquete_notarial": (
                    ESTADO_PAQUETE_NOTARIAL_EN_NOTARIA
                    if paquete_actualizado
                    else estado_paquete_actual
                ),
                "message_id": message_id,
                "thread_id": thread_id,
                "message_ids": ids_enviados,
                "thread_ids": threads_enviados,
                "cantidad_partes": len(lotes),
                "partes": respuestas_partes,
                "cantidad_documentos_firma": len(preparados),
                "cantidad_antecedentes_prime": len(
                    [x for x in detalle_archivos if x.get("tipo") == "Antecedente"]
                ),
                "cantidad_adjuntos": len(adjuntos),
                "total_adjuntos_mb": round(total_bytes / (1024 * 1024), 2),
                "cantidad_adjuntos_omitidos": len(archivos_omitidos),
                "total_omitidos_mb": round(total_omitidos_bytes / (1024 * 1024), 2),
                "archivos_omitidos": [
                    {
                        "nombre": texto(x.get("nombre")),
                        "tamano_mb": x.get("tamano_mb"),
                        "tipo": texto(x.get("tipo")),
                        "motivo": texto(x.get("motivo")),
                    }
                    for x in archivos_omitidos
                ],
                "limite_por_correo_mb": MAX_ADJUNTOS_NOTARIA_MB,
                "notificacion_archivos_omitidos": notificacion_archivos_omitidos,
                "notificaciones_firmantes": notificaciones_firmantes,
                "evento_raiz": texto(evento_raiz.get("ID_EVENTO")),
                "advertencias": advertencias,
            }
        ), 200

    except PermissionError as exc:
        if id_documento_raiz and not correo_notaria_aceptado_gmail:
            registrar_error_envio_notaria(id_documento_raiz, str(exc))
        return {"error": str(exc)}, 403
    except (ValueError, LookupError) as exc:
        if id_documento_raiz and not correo_notaria_aceptado_gmail:
            registrar_error_envio_notaria(id_documento_raiz, str(exc))
        return {"error": str(exc)}, 400
    except Exception as exc:
        traceback.print_exc()
        if correo_notaria_aceptado_gmail:
            return jsonify(
                {
                    "ok": False,
                    "gmail_acepto_correo": True,
                    "entrega_destinatario_confirmada": False,
                    "error_post_gmail": str(exc),
                    "message_id": message_id,
                    "thread_id": thread_id,
                }
            ), 200
        if id_documento_raiz:
            registrar_error_envio_notaria(id_documento_raiz, str(exc))
        return {
            "error": str(exc),
            "gmail_acepto_correo": False,
            "message_id": message_id,
            "thread_id": thread_id,
        }, 500


# -----------------------------------------------------------------------------
# Flujo Notarial - devolución desde notaría + notificaciones
# -----------------------------------------------------------------------------


def buscar_registro_prime_origen_notaria(id_documento: str) -> dict[str, Any] | None:
    selector = (
        f"FILTER({TABLA_DOCUMENTOS_PRIME}, AND("
        f"[ID_DOCUMENTO_ORIGEN] = {literal_appsheet(id_documento)}, "
        f"[ORIGEN_DOCUMENTO] = {literal_appsheet('Notaría')}))"
    )
    filas = appsheet_find(TABLA_DOCUMENTOS_PRIME, selector)
    if not filas:
        return None
    return filas[0]


def crear_registro_prime_devolucion_notaria(
    *,
    documento: dict[str, Any],
    id_propiedad_prime: str,
    archivo_devuelto: str,
    fecha: str,
) -> dict[str, Any]:
    """Publica una sola vez el archivo devuelto por notaría en Documentos_Prime."""
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    existente = buscar_registro_prime_origen_notaria(id_documento)
    if existente is not None:
        return existente

    id_documento_prime = nuevo_id()
    titulo = texto(documento.get("TITULO")) or id_documento
    tipo_documento = texto(documento.get("TIPO_DOCUMENTO")) or "Documento"

    fila = {
        "id_documento": id_documento_prime,
        "id_Propiedades": id_propiedad_prime,
        "Tipo de Documento": tipo_documento,
        "Descripcion": f"Documento devuelto por notaría - {titulo}",
        "file1": archivo_devuelto,
        "Fecha_creacion": fecha,
        "ID_DOCUMENTO_ORIGEN": id_documento,
        "ORIGEN_DOCUMENTO": "Notaría",
    }
    appsheet_action(TABLA_DOCUMENTOS_PRIME, "Add", [fila])
    return fila


def actualizar_documento_tras_devolucion_notaria(
    *,
    id_documento: str,
    id_documento_prime: str,
    usuario: str,
    fecha: str,
) -> None:
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento,
                "ESTADO": ESTADO_DOCUMENTO_PROCESO_TERMINADO,
                "ID_DOCUMENTO_PRIME_GENERADO": id_documento_prime,
                "RECIBIDO_NOTARIA_POR": usuario,
                "FECHA_RECEPCION_NOTARIA": fecha,
                "ULTIMO_ENVIADO_POR": usuario,
                "FECHA_ULTIMO_ENVIO": fecha,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "FECHA_CIERRE": fecha,
                "OBSERVACION_ACTUAL": "",
                "ACCION_SOLICITADA": "",
            }
        ],
    )


def crear_evento_devolucion_notaria(
    *,
    documento: dict[str, Any],
    id_documento_prime: str,
    usuario: str,
    fecha: str,
) -> dict[str, Any]:
    id_documento = texto(documento.get("ID_DOCUMENTO"))
    comentario_prefijo = f"Documento devuelto por notaría; Prime={id_documento_prime}."
    existentes = [
        evento
        for evento in buscar_eventos_documento(id_documento)
        if texto(evento.get("TIPO_EVENTO")) == "Respuesta de firma externa"
        and texto(evento.get("COMENTARIO")).startswith(comentario_prefijo)
    ]
    if existentes:
        existentes.sort(
            key=lambda e: (
                parsear_fecha_appsheet(e.get("FECHA_EVENTO")),
                texto(e.get("ID_EVENTO")),
            )
        )
        return existentes[-1]

    evento = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": id_documento,
        "ID_VERSION": texto(documento.get("ID_VERSION_ACTUAL")),
        "ID_APROBACION_ACTUAL": texto(documento.get("ID_APROBACION_ACTUAL")),
        "TIPO_EVENTO": "Respuesta de firma externa",
        "ESTADO_ANTERIOR": ESTADO_DOCUMENTO_EN_NOTARIA,
        "ESTADO_NUEVO": ESTADO_DOCUMENTO_PROCESO_TERMINADO,
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": (
            comentario_prefijo
            + " El archivo recibido quedó registrado en el flujo documental "
            "y publicado en Documentos_Prime."
        ),
    }
    appsheet_action(TABLA_EVENTOS, "Add", [evento])
    return evento


def notificar_responsable_documento_devuelto_notaria(
    *,
    documento: dict[str, Any],
    evento: dict[str, Any],
) -> list[dict[str, Any]]:
    """Notificación interna por cada documento recibido desde notaría."""
    responsable = obtener_responsable_firmas_documento_notarial(documento)
    titulo = texto(documento.get("TITULO")) or texto(documento.get("ID_DOCUMENTO"))
    return notificar_destinatarios_internos(
        documento=documento,
        evento=evento,
        destinatarios=[
            {
                "aprobador": responsable,
                "tipo_notificacion": "Informativa",
                "movimiento": "Documento recibido desde notaría",
                "comentario_principal": (
                    f"Se recibió desde notaría el documento {titulo}. "
                    "El archivo quedó registrado en el flujo documental y en Documentos_Prime."
                ),
                "link_documento": "",
            }
        ],
    )


def paquete_notarial_devuelto_completo(id_documento_raiz: str) -> tuple[bool, list[dict[str, Any]]]:
    integrantes = buscar_integrantes_paquete_notarial(id_documento_raiz)
    if not integrantes:
        raise LookupError("El paquete Notarial no contiene documentos asociados")
    completo = all(
        bool(texto(fila.get("ID_DOCUMENTO_PRIME_GENERADO")))
        for fila in integrantes
    )
    return completo, integrantes


def cerrar_paquete_notarial_tras_devoluciones(
    *,
    id_documento_raiz: str,
    fecha: str,
) -> None:
    """Cierra el paquete una vez que todos sus documentos volvieron de notaría."""
    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [
            {
                "ID_DOCUMENTO": id_documento_raiz,
                "ESTADO_PAQUETE_NOTARIAL": ESTADO_PAQUETE_NOTARIAL_CERRADO,
                "FECHA_ULTIMA_ACTUALIZACION": fecha,
                "OBSERVACION_ACTUAL": "",
                "ACCION_SOLICITADA": "",
            }
        ],
    )


def buscar_evento_cierre_retorno_notaria(id_documento_raiz: str) -> dict[str, Any] | None:
    prefijo = "Paquete notarial completo: todos los documentos fueron recibidos desde notaría."
    candidatos = [
        evento
        for evento in buscar_eventos_documento(id_documento_raiz)
        if texto(evento.get("TIPO_EVENTO")) == "Proceso terminado"
        and texto(evento.get("COMENTARIO")).startswith(prefijo)
    ]
    if not candidatos:
        return None
    candidatos.sort(
        key=lambda e: (
            parsear_fecha_appsheet(e.get("FECHA_EVENTO")),
            texto(e.get("ID_EVENTO")),
        )
    )
    return candidatos[-1]


def crear_evento_cierre_retorno_notaria(
    *,
    raiz: dict[str, Any],
    usuario: str,
    fecha: str,
) -> tuple[dict[str, Any], bool]:
    existente = buscar_evento_cierre_retorno_notaria(texto(raiz.get("ID_DOCUMENTO")))
    if existente is not None:
        return existente, False

    evento = {
        "ID_EVENTO": nuevo_id(),
        "ID_DOCUMENTO": texto(raiz.get("ID_DOCUMENTO")),
        "ID_VERSION": texto(raiz.get("ID_VERSION_ACTUAL")),
        "ID_APROBACION_ACTUAL": texto(raiz.get("ID_APROBACION_ACTUAL")),
        "TIPO_EVENTO": "Proceso terminado",
        "ESTADO_ANTERIOR": ESTADO_DOCUMENTO_EN_NOTARIA,
        "ESTADO_NUEVO": ESTADO_DOCUMENTO_PROCESO_TERMINADO,
        "USUARIO": usuario,
        "FECHA_EVENTO": fecha,
        "COMENTARIO": (
            "Paquete notarial completo: todos los documentos fueron recibidos desde notaría. "
            "Las devoluciones quedaron incorporadas en Documentos_Prime."
        ),
    }
    appsheet_action(TABLA_EVENTOS, "Add", [evento])
    return evento, True


def construir_email_fin_paquete_notarial(
    *,
    raiz: dict[str, Any],
    datos_notaria: dict[str, str],
) -> tuple[str, str, str]:
    titulo = texto(raiz.get("TITULO")) or texto(raiz.get("ID_DOCUMENTO"))
    asunto = f"Proceso notarial concluido — {titulo}"
    cuerpo = (
        "Estimado/a:\n\n"
        "El proceso notarial de la documentación ha concluido. "
        "Todos los documentos asociados al paquete fueron recibidos de vuelta desde la notaría "
        "y registrados correctamente en el sistema.\n\n"
        f"Notaría: {datos_notaria['nombre']}\n"
        f"Documento principal: {titulo}\n\n"
        "No se adjuntan documentos a esta notificación."
    )
    cuerpo_html = f"""
    <html><body style="font-family:Arial,Helvetica,sans-serif;color:#111827;line-height:1.6;">
      <h2>Proceso notarial concluido</h2>
      <p>Estimado/a:</p>
      <p>El proceso notarial de la documentación ha concluido. Todos los documentos asociados al paquete fueron recibidos de vuelta desde la notaría y registrados correctamente en el sistema.</p>
      <table style="border-collapse:collapse;">
        <tr><td style="padding:6px 12px 6px 0;font-weight:bold;">Documento principal</td><td>{html.escape(titulo)}</td></tr>
        <tr><td style="padding:6px 12px 6px 0;font-weight:bold;">Notaría</td><td>{html.escape(datos_notaria['nombre'])}</td></tr>
      </table>
      <p><strong>No se adjuntan documentos a esta notificación.</strong></p>
    </body></html>
    """
    return asunto, cuerpo, cuerpo_html


def enviar_notificaciones_fin_paquete_notarial(
    *,
    gmail_service: Any,
    raiz: dict[str, Any],
    firmantes: list[str],
    datos_notaria: dict[str, str],
    id_evento_cierre: str,
) -> list[dict[str, Any]]:
    """Correo externo individual a firmantes, una vez completo el retorno del paquete."""
    asunto, cuerpo, cuerpo_html = construir_email_fin_paquete_notarial(
        raiz=raiz,
        datos_notaria=datos_notaria,
    )
    resultados: list[dict[str, Any]] = []
    for email in firmantes:
        try:
            rfc_message_id = construir_rfc_message_id_notificacion(
                f"fin-notaria-{id_evento_cierre}-{email}"
            )
            respuesta = enviar_email_notificacion(
                gmail_service=gmail_service,
                destinatario=email,
                asunto=asunto,
                cuerpo_texto=cuerpo,
                cuerpo_html=cuerpo_html,
                rfc_message_id=rfc_message_id,
            )
            resultados.append(
                {
                    "ok": True,
                    "destinatario": email,
                    "message_id": respuesta["message_id"],
                    "thread_id": respuesta["thread_id"],
                }
            )
        except Exception as exc:
            traceback.print_exc()
            resultados.append(
                {
                    "ok": False,
                    "destinatario": email,
                    "error": str(exc),
                }
            )
    return resultados


def registrar_error_devolucion_notaria(id_documento: str, mensaje: str) -> None:
    try:
        appsheet_action(
            TABLA_DOCUMENTOS,
            "Edit",
            [
                {
                    "ID_DOCUMENTO": id_documento,
                    "ACCION_SOLICITADA": "",
                    "OBSERVACION_ACTUAL": texto(mensaje)[:1000],
                    "FECHA_ULTIMA_ACTUALIZACION": ahora_iso(),
                }
            ],
        )
    except Exception:
        traceback.print_exc()


@app.route("/registrar-devolucion-notaria", methods=["POST"])
def registrar_devolucion_notaria():
    """
    Registra un documento devuelto por notaría.

    - Publica el archivo en Documentos_Prime.
    - Notifica internamente al Responsable de firmas por cada devolución.
    - Cuando todo el paquete está devuelto, notifica una sola vez a los firmantes.
    """
    id_documento = ""
    try:
        validar_configuracion()
        validar_token()
        data = request.get_json(silent=True) or {}

        id_documento = texto(data.get("id_documento"))
        id_documento_raiz = texto(data.get("id_documento_raiz"))
        id_propiedad_prime = texto(data.get("id_propiedad_prime"))
        archivo_devuelto = texto(data.get("archivo_devuelto"))
        usuario = texto(data.get("usuario"))

        if not id_documento:
            return {"error": "Falta id_documento"}, 400

        documento = buscar_documento(id_documento)
        id_raiz_documento = texto(documento.get("ID_DOCUMENTO_RAIZ_NOTARIAL"))
        id_documento_raiz = id_documento_raiz or id_raiz_documento
        if not id_documento_raiz:
            raise ValueError("El documento no tiene ID_DOCUMENTO_RAIZ_NOTARIAL")
        if id_raiz_documento and id_raiz_documento != id_documento_raiz:
            raise ValueError("id_documento_raiz no coincide con el documento")

        raiz = buscar_documento(id_documento_raiz)
        if texto(raiz.get("TIPO_FIRMA")) != "Notarial":
            raise ValueError("El paquete no corresponde a Firma Notarial")

        # Idempotencia antes de validar estados: una devolución ya registrada puede
        # tener ESTADO='Proceso terminado' y el paquete puede estar 'Cerrado'.
        id_prime_existente = texto(documento.get("ID_DOCUMENTO_PRIME_GENERADO"))
        if id_prime_existente:
            completo, _ = paquete_notarial_devuelto_completo(id_documento_raiz)
            return jsonify(
                {
                    "ok": True,
                    "ya_procesado": True,
                    "id_documento": id_documento,
                    "id_documento_raiz": id_documento_raiz,
                    "id_documento_prime": id_prime_existente,
                    "estado_documento": texto(documento.get("ESTADO")),
                    "estado_paquete_notarial": texto(raiz.get("ESTADO_PAQUETE_NOTARIAL")),
                    "paquete_completo": completo,
                }
            )

        if texto(raiz.get("ESTADO_PAQUETE_NOTARIAL")) != ESTADO_PAQUETE_NOTARIAL_EN_NOTARIA:
            raise ValueError(
                "Solo se puede registrar devolución cuando el paquete está En notaría. "
                f"Estado actual: {texto(raiz.get('ESTADO_PAQUETE_NOTARIAL'))!r}"
            )
        if texto(documento.get("ESTADO")) != ESTADO_DOCUMENTO_EN_NOTARIA:
            raise ValueError(
                "Solo se puede registrar devolución de un documento En notaría. "
                f"Estado actual: {texto(documento.get('ESTADO'))!r}"
            )

        id_propiedad_registrada = texto(raiz.get("ID_PROPIEDAD_PRIME_NOTARIA"))
        id_propiedad_prime = id_propiedad_prime or id_propiedad_registrada
        if not id_propiedad_prime:
            raise ValueError("No existe ID_PROPIEDAD_PRIME_NOTARIA en la raíz")
        if id_propiedad_registrada and id_propiedad_prime != id_propiedad_registrada:
            raise ValueError("id_propiedad_prime no coincide con el paquete")

        archivo_devuelto = archivo_devuelto or texto(
            documento.get("ARCHIVO_DEVUELTO_NOTARIA_TEMP")
        )
        if not archivo_devuelto:
            raise ValueError("No se indicó ARCHIVO_DEVUELTO_NOTARIA_TEMP")

        usuario = usuario or texto(documento.get("RECIBIDO_NOTARIA_POR"))
        if not usuario or not _EMAIL_RE.fullmatch(usuario.lower()):
            raise ValueError("El usuario que registra la devolución no es válido")

        fecha = ahora_iso()
        registro_prime = crear_registro_prime_devolucion_notaria(
            documento=documento,
            id_propiedad_prime=id_propiedad_prime,
            archivo_devuelto=archivo_devuelto,
            fecha=fecha,
        )
        id_documento_prime = primer_valor(
            registro_prime,
            "id_documento",
            "ID_DOCUMENTO",
        )
        if not id_documento_prime:
            raise RuntimeError("Documentos_Prime no devolvió id_documento")

        actualizar_documento_tras_devolucion_notaria(
            id_documento=id_documento,
            id_documento_prime=id_documento_prime,
            usuario=usuario,
            fecha=fecha,
        )

        documento_actualizado = buscar_documento(id_documento)
        evento_devolucion = crear_evento_devolucion_notaria(
            documento=documento_actualizado,
            id_documento_prime=id_documento_prime,
            usuario=usuario,
            fecha=fecha,
        )

        advertencias: list[str] = []
        notificaciones_internas: list[dict[str, Any]] = []
        try:
            notificaciones_internas = notificar_responsable_documento_devuelto_notaria(
                documento=documento_actualizado,
                evento=evento_devolucion,
            )
            fallidas = [x for x in notificaciones_internas if not x.get("ok")]
            if fallidas:
                advertencias.append(
                    f"{len(fallidas)} notificación(es) internas quedaron omitidas o con error."
                )
        except Exception as exc_notificacion:
            traceback.print_exc()
            advertencias.append(
                "La devolución quedó registrada, pero falló la notificación al Responsable de firmas: "
                + str(exc_notificacion)
            )

        paquete_completo, integrantes = paquete_notarial_devuelto_completo(
            id_documento_raiz
        )
        notificaciones_firmantes: list[dict[str, Any]] = []
        evento_cierre_id = ""

        if paquete_completo:
            raiz_actual = buscar_documento(id_documento_raiz)
            evento_cierre, creado = crear_evento_cierre_retorno_notaria(
                raiz=raiz_actual,
                usuario=usuario,
                fecha=fecha,
            )
            evento_cierre_id = texto(evento_cierre.get("ID_EVENTO"))

            try:
                cerrar_paquete_notarial_tras_devoluciones(
                    id_documento_raiz=id_documento_raiz,
                    fecha=fecha,
                )
            except Exception as exc_cierre_paquete:
                traceback.print_exc()
                advertencias.append(
                    "Todos los documentos fueron recibidos, pero no fue posible cambiar "
                    "ESTADO_PAQUETE_NOTARIAL a 'Cerrado': " + str(exc_cierre_paquete)
                )

            # Solo el primer cierre dispara correo externo a los firmantes.
            if creado:
                firmantes = normalizar_destinatarios(
                    raiz_actual.get("DESTINATARIOS_FIRMANTES_NOTARIA")
                )
                if firmantes:
                    try:
                        id_notaria = texto(raiz_actual.get("ID_NOTARIA"))
                        if not id_notaria:
                            raise ValueError("La raíz no tiene ID_NOTARIA")
                        datos_notaria = normalizar_datos_notaria(
                            buscar_notaria_por_id(id_notaria)
                        )
                        gmail_service = obtener_gmail_service()
                        notificaciones_firmantes = enviar_notificaciones_fin_paquete_notarial(
                            gmail_service=gmail_service,
                            raiz=raiz_actual,
                            firmantes=firmantes,
                            datos_notaria=datos_notaria,
                            id_evento_cierre=evento_cierre_id,
                        )
                        fallidas = [
                            x for x in notificaciones_firmantes if not x.get("ok")
                        ]
                        if fallidas:
                            advertencias.append(
                                f"{len(fallidas)} notificación(es) finales a firmantes fallaron."
                            )
                    except Exception as exc_firmantes:
                        traceback.print_exc()
                        advertencias.append(
                            "El paquete quedó completo, pero falló la notificación final a firmantes: "
                            + str(exc_firmantes)
                        )
                else:
                    advertencias.append(
                        "El paquete quedó completo, pero no hay DESTINATARIOS_FIRMANTES_NOTARIA."
                    )

        return jsonify(
            {
                "ok": True,
                "ya_procesado": False,
                "id_documento": id_documento,
                "id_documento_raiz": id_documento_raiz,
                "id_documento_prime": id_documento_prime,
                "id_evento_devolucion": texto(evento_devolucion.get("ID_EVENTO")),
                "paquete_completo": paquete_completo,
                "cantidad_documentos_paquete": len(integrantes),
                "id_evento_cierre": evento_cierre_id,
                "estado_documento": ESTADO_DOCUMENTO_PROCESO_TERMINADO,
                "estado_paquete_notarial": (
                    ESTADO_PAQUETE_NOTARIAL_CERRADO
                    if paquete_completo
                    else ESTADO_PAQUETE_NOTARIAL_EN_NOTARIA
                ),
                "notificaciones_internas": notificaciones_internas,
                "notificaciones_firmantes": notificaciones_firmantes,
                "advertencias": advertencias,
            }
        )

    except PermissionError as exc:
        if id_documento:
            registrar_error_devolucion_notaria(id_documento, str(exc))
        return {"error": str(exc)}, 403
    except (ValueError, LookupError) as exc:
        if id_documento:
            registrar_error_devolucion_notaria(id_documento, str(exc))
        return {"error": str(exc)}, 400
    except Exception as exc:
        traceback.print_exc()
        if id_documento:
            registrar_error_devolucion_notaria(id_documento, str(exc))
        return {"error": str(exc)}, 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)
