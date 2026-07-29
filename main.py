from __future__ import annotations

import json
import os
import re
import traceback
import uuid
from datetime import datetime
from typing import Any
from urllib.parse import quote
from zoneinfo import ZoneInfo

import requests
from flask import Flask, jsonify, request
from google.auth.transport.requests import Request as GoogleAuthRequest
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build


app = Flask(__name__)


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

DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]
CHILE_TZ = ZoneInfo("America/Santiago")


# -----------------------------------------------------------------------------
# Utilidades generales
# -----------------------------------------------------------------------------


def ahora_iso() -> str:
    return datetime.now(CHILE_TZ).isoformat(timespec="seconds")


def nuevo_id() -> str:
    return str(uuid.uuid4())


def texto(valor: Any) -> str:
    if valor is None:
        return ""
    return str(valor).strip()


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
    return texto(valor).lower() in {"true", "yes", "si", "sí", "1", "y"}


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


def obtener_drive_service():
    credentials = Credentials(
        token=None,
        refresh_token=GOOGLE_OAUTH_REFRESH_TOKEN,
        token_uri=GOOGLE_OAUTH_TOKEN_URI,
        client_id=GOOGLE_OAUTH_CLIENT_ID,
        client_secret=GOOGLE_OAUTH_CLIENT_SECRET,
        scopes=DRIVE_SCOPES,
    )

    credentials.refresh(GoogleAuthRequest())

    return build(
        "drive",
        "v3",
        credentials=credentials,
        cache_discovery=False,
    )


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


def asegurar_permiso_editor(
    drive_service: Any,
    file_id: str,
    email: str,
) -> str:
    if not email or "@" not in email:
        raise ValueError(
            f"El encargado de orden 1 no tiene un correo válido: {email!r}"
        )

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

        rol_actual = texto(permiso.get("role")).lower()
        if rol_actual in {"owner", "organizer", "fileorganizer", "writer"}:
            return texto(permiso.get("id"))

        permiso_actualizado = (
            drive_service.permissions()
            .update(
                fileId=file_id,
                permissionId=permiso["id"],
                body={"role": "writer"},
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
                "role": "writer",
                "emailAddress": email,
            },
            fields="id",
            sendNotificationEmail=DRIVE_SEND_NOTIFICATION_EMAIL,
            supportsAllDrives=True,
        )
        .execute()
    )

    return texto(permiso_creado.get("id"))


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

    response = requests.post(
        url,
        headers=headers,
        json=payload,
        timeout=90,
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
    selector = (
        f"ORDERBY("
        f"FILTER({TABLA_APROBADORES}, "
        f"AND("
        f"[ID_PLANTILLA] = {literal_appsheet(id_plantilla)}, "
        f"[VIGENTE] = TRUE"
        f")), "
        f"[ORDEN], FALSE"
        f")"
    )

    filas = appsheet_find(TABLA_APROBADORES, selector)

    if not filas:
        raise ValueError(
            "La plantilla no tiene una cadena vigente en "
            f"{TABLA_APROBADORES}"
        )

    filas_ordenadas = sorted(
        filas,
        key=lambda fila: entero(fila.get("ORDEN"), "ORDEN"),
    )

    ordenes = [entero(fila.get("ORDEN"), "ORDEN") for fila in filas_ordenadas]
    if len(ordenes) != len(set(ordenes)):
        raise ValueError(
            "La cadena de aprobación tiene dos o más responsables con el mismo ORDEN"
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
        if id_documento:
            marcar_documento_error(id_documento, str(exc))
        return {"error": str(exc)}, 400

    except Exception as exc:
        traceback.print_exc()
        if id_documento:
            marcar_documento_error(id_documento, str(exc))
        return {"error": str(exc)}, 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)
