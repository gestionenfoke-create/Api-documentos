from __future__ import annotations

import os
import traceback
import uuid
from datetime import datetime
from typing import Any

import google.auth
import requests
from flask import Flask, jsonify, request
from googleapiclient.discovery import build


app = Flask(__name__)


APPSHEET_APP_ID = os.environ.get("APPSHEET_APP_ID")
APPSHEET_ACCESS_KEY = os.environ.get("APPSHEET_ACCESS_KEY")

TABLA_DOCUMENTOS = os.environ.get(
    "TABLA_DOCUMENTOS",
    "Documentos",
)

TABLA_VERSIONES = os.environ.get(
    "TABLA_VERSIONES",
    "Documento_Versiones",
)

WEBHOOK_TOKEN = os.environ.get("WEBHOOK_TOKEN")


def validar_configuracion() -> None:
    faltantes = []

    if not APPSHEET_APP_ID:
        faltantes.append("APPSHEET_APP_ID")

    if not APPSHEET_ACCESS_KEY:
        faltantes.append("APPSHEET_ACCESS_KEY")

    if not WEBHOOK_TOKEN:
        faltantes.append("WEBHOOK_TOKEN")

    if faltantes:
        raise RuntimeError(
            "Faltan variables de entorno: " + ", ".join(faltantes)
        )


def validar_token() -> None:
    token_recibido = request.headers.get("X-Webhook-Token", "")

    if token_recibido != WEBHOOK_TOKEN:
        raise PermissionError("Token de webhook inválido")


def obtener_drive_service():
    credentials, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/drive"]
    )

    return build(
        "drive",
        "v3",
        credentials=credentials,
        cache_discovery=False,
    )


def appsheet_action(
    table_name: str,
    action: str,
    rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    url = (
        f"https://www.appsheet.com/api/v2/apps/"
        f"{APPSHEET_APP_ID}/tables/{table_name}/Action"
    )

    headers = {
        "ApplicationAccessKey": APPSHEET_ACCESS_KEY,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

    payload = {
        "Action": action,
        "Properties": {
            "Locale": "es-CL",
            "Location": "-33.4489,-70.6693",
            "Timezone": "America/Santiago",
        },
        "Rows": rows,
    }

    response = requests.post(
        url,
        headers=headers,
        json=payload,
        timeout=90,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Error AppSheet {table_name}: "
            f"{response.status_code} - {response.text}"
        )

    if not response.text.strip():
        return []

    return response.json()


def copiar_plantilla(
    template_id: str,
    folder_id: str,
    nombre_documento: str,
) -> dict[str, str]:
    drive_service = obtener_drive_service()
    
 # ======= AGREGAR ESTO =======
    about = drive_service.about().get(
        fields="user,storageQuota"
    ).execute()

    print("===== DRIVE INFO =====")
    print(about)
    print("======================")
    # ============================


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


def crear_registro_version(
    id_documento: str,
    google_doc_id: str,
    google_doc_url: str,
    creado_por: str,
) -> str:
    id_version = str(uuid.uuid4())

    fila_version = {
        "ID_VERSION": id_version,
        "ID_DOCUMENTO": id_documento,
        "NUMERO_VERSION": 1,
        "GOOGLE_DOC_ID": google_doc_id,
        "GOOGLE_DOC_URL": google_doc_url,
        "ESTADO_VERSION": "Borrador",
        "FECHA_CREACION": datetime.now().isoformat(),
        "CREADO_POR": creado_por,
    }

    appsheet_action(
        TABLA_VERSIONES,
        "Add",
        [fila_version],
    )

    return id_version


def actualizar_documento(
    id_documento: str,
    google_doc_id: str,
    google_doc_url: str,
    id_version: str,
) -> None:
    fila_documento = {
        "ID_DOCUMENTO": id_documento,
        "VERSION_ACTUAL": 1,
        "GOOGLE_DOC_ID": google_doc_id,
        "GOOGLE_DOC_URL": google_doc_url,
        "ID_VERSION_ACTUAL": id_version,
        "ESTADO": "Borrador",
    }

    appsheet_action(
        TABLA_DOCUMENTOS,
        "Edit",
        [fila_documento],
    )


@app.route("/")
def home():
    return "API Documentos funcionando"


@app.route("/health")
def health():
    return {
        "status": "ok",
        "api": "documentos",
        "server_time": datetime.now().isoformat(),
    }


@app.route("/crear-documento", methods=["POST"])
def crear_documento():
    try:
        validar_configuracion()
        validar_token()

        data = request.get_json(silent=True) or {}

        id_documento = str(
            data.get("id_documento", "")
        ).strip()

        template_id = str(
            data.get("template_id", "")
        ).strip()

        folder_id = str(
            data.get("folder_id", "")
        ).strip()

        titulo = str(
            data.get("titulo", "")
        ).strip()

        creado_por = str(
            data.get("creado_por", "")
        ).strip()

        if not id_documento:
            return {"error": "Falta id_documento"}, 400

        if not template_id:
            return {"error": "Falta template_id"}, 400

        if not folder_id:
            return {"error": "Falta folder_id"}, 400

        if not titulo:
            titulo = f"Documento_{id_documento}"

        nombre_documento = f"{titulo}_V01"

        copia = copiar_plantilla(
            template_id=template_id,
            folder_id=folder_id,
            nombre_documento=nombre_documento,
        )

        id_version = crear_registro_version(
            id_documento=id_documento,
            google_doc_id=copia["id"],
            google_doc_url=copia["url"],
            creado_por=creado_por,
        )

        actualizar_documento(
            id_documento=id_documento,
            google_doc_id=copia["id"],
            google_doc_url=copia["url"],
            id_version=id_version,
        )

        return jsonify(
            {
                "ok": True,
                "id_documento": id_documento,
                "id_version": id_version,
                "numero_version": 1,
                "google_doc_id": copia["id"],
                "google_doc_url": copia["url"],
                "nombre_archivo": copia["name"],
            }
        )

    except PermissionError as exc:
        return {"error": str(exc)}, 403

    except Exception as exc:
        traceback.print_exc()
        return {"error": str(exc)}, 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)
