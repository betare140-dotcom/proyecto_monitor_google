"""Integración de Google OAuth, Drive y Docs para el monitor político.

El módulo no conoce Streamlit salvo en ``StreamlitGoogleOAuth``. La clase
``GoogleReportService`` recibe credenciales OAuth del usuario, copia un Google
Doc que funciona como plantilla visual y genera/actualiza la zona automática
del reporte sin modificar la fotografía previa.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
import json
import re
import secrets
from typing import Any, Iterable, Mapping

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload


GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
]

AUTO_START = "[[MONITOREO_AUTOMATICO_INICIO]]"
AUTO_END = "[[MONITOREO_AUTOMATICO_FIN]]"
PERIOD_TABLE = "[[TABLA_PERIODO]]"
BALANCE_TABLE = "[[TABLA_BALANCE]]"
PAGE_BREAK = "[[SALTO_RESUMEN]]"
SOURCE_MARKER = "[[FUENTE_NOTA]]"

MIME_GOOGLE_DOC = "application/vnd.google-apps.document"
MIME_JSON = "application/json"

CHANNEL_ORDER = [
    "ENTREVISTAS",
    "TELEVISIÓN",
    "RADIO",
    "PRENSA LOCAL",
    "PORTALES DIGITALES",
    "COLUMNAS",
    "REDES SOCIALES",
]


def extraer_id_google(valor: str) -> str:
    """Extrae un ID de Doc/carpeta o acepta directamente un ID de Drive."""
    texto = str(valor or "").strip()
    if not texto:
        raise ValueError("Falta la liga o el identificador de Google Drive.")
    patrones = [
        r"/document/d/([A-Za-z0-9_-]+)",
        r"/folders/([A-Za-z0-9_-]+)",
        r"/file/d/([A-Za-z0-9_-]+)",
        r"[?&]id=([A-Za-z0-9_-]+)",
    ]
    for patron in patrones:
        coincidencia = re.search(patron, texto)
        if coincidencia:
            return coincidencia.group(1)
    if re.fullmatch(r"[A-Za-z0-9_-]{20,}", texto):
        return texto
    raise ValueError("La liga de Google Drive o Google Docs no es válida.")


def _safe_file_name(value: str) -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]+", " ", str(value or ""))
    return re.sub(r"\s+", " ", cleaned).strip()[:180]


def _as_iso_date(value: Any) -> str:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    parsed = datetime.fromisoformat(str(value)[:10])
    return parsed.date().isoformat()


def _note_key(note: Mapping[str, Any]) -> str:
    """Elimina cargas repetidas, pero conserva republicaciones distintas."""
    url = str(note.get("url") or "").strip().lower()
    if url:
        return "url|" + url
    fields = [
        note.get("fecha"),
        note.get("canal"),
        note.get("fuente"),
        note.get("contenido"),
        note.get("origen"),
    ]
    normalized = "|".join(re.sub(r"\s+", " ", str(x or "")).strip().lower() for x in fields)
    return "row|" + normalized


def normalizar_modelo(modelo: Mapping[str, Any]) -> dict[str, Any]:
    actor = str(modelo.get("actor") or "").strip().upper()
    if not actor:
        raise ValueError("Falta el nombre del actor político.")
    inicio = _as_iso_date(modelo["fecha_inicio"])
    fin = _as_iso_date(modelo["fecha_fin"])
    if inicio > fin:
        inicio, fin = fin, inicio

    notes: list[dict[str, Any]] = []
    seen: set[str] = set()
    for raw in modelo.get("notas", []):
        note = {
            "fecha": _as_iso_date(raw.get("fecha")),
            "canal": str(raw.get("canal") or "PORTALES DIGITALES").strip().upper(),
            "sentimiento": (
                "NEGATIVA"
                if str(raw.get("sentimiento") or "").strip().upper() == "NEGATIVA"
                else "POSITIVA"
            ),
            "fuente": re.sub(r"\s+", " ", str(raw.get("fuente") or "Sin fuente")).strip(),
            "contenido": re.sub(r"\s+", " ", str(raw.get("contenido") or "")).strip(),
            "url": str(raw.get("url") or "").strip(),
            "origen": str(raw.get("origen") or "TRADICIONALES").strip().upper(),
        }
        if note["canal"] not in CHANNEL_ORDER:
            note["canal"] = "PORTALES DIGITALES"
        key = _note_key(note)
        if key not in seen:
            seen.add(key)
            notes.append(note)

    notes.sort(key=lambda n: (n["fecha"], CHANNEL_ORDER.index(n["canal"]), n["fuente"].lower()))
    return {
        "version": 1,
        "actor": actor,
        "fecha_inicio": inicio,
        "fecha_fin": fin,
        "temas_informativos": [
            re.sub(r"^\s*\d+[.)-]?\s*", "", str(x)).strip()
            for x in modelo.get("temas_informativos", [])
            if str(x).strip()
        ][:3],
        "temas_negativos": [
            re.sub(r"^\s*\d+[.)-]?\s*", "", str(x)).strip()
            for x in modelo.get("temas_negativos", [])
            if str(x).strip()
        ][:3],
        "notas": notes,
    }


def combinar_modelos(base: Mapping[str, Any], nuevo: Mapping[str, Any]) -> dict[str, Any]:
    """Conserva actor/periodo iniciales y agrega notas que todavía no existen."""
    first = normalizar_modelo(base)
    second = normalizar_modelo({
        **nuevo,
        "actor": first["actor"],
        "fecha_inicio": first["fecha_inicio"],
        "fecha_fin": first["fecha_fin"],
    })
    merged_notes: list[dict[str, Any]] = []
    seen: set[str] = set()
    for note in first["notas"] + second["notas"]:
        key = _note_key(note)
        if key not in seen:
            seen.add(key)
            merged_notes.append(note)

    def merge_topics(key: str) -> list[str]:
        output: list[str] = []
        signatures: set[str] = set()
        for topic in first.get(key, []) + second.get(key, []):
            signature = re.sub(r"\W+", "", topic.lower())
            if signature and signature not in signatures:
                signatures.add(signature)
                output.append(topic)
        return output[:3]

    return normalizar_modelo({
        "actor": first["actor"],
        "fecha_inicio": first["fecha_inicio"],
        "fecha_fin": first["fecha_fin"],
        "temas_informativos": merge_topics("temas_informativos"),
        "temas_negativos": merge_topics("temas_negativos"),
        "notas": merged_notes,
    })


def calcular_metricas(modelo: Mapping[str, Any]) -> dict[str, Any]:
    notes = list(modelo.get("notas", []))
    positive = [n for n in notes if n.get("sentimiento") != "NEGATIVA"]
    negative = [n for n in notes if n.get("sentimiento") == "NEGATIVA"]

    def by_channel(items: Iterable[Mapping[str, Any]]) -> dict[str, int]:
        result = {channel: 0 for channel in CHANNEL_ORDER}
        for item in items:
            channel = str(item.get("canal") or "PORTALES DIGITALES")
            result[channel if channel in result else "PORTALES DIGITALES"] += 1
        return result

    return {
        "positiva": len(positive),
        "negativa": len(negative),
        "total": len(notes),
        "positivos_canal": by_channel(positive),
        "negativos_canal": by_channel(negative),
    }


def periodo_espanol(start_value: Any, end_value: Any) -> str:
    months = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
    }
    start = datetime.fromisoformat(_as_iso_date(start_value)).date()
    end = datetime.fromisoformat(_as_iso_date(end_value)).date()
    if start == end:
        return f"{start.day} de {months[start.month]} de {start.year}"
    if start.year == end.year and start.month == end.month:
        return f"{start.day} al {end.day} de {months[end.month]} de {end.year}"
    if start.year == end.year:
        return (
            f"{start.day} de {months[start.month]} al {end.day} de "
            f"{months[end.month]} de {end.year}"
        )
    return (
        f"{start.day} de {months[start.month]} de {start.year} al "
        f"{end.day} de {months[end.month]} de {end.year}"
    )


class StreamlitGoogleOAuth:
    """OAuth por usuario para aplicaciones Streamlit."""

    SESSION_CREDENTIALS = "google_oauth_credentials"
    SESSION_STATE = "google_oauth_state"

    def __init__(self, streamlit_module: Any, config: Mapping[str, Any]):
        self.st = streamlit_module
        self.client_id = str(config.get("client_id") or "").strip()
        self.client_secret = str(config.get("client_secret") or "").strip()
        self.redirect_uri = str(config.get("redirect_uri") or "").strip()
        if not all([self.client_id, self.client_secret, self.redirect_uri]):
            raise RuntimeError(
                "Faltan client_id, client_secret o redirect_uri en "
                ".streamlit/secrets.toml."
            )

    @property
    def client_config(self) -> dict[str, Any]:
        return {
            "web": {
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": [self.redirect_uri],
            }
        }

    def _flow(self, state: str | None = None) -> Flow:
        flow = Flow.from_client_config(
            self.client_config,
            scopes=GOOGLE_SCOPES,
            state=state,
        )
        flow.redirect_uri = self.redirect_uri
        return flow

    def procesar_retorno(self) -> bool:
        code = self.st.query_params.get("code")
        returned_state = self.st.query_params.get("state")
        if not code:
            return False
        expected_state = self.st.session_state.get(self.SESSION_STATE)
        if not expected_state or returned_state != expected_state:
            self.st.query_params.clear()
            raise RuntimeError("Google devolvió una sesión de autorización no válida.")
        flow = self._flow(expected_state)
        flow.fetch_token(code=code)
        self.st.session_state[self.SESSION_CREDENTIALS] = flow.credentials.to_json()
        self.st.session_state.pop(self.SESSION_STATE, None)
        self.st.query_params.clear()
        return True

    def credentials(self) -> Credentials | None:
        raw = self.st.session_state.get(self.SESSION_CREDENTIALS)
        if not raw:
            return None
        try:
            credentials = Credentials.from_authorized_user_info(json.loads(raw), GOOGLE_SCOPES)
            if credentials.expired and credentials.refresh_token:
                credentials.refresh(Request())
                self.st.session_state[self.SESSION_CREDENTIALS] = credentials.to_json()
            if not credentials.valid:
                return None
            return credentials
        except Exception:
            self.st.session_state.pop(self.SESSION_CREDENTIALS, None)
            return None

    def authorization_url(self) -> str:
        state = secrets.token_urlsafe(32)
        flow = self._flow(state)
        url, generated_state = flow.authorization_url(
            access_type="offline",
            include_granted_scopes="true",
            prompt="consent",
        )
        self.st.session_state[self.SESSION_STATE] = generated_state
        return url

    def logout(self) -> None:
        self.st.session_state.pop(self.SESSION_CREDENTIALS, None)
        self.st.session_state.pop(self.SESSION_STATE, None)


@dataclass
class GoogleReportResult:
    document_id: str
    document_url: str
    title: str
    state_file_id: str | None = None


class GoogleReportService:
    """Crea y actualiza reportes nativos de Google Docs."""

    def __init__(self, credentials: Credentials):
        self.drive = build("drive", "v3", credentials=credentials, cache_discovery=False)
        self.docs = build("docs", "v1", credentials=credentials, cache_discovery=False)

    def copiar_plantilla(
        self,
        template_url: str,
        output_name: str,
        folder_url: str | None = None,
    ) -> GoogleReportResult:
        template_id = extraer_id_google(template_url)
        metadata = self.drive.files().get(
            fileId=template_id,
            fields="id,name,mimeType,parents",
            supportsAllDrives=True,
        ).execute()
        if metadata.get("mimeType") != MIME_GOOGLE_DOC:
            raise ValueError("La plantilla debe ser un documento nativo de Google Docs.")
        parents = []
        if folder_url:
            parents = [extraer_id_google(folder_url)]
        elif metadata.get("parents"):
            parents = list(metadata["parents"][:1])
        body: dict[str, Any] = {"name": _safe_file_name(output_name)}
        if parents:
            body["parents"] = parents
        copied = self.drive.files().copy(
            fileId=template_id,
            body=body,
            supportsAllDrives=True,
            fields="id,name,webViewLink,parents",
        ).execute()
        document_id = copied["id"]
        return GoogleReportResult(
            document_id=document_id,
            document_url=copied.get("webViewLink") or f"https://docs.google.com/document/d/{document_id}/edit",
            title=copied.get("name") or output_name,
        )

    def crear_desde_plantilla(
        self,
        template_url: str,
        modelo: Mapping[str, Any],
        folder_url: str | None = None,
    ) -> GoogleReportResult:
        normalized = normalizar_modelo(modelo)
        period = periodo_espanol(normalized["fecha_inicio"], normalized["fecha_fin"])
        title = f"{normalized['actor'].title()} - {period}"
        result = self.copiar_plantilla(template_url, title, folder_url)
        self.renderizar(result.document_id, normalized)
        state_id = self.guardar_estado(result.document_id, normalized)
        result.state_file_id = state_id
        return result

    def actualizar_con_redes(
        self,
        document_url: str,
        modelo_redes: Mapping[str, Any],
    ) -> GoogleReportResult:
        document_id = extraer_id_google(document_url)
        current = self.cargar_estado(document_id)
        merged = combinar_modelos(current, modelo_redes)
        self.renderizar(document_id, merged)
        state_id = self.guardar_estado(document_id, merged)
        metadata = self.drive.files().get(
            fileId=document_id,
            fields="id,name,webViewLink",
            supportsAllDrives=True,
        ).execute()
        return GoogleReportResult(
            document_id=document_id,
            document_url=metadata.get("webViewLink") or f"https://docs.google.com/document/d/{document_id}/edit",
            title=metadata.get("name") or merged["actor"],
            state_file_id=state_id,
        )

    def _get_document(self, document_id: str) -> dict[str, Any]:
        try:
            return self.docs.documents().get(
                documentId=document_id,
                includeTabsContent=True,
            ).execute()
        except TypeError:
            return self.docs.documents().get(documentId=document_id).execute()

    @staticmethod
    def _tab_and_body(document: Mapping[str, Any]) -> tuple[str | None, dict[str, Any]]:
        tabs = document.get("tabs") or []
        if tabs:
            tab = tabs[0]
            tab_id = (tab.get("tabProperties") or {}).get("tabId")
            document_tab = tab.get("documentTab") or {}
            return tab_id, document_tab.get("body") or {}
        return None, dict(document.get("body") or {})

    @staticmethod
    def _with_tab(location: dict[str, Any], tab_id: str | None) -> dict[str, Any]:
        result = dict(location)
        if tab_id:
            result["tabId"] = tab_id
        return result

    def _batch(self, document_id: str, requests: list[dict[str, Any]]) -> dict[str, Any]:
        if not requests:
            return {}
        return self.docs.documents().batchUpdate(
            documentId=document_id,
            body={"requests": requests},
        ).execute()

    @staticmethod
    def _paragraph_text(paragraph: Mapping[str, Any]) -> str:
        output: list[str] = []
        for element in paragraph.get("elements", []):
            if "textRun" in element:
                output.append(element["textRun"].get("content", ""))
        return "".join(output)

    @classmethod
    def _collect_paragraphs(
        cls,
        content: Iterable[Mapping[str, Any]],
        in_table: bool = False,
    ) -> list[dict[str, Any]]:
        output: list[dict[str, Any]] = []
        for element in content or []:
            if "paragraph" in element:
                output.append({
                    "startIndex": element.get("startIndex"),
                    "endIndex": element.get("endIndex"),
                    "paragraph": element["paragraph"],
                    "text": cls._paragraph_text(element["paragraph"]),
                    "in_table": in_table,
                })
            table = element.get("table")
            if table:
                for row in table.get("tableRows", []):
                    for cell in row.get("tableCells", []):
                        output.extend(cls._collect_paragraphs(cell.get("content", []), True))
            toc = element.get("tableOfContents")
            if toc:
                output.extend(cls._collect_paragraphs(toc.get("content", []), in_table))
        return output

    @staticmethod
    def _collect_tables(content: Iterable[Mapping[str, Any]]) -> list[dict[str, Any]]:
        output: list[dict[str, Any]] = []
        for element in content or []:
            if "table" in element:
                output.append(element)
                for row in element["table"].get("tableRows", []):
                    for cell in row.get("tableCells", []):
                        output.extend(GoogleReportService._collect_tables(cell.get("content", [])))
        return output

    def _find_paragraph(self, document: Mapping[str, Any], exact_text: str) -> dict[str, Any]:
        _, body = self._tab_and_body(document)
        for paragraph in self._collect_paragraphs(body.get("content", [])):
            if paragraph["text"].strip() == exact_text:
                return paragraph
        raise ValueError(f"No se encontró el marcador interno {exact_text}.")

    def _clear_auto_region(self, document_id: str) -> None:
        document = self._get_document(document_id)
        tab_id, _ = self._tab_and_body(document)
        try:
            start = self._find_paragraph(document, AUTO_START)
            end = self._find_paragraph(document, AUTO_END)
        except ValueError:
            return
        self._batch(document_id, [{
            "deleteContentRange": {
                "range": self._with_tab({
                    "startIndex": start["startIndex"],
                    "endIndex": max(start["startIndex"] + 1, end["endIndex"] - 1),
                }, tab_id)
            }
        }])

    @staticmethod
    def _format_date_label(value: str) -> str:
        parsed = datetime.fromisoformat(value).date()
        return parsed.strftime("%d.%m.%y")

    def _build_skeleton(self, model: Mapping[str, Any]) -> str:
        metrics = calcular_metricas(model)
        lines: list[str] = [
            AUTO_START,
            model["actor"],
            PERIOD_TABLE,
            BALANCE_TABLE,
            PAGE_BREAK,
            "RESUMEN",
            f"TOTAL NOTAS INFORMATIVAS: {metrics['positiva']}",
        ]
        for channel in CHANNEL_ORDER:
            lines.append(f"{channel}: {metrics['positivos_canal'][channel]}")
        lines.append(f"TOTAL NOTAS NEGATIVAS: {metrics['negativa']}")
        for channel in CHANNEL_ORDER:
            lines.append(f"{channel}: {metrics['negativos_canal'][channel]}")
        lines.extend(["Temas relevantes:"])
        topics_info = model.get("temas_informativos") or [
            "Sin tema disponible: no fue posible generar un resumen informativo."
        ]
        lines.extend(f"{index}. {topic}" for index, topic in enumerate(topics_info[:3], 1))
        lines.append("Temas negativos:")
        topics_negative = model.get("temas_negativos") or [
            "Sin incidencias negativas: no se registraron temas negativos en el periodo."
        ]
        lines.extend(f"{index}. {topic}" for index, topic in enumerate(topics_negative[:3], 1))
        lines.append("DESGLOSE")

        notes_by_date: dict[str, list[Mapping[str, Any]]] = {}
        for note in model.get("notas", []):
            notes_by_date.setdefault(note["fecha"], []).append(note)
        for day in sorted(notes_by_date):
            day_notes = notes_by_date[day]
            positive = [n for n in day_notes if n["sentimiento"] != "NEGATIVA"]
            negative = [n for n in day_notes if n["sentimiento"] == "NEGATIVA"]
            lines.append(self._format_date_label(day))
            lines.append(f"TOTAL DE IMPACTOS INFORMATIVOS: {len(positive)}")
            self._append_note_groups(lines, positive)
            lines.append(f"TOTAL DE IMPACTOS NEGATIVOS: {len(negative)}")
            self._append_note_groups(lines, negative)
        lines.append(AUTO_END)
        return "\n".join(lines) + "\n"

    @staticmethod
    def _append_note_groups(lines: list[str], notes: list[Mapping[str, Any]]) -> None:
        for channel in CHANNEL_ORDER:
            group = [n for n in notes if n.get("canal") == channel]
            if not group:
                continue
            count = f"({len(group)})" if channel == "REDES SOCIALES" else str(len(group))
            lines.append(f"{channel}: {count}")
            for note in group:
                lines.append(SOURCE_MARKER + str(note.get("fuente") or "Sin fuente"))
                lines.append(str(note.get("contenido") or "Sin contenido disponible"))
                if note.get("url"):
                    lines.append(str(note["url"]))

    def renderizar(self, document_id: str, modelo: Mapping[str, Any]) -> None:
        model = normalizar_modelo(modelo)
        self._clear_auto_region(document_id)
        document = self._get_document(document_id)
        tab_id, body = self._tab_and_body(document)
        content = body.get("content", [])
        if not content:
            raise RuntimeError("El Google Doc no contiene un cuerpo editable.")
        insertion_index = max(1, int(content[-1].get("endIndex", 2)) - 1)
        self._batch(document_id, [{
            "insertText": {
                "location": self._with_tab({"index": insertion_index}, tab_id),
                "text": "\n" + self._build_skeleton(model),
            }
        }])

        period = periodo_espanol(model["fecha_inicio"], model["fecha_fin"])
        self._insert_and_fill_table(
            document_id,
            PERIOD_TABLE,
            [[f"PERIODO DE MEDICIÓN: {period}"]],
            merge_first_row=False,
        )
        metrics = calcular_metricas(model)
        self._insert_and_fill_table(
            document_id,
            BALANCE_TABLE,
            [
                ["PRENSA, TV, RADIO, PORTALES, REDES SOCIALES Y COLUMNAS."],
                ["POSITIVA", "NEGATIVA", "TOTAL DE IMPACTOS"],
                [str(metrics["positiva"]), str(metrics["negativa"]), str(metrics["total"])],
            ],
            merge_first_row=True,
        )
        self._replace_with_page_break(document_id)
        self._apply_report_format(document_id)
        self._batch(document_id, [{
            "replaceAllText": {
                "containsText": {"text": SOURCE_MARKER, "matchCase": True},
                "replaceText": "",
            }
        }])

    def _insert_and_fill_table(
        self,
        document_id: str,
        marker: str,
        values: list[list[str]],
        merge_first_row: bool,
    ) -> None:
        document = self._get_document(document_id)
        tab_id, _ = self._tab_and_body(document)
        paragraph = self._find_paragraph(document, marker)
        start = int(paragraph["startIndex"])
        columns = 3 if merge_first_row else 1
        rows = 3 if merge_first_row else 1
        self._batch(document_id, [
            {
                "deleteContentRange": {
                    "range": self._with_tab({
                        "startIndex": start,
                        "endIndex": int(paragraph["endIndex"]) - 1,
                    }, tab_id)
                }
            },
            {
                "insertTable": {
                    "rows": rows,
                    "columns": columns,
                    "location": self._with_tab({"index": start}, tab_id),
                }
            },
        ])

        document = self._get_document(document_id)
        tab_id, body = self._tab_and_body(document)
        tables = self._collect_tables(body.get("content", []))
        candidates = [
            table for table in tables
            if len((table.get("table") or {}).get("tableRows", [])) == rows
        ]
        if not candidates:
            raise RuntimeError("Google Docs no devolvió la tabla insertada.")
        table_element = min(candidates, key=lambda t: abs(int(t.get("startIndex", 0)) - start))
        table_start = int(table_element["startIndex"])
        if merge_first_row:
            self._batch(document_id, [{
                "mergeTableCells": {
                    "tableRange": {
                        "tableCellLocation": {
                            "tableStartLocation": self._with_tab({"index": table_start}, tab_id),
                            "rowIndex": 0,
                            "columnIndex": 0,
                        },
                        "rowSpan": 1,
                        "columnSpan": 3,
                    }
                }
            }])
            document = self._get_document(document_id)
            tab_id, body = self._tab_and_body(document)
            tables = self._collect_tables(body.get("content", []))
            table_element = min(tables, key=lambda t: abs(int(t.get("startIndex", 0)) - table_start))
            table_start = int(table_element["startIndex"])

        rows_data = table_element["table"].get("tableRows", [])
        insertions: list[tuple[int, str]] = []
        for row_index, row_values in enumerate(values):
            cells = rows_data[row_index].get("tableCells", [])
            for column_index, value in enumerate(row_values):
                if column_index >= len(cells):
                    continue
                cell_content = cells[column_index].get("content", [])
                if not cell_content:
                    continue
                insertions.append((int(cell_content[0]["startIndex"]), value))
        requests = [{
            "insertText": {
                "location": self._with_tab({"index": index}, tab_id),
                "text": value,
            }
        } for index, value in sorted(insertions, reverse=True)]
        self._batch(document_id, requests)
        self._format_table(document_id, table_start, rows, columns, merge_first_row)

    @staticmethod
    def _rgb(hex_color: str) -> dict[str, float]:
        value = hex_color.lstrip("#")
        return {
            "red": int(value[0:2], 16) / 255,
            "green": int(value[2:4], 16) / 255,
            "blue": int(value[4:6], 16) / 255,
        }

    def _table_border(self) -> dict[str, Any]:
        return {
            "color": {"color": {"rgbColor": self._rgb("808080")}},
            "width": {"magnitude": 0.75, "unit": "PT"},
            "dashStyle": "SOLID",
        }

    def _format_table(
        self,
        document_id: str,
        expected_start: int,
        rows: int,
        columns: int,
        merged: bool,
    ) -> None:
        document = self._get_document(document_id)
        tab_id, body = self._tab_and_body(document)
        tables = self._collect_tables(body.get("content", []))
        table_element = min(tables, key=lambda t: abs(int(t.get("startIndex", 0)) - expected_start))
        table_start = int(table_element["startIndex"])
        row_data = table_element["table"].get("tableRows", [])
        requests: list[dict[str, Any]] = []
        border = self._table_border()
        base_style = {
            "contentAlignment": "MIDDLE",
            "borderTop": border,
            "borderBottom": border,
            "borderLeft": border,
            "borderRight": border,
            "paddingTop": {"magnitude": 3, "unit": "PT"},
            "paddingBottom": {"magnitude": 3, "unit": "PT"},
            "paddingLeft": {"magnitude": 4, "unit": "PT"},
            "paddingRight": {"magnitude": 4, "unit": "PT"},
        }
        for row_index, row in enumerate(row_data):
            cells = row.get("tableCells", [])
            span = 1 if merged and row_index == 0 else len(cells)
            background = "E7E6E6" if merged and row_index <= 1 else "FFFFFF"
            requests.append({
                "updateTableCellStyle": {
                    "tableRange": {
                        "tableCellLocation": {
                            "tableStartLocation": self._with_tab({"index": table_start}, tab_id),
                            "rowIndex": row_index,
                            "columnIndex": 0,
                        },
                        "rowSpan": 1,
                        "columnSpan": span,
                    },
                    "tableCellStyle": {
                        **base_style,
                        "backgroundColor": {"color": {"rgbColor": self._rgb(background)}},
                    },
                    "fields": (
                        "backgroundColor,contentAlignment,borderTop,borderBottom,"
                        "borderLeft,borderRight,paddingTop,paddingBottom,paddingLeft,paddingRight"
                    ),
                }
            })
            for cell in cells:
                for paragraph in self._collect_paragraphs(cell.get("content", []), True):
                    start = int(paragraph["startIndex"])
                    end = max(start + 1, int(paragraph["endIndex"]) - 1)
                    if end <= start:
                        continue
                    requests.extend([
                        {
                            "updateTextStyle": {
                                "range": self._with_tab({"startIndex": start, "endIndex": end}, tab_id),
                                "textStyle": {
                                    "weightedFontFamily": {"fontFamily": "Verdana"},
                                    "fontSize": {"magnitude": 12, "unit": "PT"},
                                    "bold": True,
                                    "underline": False,
                                },
                                "fields": "weightedFontFamily,fontSize,bold,underline",
                            }
                        },
                        {
                            "updateParagraphStyle": {
                                "range": self._with_tab({
                                    "startIndex": int(paragraph["startIndex"]),
                                    "endIndex": int(paragraph["endIndex"]),
                                }, tab_id),
                                "paragraphStyle": {"alignment": "CENTER"},
                                "fields": "alignment",
                            }
                        },
                    ])
        if columns == 3:
            requests.append({
                "updateTableColumnProperties": {
                    "tableStartLocation": self._with_tab({"index": table_start}, tab_id),
                    "columnIndices": [],
                    "tableColumnProperties": {
                        "widthType": "FIXED_WIDTH",
                        "width": {"magnitude": 170, "unit": "PT"},
                    },
                    "fields": "widthType,width",
                }
            })
        self._batch(document_id, requests)

    def _replace_with_page_break(self, document_id: str) -> None:
        document = self._get_document(document_id)
        tab_id, _ = self._tab_and_body(document)
        paragraph = self._find_paragraph(document, PAGE_BREAK)
        start = int(paragraph["startIndex"])
        self._batch(document_id, [
            {
                "deleteContentRange": {
                    "range": self._with_tab({
                        "startIndex": start,
                        "endIndex": int(paragraph["endIndex"]) - 1,
                    }, tab_id)
                }
            },
            {
                "insertPageBreak": {
                    "location": self._with_tab({"index": start}, tab_id)
                }
            },
        ])

    def _text_style_request(
        self,
        tab_id: str | None,
        start: int,
        end: int,
        style: Mapping[str, Any],
        fields: str,
    ) -> dict[str, Any] | None:
        if end <= start:
            return None
        return {
            "updateTextStyle": {
                "range": self._with_tab({"startIndex": start, "endIndex": end}, tab_id),
                "textStyle": dict(style),
                "fields": fields,
            }
        }

    def _paragraph_style_request(
        self,
        tab_id: str | None,
        paragraph: Mapping[str, Any],
        style: Mapping[str, Any],
        fields: str,
    ) -> dict[str, Any]:
        return {
            "updateParagraphStyle": {
                "range": self._with_tab({
                    "startIndex": int(paragraph["startIndex"]),
                    "endIndex": int(paragraph["endIndex"]),
                }, tab_id),
                "paragraphStyle": dict(style),
                "fields": fields,
            }
        }

    def _apply_report_format(self, document_id: str) -> None:
        document = self._get_document(document_id)
        tab_id, body = self._tab_and_body(document)
        paragraphs = [
            p for p in self._collect_paragraphs(body.get("content", []))
            if not p["in_table"]
        ]
        start_marker = next(p for p in paragraphs if p["text"].strip() == AUTO_START)
        end_marker = next(p for p in paragraphs if p["text"].strip() == AUTO_END)
        scoped = [
            p for p in paragraphs
            if int(start_marker["startIndex"]) <= int(p["startIndex"]) <= int(end_marker["startIndex"])
        ]
        requests: list[dict[str, Any]] = []
        texts = [p["text"].rstrip("\n") for p in scoped]
        url_indexes = {i for i, text in enumerate(texts) if re.fullmatch(r"https?://\S+", text.strip())}
        negative_topics = False
        in_breakdown = False
        for index, paragraph in enumerate(scoped):
            raw = texts[index]
            text = raw.strip()
            visible_start = int(paragraph["startIndex"])
            visible_end = max(visible_start, int(paragraph["endIndex"]) - 1)
            text_style: dict[str, Any] = {}
            fields: list[str] = []
            paragraph_style: dict[str, Any] = {}
            paragraph_fields: list[str] = []

            base_request = self._text_style_request(
                tab_id,
                visible_start,
                visible_end,
                {
                    "weightedFontFamily": {"fontFamily": "Verdana"},
                    "fontSize": {"magnitude": 12, "unit": "PT"},
                    "foregroundColor": {"color": {"rgbColor": self._rgb("000000")}},
                    "backgroundColor": {"color": {"rgbColor": self._rgb("FFFFFF")}},
                    "bold": False,
                    "italic": False,
                    "underline": False,
                },
                "weightedFontFamily,fontSize,foregroundColor,backgroundColor,bold,italic,underline",
            )
            if base_request:
                requests.append(base_request)

            if text in {AUTO_START, AUTO_END}:
                text_style = {
                    "foregroundColor": {"color": {"rgbColor": self._rgb("FFFFFF")}},
                    "fontSize": {"magnitude": 1, "unit": "PT"},
                    "underline": False,
                }
                fields = ["foregroundColor", "fontSize", "underline"]
                paragraph_style = {
                    "spaceAbove": {"magnitude": 0, "unit": "PT"},
                    "spaceBelow": {"magnitude": 0, "unit": "PT"},
                }
                paragraph_fields = ["spaceAbove", "spaceBelow"]
            elif index == 1:
                text_style = {
                    "bold": True,
                    "foregroundColor": {"color": {"rgbColor": self._rgb("003366")}},
                }
                fields = ["bold", "foregroundColor"]
                paragraph_style = {"alignment": "CENTER"}
                paragraph_fields = ["alignment"]
            elif text in {"RESUMEN", "DESGLOSE"}:
                text_style = {
                    "bold": True,
                    "foregroundColor": {"color": {"rgbColor": self._rgb("8B1A1A")}},
                }
                fields = ["bold", "foregroundColor"]
                paragraph_style = {
                    "alignment": "CENTER",
                    "spaceAbove": {"magnitude": 8, "unit": "PT"},
                    "spaceBelow": {"magnitude": 4, "unit": "PT"},
                }
                paragraph_fields = ["alignment", "spaceAbove", "spaceBelow"]
                if text == "DESGLOSE":
                    negative_topics = False
                    in_breakdown = True
            elif text.startswith("TOTAL NOTAS INFORMATIVAS:"):
                text_style = {
                    "bold": True,
                    "backgroundColor": {"color": {"rgbColor": self._rgb("00FF00")}},
                }
                fields = ["bold", "backgroundColor"]
            elif text.startswith("TOTAL NOTAS NEGATIVAS:"):
                text_style = {
                    "bold": True,
                    "foregroundColor": {"color": {"rgbColor": self._rgb("FFFFFF")}},
                    "backgroundColor": {"color": {"rgbColor": self._rgb("FF0000")}},
                }
                fields = ["bold", "foregroundColor", "backgroundColor"]
            elif text == "Temas relevantes:":
                text_style = {"bold": True}
                fields = ["bold"]
                negative_topics = False
            elif text == "Temas negativos:":
                text_style = {
                    "bold": True,
                    "foregroundColor": {"color": {"rgbColor": self._rgb("8B1A1A")}},
                }
                fields = ["bold", "foregroundColor"]
                negative_topics = True
            elif negative_topics and re.match(r"^\d+\.\s", text):
                text_style = {
                    "bold": True,
                    "foregroundColor": {"color": {"rgbColor": self._rgb("8B1A1A")}},
                }
                fields = ["bold", "foregroundColor"]
            elif re.fullmatch(r"\d{2}\.\d{2}\.\d{2}", text):
                text_style = {
                    "bold": True,
                    "underline": False,
                    "backgroundColor": {"color": {"rgbColor": self._rgb("00FFFF")}},
                }
                fields = ["bold", "underline", "backgroundColor"]
                paragraph_style = {
                    "spaceAbove": {"magnitude": 10, "unit": "PT"},
                    "spaceBelow": {"magnitude": 2, "unit": "PT"},
                }
                paragraph_fields = ["spaceAbove", "spaceBelow"]
            elif text.startswith("TOTAL DE IMPACTOS INFORMATIVOS:"):
                text_style = {
                    "bold": True,
                    "backgroundColor": {"color": {"rgbColor": self._rgb("FFFF00")}},
                }
                fields = ["bold", "backgroundColor"]
            elif text.startswith("TOTAL DE IMPACTOS NEGATIVOS:"):
                text_style = {
                    "bold": True,
                    "foregroundColor": {"color": {"rgbColor": self._rgb("FFFFFF")}},
                    "backgroundColor": {"color": {"rgbColor": self._rgb("FF0000")}},
                }
                fields = ["bold", "foregroundColor", "backgroundColor"]
            elif re.fullmatch(
                r"(?:ENTREVISTAS|TELEVISIÓN|RADIO|PRENSA LOCAL|PORTALES DIGITALES|COLUMNAS|REDES SOCIALES):\s*\(?\d+\)?",
                text,
            ):
                text_style = {
                    "bold": True,
                }
                fields = ["bold"]
                if in_breakdown:
                    text_style["backgroundColor"] = {
                        "color": {"rgbColor": self._rgb("00FFFF" if text.startswith("REDES") else "FFFF00")}
                    }
                    fields.append("backgroundColor")
            elif text.startswith(SOURCE_MARKER):
                text_style = {"bold": True}
                fields = ["bold"]
                paragraph_style = {
                    "spaceAbove": {"magnitude": 4, "unit": "PT"},
                    "spaceBelow": {"magnitude": 1, "unit": "PT"},
                }
                paragraph_fields = ["spaceAbove", "spaceBelow"]
            elif index in url_indexes:
                text_style = {
                    "foregroundColor": {"color": {"rgbColor": self._rgb("0066CC")}},
                    "underline": True,
                    "link": {"url": text},
                }
                fields = ["foregroundColor", "underline", "link"]
                paragraph_style = {"spaceBelow": {"magnitude": 6, "unit": "PT"}}
                paragraph_fields = ["spaceBelow"]

            request = self._text_style_request(
                tab_id, visible_start, visible_end, text_style, ",".join(fields)
            ) if fields else None
            if request:
                requests.append(request)
            if paragraph_fields:
                requests.append(self._paragraph_style_request(
                    tab_id, paragraph, paragraph_style, ",".join(paragraph_fields)
                ))

        self._batch(document_id, requests)

    def _state_query(self, document_id: str) -> str:
        safe = document_id.replace("'", "\\'")
        return (
            "trashed = false and "
            f"appProperties has {{ key='monitoring_report_id' and value='{safe}' }}"
        )

    def _document_parent(self, document_id: str) -> str | None:
        metadata = self.drive.files().get(
            fileId=document_id,
            fields="parents",
            supportsAllDrives=True,
        ).execute()
        parents = metadata.get("parents") or []
        return parents[0] if parents else None

    def guardar_estado(self, document_id: str, modelo: Mapping[str, Any]) -> str:
        payload = json.dumps(normalizar_modelo(modelo), ensure_ascii=False, indent=2).encode("utf-8")
        listed = self.drive.files().list(
            q=self._state_query(document_id),
            spaces="drive",
            fields="files(id,name)",
            pageSize=10,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        media = MediaIoBaseUpload(BytesIO(payload), mimetype=MIME_JSON, resumable=False)
        files = listed.get("files") or []
        if files:
            file_id = files[0]["id"]
            self.drive.files().update(
                fileId=file_id,
                media_body=media,
                supportsAllDrives=True,
            ).execute()
            return file_id
        parent = self._document_parent(document_id)
        body: dict[str, Any] = {
            "name": f"_estado_monitoreo_{document_id}.json",
            "mimeType": MIME_JSON,
            "appProperties": {"monitoring_report_id": document_id},
        }
        if parent:
            body["parents"] = [parent]
        created = self.drive.files().create(
            body=body,
            media_body=media,
            fields="id",
            supportsAllDrives=True,
        ).execute()
        return created["id"]

    def cargar_estado(self, document_id_or_url: str) -> dict[str, Any]:
        document_id = extraer_id_google(document_id_or_url)
        listed = self.drive.files().list(
            q=self._state_query(document_id),
            spaces="drive",
            fields="files(id,name,modifiedTime)",
            orderBy="modifiedTime desc",
            pageSize=10,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files = listed.get("files") or []
        if not files:
            raise ValueError(
                "Este documento no tiene el archivo de control del monitoreo. "
                "Debe generarse primero desde la opción Primera extracción."
            )
        buffer = BytesIO()
        downloader = MediaIoBaseDownload(
            buffer,
            self.drive.files().get_media(fileId=files[0]["id"], supportsAllDrives=True),
        )
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buffer.seek(0)
        return normalizar_modelo(json.loads(buffer.read().decode("utf-8")))
