from datetime import datetime
from copy import deepcopy
import io
import json
import os
import re
import time
import unicodedata
import zipfile
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml import OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor
from google import genai
import pandas as pd
import streamlit as st

from google_workspace_report import (
    GoogleReportService,
    StreamlitGoogleOAuth,
    extraer_id_google,
)

MESES_ES = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}

st.set_page_config(
    page_title="Monitoreo Político Oficial", page_icon="📊", layout="centered"
)

st.title("📊 Generador de Monitoreo de Actores Políticos")
st.write(
    "Sistema universal de monitoreo político para cualquier perfil o cargo"
    " público. Permite procesar archivos individuales o fusionar múltiples"
    " archivos (Radio/TV, Portales Web y Redes) en un solo reporte oficial"
    " consolidado."
)

def obtener_api_key_gemini():
  """Lee la clave sin incrustarla en el repositorio."""
  api_key = os.getenv("GEMINI_API_KEY", "").strip()
  if api_key:
    return api_key
  try:
    return str(st.secrets["GEMINI_API_KEY"]).strip()
  except Exception:
    return ""


GEMINI_API_KEY = obtener_api_key_gemini()
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite").strip()
UMBRAL_CONFIANZA_IA = 0.60
CLIENTE_GEMINI = (
    genai.Client(
        api_key=GEMINI_API_KEY,
        http_options={"timeout": 60000},
    )
    if GEMINI_API_KEY
    else None
)
if not GEMINI_API_KEY:
  st.warning(
      "Falta configurar GEMINI_API_KEY. Los reportes tradicionales seguirán"
      " funcionando con el sentimiento incluido en sus archivos, pero necesitas"
      " la clave para analizar redes sociales."
  )

# ==============================================================================
# BASE DE CONOCIMIENTO Y CRITERIOS UNIVERSALES DE CLASIFICACIÓN POLÍTICA
# ==============================================================================
SYSTEM_PROMPT_UNIVERSAL = """
Actúa como analista senior de reputación y comunicación política. Evalúa cada
publicación DESDE LA PERSPECTIVA DEL ACTOR POLÍTICO OBJETIVO: imagina que el
propio actor lee la nota y determina si la mención beneficia/informa sobre su
imagen o si puede perjudicarla, cuestionarla o abrir una crisis.

El reporte utiliza dos grupos:
- POSITIVA: también incluye menciones neutrales o meramente informativas.
- NEGATIVA: incluye cualquier afectación reputacional directa o atribuible.

PROCESO OBLIGATORIO PARA CADA PUBLICACIÓN:
1. Determina si realmente habla del actor objetivo. Si es un homónimo, otra
   persona, una coincidencia de nombre o el actor no aparece, marca
   relevante=false. No inventes una relación.
2. En resúmenes con varias noticias, aísla únicamente el fragmento donde aparece
   el actor. No traslades al actor hechos negativos de otras notas.
3. Marca NEGATIVA cuando el texto atribuya, denuncie, sugiera o reproduzca una
   crítica que razonablemente pueda perjudicar al actor, aunque sea una cita,
   acusación no comprobada, pregunta retórica, opinión, sátira o declaración de
   un adversario. No evalúes si la acusación es verdadera; evalúa su impacto.
4. Son NEGATIVAS, entre otras: señalamientos de corrupción, ilegalidad, abuso,
   opacidad, nepotismo, desvío o uso indebido de recursos; promoción o campaña
   anticipada; bardas o lonas cuestionadas; investigaciones, sanciones o
   denuncias; críticas a capacidad, legitimidad o resultados; reclamos por
   fallas del área bajo su responsabilidad; favoritismo, piso disparejo,
   ventajas injustas, conflictos internos y hashtags ofensivos dirigidos al
   actor.
5. Si una publicación mezcla logros con una crítica directa relevante, la
   crítica prevalece y se clasifica NEGATIVA.
6. Marca POSITIVA cuando la mención sea favorable o informativa sin reproche:
   agenda, declaraciones, eventos, obras, apoyos, convenios, resultados,
   reconocimientos, respaldos, defensa frente a ataques o especulación política
   descriptiva sin cuestionamiento.

EJEMPLOS DE CALIBRACIÓN:
- "Laura Artemisa: el arte de violar la ley" -> NEGATIVA.
- "Acusan promoción anticipada y cuestionan el origen de recursos para bardas"
  -> NEGATIVA.
- "En Morena piden piso parejo y señalan una ventaja injusta para Artemisa"
  -> NEGATIVA.
- "Laura Artemisa entregó recursos para obra comunitaria" -> POSITIVA.
- Una nota sobre una soprano llamada Laura Artemisa, distinta de la funcionaria
  evaluada -> relevante=false.

Devuelve una decisión por cada id recibido. La explicación debe ser breve,
concreta y referirse al impacto sobre el actor.
"""


SCHEMA_CLASIFICACION = {
    "type": "object",
    "properties": {
        "resultados": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "id": {"type": "integer"},
                    "relevante": {"type": "boolean"},
                    "sentimiento": {
                        "type": "string",
                        "enum": ["POSITIVA", "NEGATIVA"],
                    },
                    "confianza": {
                        "type": "number",
                        "minimum": 0,
                        "maximum": 1,
                    },
                    "motivo": {"type": "string"},
                },
                "required": [
                    "id",
                    "relevante",
                    "sentimiento",
                    "confianza",
                    "motivo",
                ],
            },
        }
    },
    "required": ["resultados"],
}


def quitar_acentos(texto):
  if texto is None or pd.isna(texto):
    return ""
  return "".join(
      c
      for c in unicodedata.normalize("NFD", str(texto))
      if unicodedata.category(c) != "Mn"
  ).lower()


def normalizar_cadena(texto):
  t = quitar_acentos(texto)
  return re.sub(r"[^a-z0-9]", "", t)


def normalizar_nombre_candidato(nombre):
  t = quitar_acentos(str(nombre).strip())
  t_clean = re.sub(r"[^a-z0-9]", "", t)

  alias_map = {
      "pepechedraui": "JOSÉ CHEDRAUI BUDIB",
      "josechedraui": "JOSÉ CHEDRAUI BUDIB",
      "rmvb": "RAFAEL MORENO VALLE BUITRÓN",
      "rafamorenovalle": "RAFAEL MORENO VALLE BUITRÓN",
      "rafaelmorenovalle": "RAFAEL MORENO VALLE BUITRÓN",
      "carolinabeau": "CAROLINA BEAUREGARD MARTÍNEZ",
      "carolinabeauregard": "CAROLINA BEAUREGARD MARTÍNEZ",
      "genovevahuerta": "GENOVEVA HUERTA VILLEGAS",
      "gabrielasanchez": "GABRIELA SÁNCHEZ SAAVEDRA",
      "gabylabonita": "GABRIELA SÁNCHEZ SAAVEDRA",
      "bonitasanchez": "GABRIELA SÁNCHEZ SAAVEDRA",
      "lauraartemisa": "LAURA ARTEMISA GARCÍA CHÁVEZ",
      "lupitacuautle": "GUADALUPE CUAUTLE TORRES",
      "guadalupecuautle": "GUADALUPE CUAUTLE TORRES",
      "tonantzinfernandez": "TONANTZIN FERNÁNDEZ DÍAZ",
      "lizsanchez": "LIZETH SÁNCHEZ GARCÍA",
      "lizethsanchez": "LIZETH SÁNCHEZ GARCÍA",
      "nestorcamarillo": "NESTOR CAMARILLO MEDINA",
      "celinapena": "CELINA PEÑA GUZMÁN",
      "rodrigoabdala": "RODRIGO ABDALA DARTIGUES",
      "delfinapozos": "DELFINA POZOS VERGARA",
      "xitlalicceja": "XITLALIC CEJA GARCÍA",
      "blancaalcala": "BLANCA ALCALÁ RUIZ",
  }
  for alias, canon in alias_map.items():
    if alias in t_clean or t_clean in alias:
      return canon
  return str(nombre).strip().upper()


PATRONES_INSTITUCIONALES_GENERICOS = [
    r"ayuntamiento",
    r"gobierno",
    r"gobpue",
    r"gobiernoded",
    r"ayto",
    r"snandresoficial",
    r"cholulaoficial",
    r"secretaria",
    r"dependencia",
    r"organismo",
    r"sindicatura",
    r"presidencia",
    r"comunicacionsocial",
    r"seguridadciudadana",
    r"seguridadpublica",
    r"proteccioncivil",
    r"policiamunicipal",
    r"policiastat",
    r"serviciospublicos",
    r"servicios_pub",
    r"obraspublicas",
    r"desarrollourbano",
    r"desarrolloeconomico",
    r"dif",
    r"sistemadif",
    r"difestatal",
    r"difmunicipal",
    r"institutodelamujer",
    r"institutodelajuventud",
    r"organismodeagua",
    r"serviciodelimpia",
    r"organismolimpia",
    r"derechoshumanos",
    r"casadecultura",
    r"bienestarsanandres",
    r"bienestarcholula",
    r"ssppc",
    r"ssppcsnandres",
    r"deportegobpue",
]

ABREV_APELLIDOS = {
    "fdz": "fernandez",
    "hdez": "hernandez",
    "mtz": "martinez",
    "glez": "gonzalez",
    "gcia": "garcia",
    "lpz": "lopez",
    "prz": "perez",
    "sdo": "saavedra",
    "chvz": "chavez",
}


def es_cuenta_del_actor_universal(autor, handle, actor_target):
  aut = normalizar_cadena(autor)
  hnd = normalizar_cadena(handle)
  act = normalizar_cadena(actor_target)

  if not act:
    return False

  if act in aut or act in hnd or aut in act:
    return True

  tokens_actor = [
      t for t in re.findall(r"\w+", quitar_acentos(actor_target)) if len(t) > 2
  ]
  for abrev, apellido_completo in ABREV_APELLIDOS.items():
    if apellido_completo in tokens_actor:
      tokens_actor.append(abrev)

  nombres_distintivos = [
      t
      for t in tokens_actor
      if len(t) >= 5 and t not in ["perez", "lopez", "garcia", "martinez"]
  ]
  for n in nombres_distintivos:
    if n in aut or n in hnd:
      otros_tokens = [t for t in tokens_actor if t != n]
      if any(ot in aut or ot in hnd for ot in otros_tokens):
        return True
      if hnd == n or aut == n:
        return True

  coincidencias_aut = sum(1 for t in tokens_actor if t in aut)
  coincidencias_hnd = sum(1 for t in tokens_actor if t in hnd)
  if coincidencias_aut >= 2 or coincidencias_hnd >= 2:
    return True

  return False


def es_institucional_universal(autor, handle, texto_post=""):
  texto = str(autor) + " " + str(handle)
  texto_norm = normalizar_cadena(texto)
  for p in PATRONES_INSTITUCIONALES_GENERICOS:
    if re.search(p, texto_norm):
      return True

  t_clean = quitar_acentos(texto_post).lower()
  frases_institucionales = [
      "nuestra presidenta municipal",
      "nuestra presidenta honoraria",
      "nuestro presidente municipal",
      "mi mas sincero agradecimiento",
      "nuestro compromiso de seguir",
      "desde el gobierno municipal",
      "agradecemos a nuestra presidenta",
      "agradecemos a nuestro presidente",
      "los invitamos a participar en nuestro",
  ]
  if any(f in t_clean for f in frases_institucionales):
    return True

  return False


def limpiar_dataframe_redes_automatico(df_raw, actor_nombre_target):
  def es_descartable(row):
    autor = obtener_campo(
        row,
        [
            "Autor",
            "Author name",
            "Author",
            "User",
            "Username",
            "Handle",
            "Fuente",
            "Nombre del Medio",
            "Canal",
        ],
    )
    handle = obtener_campo(
        row,
        [
            "Author handle (@username)",
            "Handle",
            "Username",
            "Screen Name",
            "Account",
            "Perfil",
        ],
    )
    detalle = obtener_campo(
        row,
        [
            "Contenido",
            "Detail",
            "Titulo",
            "Título",
            "Summary",
            "Síntesis",
            "Nota",
        ],
    )
    link = obtener_campo(
        row, ["Link URL Medio", "URL", "Enlace", "Link", "Link de Nota"]
    )

    # 1. Comprobar cuenta del actor político
    if es_cuenta_del_actor_universal(autor, handle, actor_nombre_target):
      return True

    # 2. Comprobar cuenta institucional / DIF / Dependencia
    if es_institucional_universal(autor, handle, detalle):
      return True

    # 3. Comprobar si el link del post apunta al perfil propio del actor
    if link and actor_nombre_target:
      link_norm = normalizar_cadena(link)
      tokens_actor = [
          t
          for t in re.findall(r"\w+", quitar_acentos(actor_nombre_target))
          if len(t) >= 5
      ]
      for n in tokens_actor:
        if (
            f"x.com/{n}" in link.lower()
            or f"twitter.com/{n}" in link.lower()
            or f"facebook.com/{n}" in link.lower()
            or f"{n}fdz" in link_norm
            or f"{n}t" in link_norm
        ):
          return True

    det_lower = quitar_acentos(detalle)
    if any(
        p in det_lower
        for p in [
            "mis ahijados",
            "con toda la actitud #graciasdios",
            "primeracomunion",
            "en familia festejando",
        ]
    ):
      return True

    return False

  mask_descarte = df_raw.apply(es_descartable, axis=1)
  df_limpio = df_raw[~mask_descarte].copy()
  total_descartadas = mask_descarte.sum()

  return df_limpio, total_descartadas


def cargar_archivo_seguro(file):
  try:
    file_bytes = file.read()
    file.seek(0)
  except Exception:
    file_bytes = file

  try:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
  except Exception:
    pass

  try:
    df_csv = pd.read_csv(io.BytesIO(file_bytes), on_bad_lines="skip")
    return {"Reporte": df_csv}
  except Exception:
    pass

  try:
    df_csv = pd.read_csv(
        io.BytesIO(file_bytes), encoding="latin1", on_bad_lines="skip"
    )
    return {"Reporte": df_csv}
  except Exception:
    pass

  return {}


def obtener_campo(row, lista_cols):
  for c in lista_cols:
    if c in row.index:
      val = row[c]
      if val is not None and not pd.isna(val):
        v = str(val).strip()
        if v and v.lower() not in ["nan", "none", "null"]:
          return v
    for col_existente in row.index:
      col_str = str(col_existente).strip()
      if col_str.lower() == c.lower() or quitar_acentos(
          col_str
      ) == quitar_acentos(c):
        val = row[col_existente]
        if val is not None and not pd.isna(val):
          v = str(val).strip()
          if v and v.lower() not in ["nan", "none", "null"]:
            return v
  return ""


def obtener_columna_serie(df_data, lista_posibles_cols):
  for c in lista_posibles_cols:
    if c in df_data.columns:
      return df_data[c]
    for col_existente in df_data.columns:
      col_str = str(col_existente).strip()
      if col_str.lower() == c.lower() or quitar_acentos(
          col_str
      ) == quitar_acentos(c):
        return df_data[col_existente]
  return pd.Series([""] * len(df_data), index=df_data.index)


def parsear_fecha_perfecta(val):
  if val is None or pd.isna(val) or str(val).strip() in ["nan", "None", ""]:
    return pd.NaT

  if isinstance(val, (datetime, pd.Timestamp)):
    return pd.Timestamp(val).to_pydatetime()

  s = str(val).strip()
  if "," in s:
    s = s.split(",")[0].strip()
  s_date = re.split(r"[ T]", s, maxsplit=1)[0].strip()

  match_iso = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", s_date)
  if match_iso:
    try:
      y, m, d = map(int, match_iso.groups())
      return datetime(y, m, d)
    except ValueError:
      return pd.NaT

  match_latam = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", s_date)
  if match_latam:
    try:
      d, m, y = map(int, match_latam.groups())
      if y < 100:
        y += 2000
      return datetime(y, m, d)
    except ValueError:
      return pd.NaT

  try:
    return pd.to_datetime(s, dayfirst=True, errors="coerce")
  except Exception:
    return pd.NaT


def limpiar_texto(texto):
  if not isinstance(texto, str) or texto == "nan":
    return ""
  texto_sin_urls = re.sub(r"https?://\S+", "", texto)
  texto_sin_emojis = re.sub(r":[a-zA-Z0-9_\-|]+:", "", texto_sin_urls)
  texto_sin_emojis = re.sub(r"[\U00010000-\U0010ffff]", "", texto_sin_emojis)
  lineas = [
      re.sub(r"[ \t]+", " ", line).strip()
      for line in texto_sin_emojis.split("\n")
      if line.strip()
  ]
  return "\n".join(lineas)


def estandarizar_categoria_medio(m_type_raw):
  if m_type_raw is None or pd.isna(m_type_raw):
    return "PORTALES DIGITALES"
  t_norm = quitar_acentos(str(m_type_raw).strip())
  if "tele" in t_norm or "tv" in t_norm:
    return "TELEVISIÓN"
  elif "rad" in t_norm or "fm" in t_norm or "am" in t_norm:
    return "RADIO"
  elif any(k in t_norm for k in ["prensa", "periodico", "diario", "impreso"]):
    # Se evalúa antes que LOCAL para que PRENSA LOCAL no se convierta en portal.
    return "PRENSA LOCAL"
  elif any(k in t_norm for k in ["columna", "opinion"]):
    return "COLUMNAS"
  elif (
      re.search(r"\b(redes?|social(?:es)?|twitter|facebook|instagram|tiktok|youtube)\b", t_norm)
      or t_norm == "x"
  ):
    return "REDES SOCIALES"
  elif any(
      k in t_norm
      for k in [
          "portal", "web", "online", "internet", "digital", "comun",
          "local", "locales",
      ]
  ):
    return "PORTALES DIGITALES"
  return "PORTALES DIGITALES"


def reparar_desfase_columnas_excel(df):
  if df is None or df.empty:
    return df

  cols = [str(c).strip() for c in df.columns]
  es_desfasado = False
  if "Hora" in cols:
    sample_hora = df["Hora"].dropna().astype(str).head(10).tolist()
    if any(
        re.match(
            r"^(Puebla|M[eé]xico|Tlaxcala|Veracruz|CDMX|Hidalgo|Nacional|Internacional)$",
            v.strip(),
            re.I,
        )
        for v in sample_hora
    ):
      es_desfasado = True

  if "Alcance" in cols and not es_desfasado:
    sample_alcance = df["Alcance"].dropna().astype(str).head(10).tolist()
    if any(v.startswith("http") for v in sample_alcance):
      es_desfasado = True

  if es_desfasado:
    columnas_reales_ordenadas = [
        "ID Nota",
        "Menu",
        "Titulo",
        "Autor",
        "Fecha",
        "Estado",
        "Pais",
        "Nombre del Medio",
        "Tipo de Medio",
        "Tipo de Nota",
        "Sentimiento",
        "Costo",
        "Alcance",
        "Link URL Medio",
        "Link de Nota",
    ]

    num_cols = len(df.columns)
    df_reparado = pd.DataFrame()

    for idx, col_name in enumerate(columnas_reales_ordenadas):
      if idx < num_cols:
        df_reparado[col_name] = df.iloc[:, idx].values
      else:
        df_reparado[col_name] = ""

    df_reparado["Hora"] = ""
    return df_reparado

  return df


def construir_registro_para_ia(row, local_id):
  """Combina título y contenido; no entrega a la IA solo la primera columna."""
  titulo = obtener_campo(
      row,
      ["Titulo", "Título", "Title", "Encabezado", "Tema"],
  )
  contenido = obtener_campo(
      row,
      ["Contenido", "Detail", "Summary", "Síntesis", "Sintesis", "Nota"],
  )
  autor = obtener_campo(
      row,
      ["Autor", "Author name", "Author", "Fuente", "Nombre del Medio"],
  )
  handle = obtener_campo(
      row,
      ["Author handle (@username)", "Handle", "Username", "Screen Name"],
  )
  medio = obtener_campo(
      row,
      ["Nombre del Medio", "Media name", "Fuente", "Medio", "Canal"],
  )

  partes = []
  for valor in [titulo, contenido]:
    valor_limpio = limpiar_texto_para_resumen(str(valor))
    if valor_limpio and valor_limpio not in partes:
      partes.append(valor_limpio)

  texto = "\n".join(partes).strip()
  if len(texto) > 4500:
    texto = texto[:4500]

  return {
      "id": local_id,
      "autor": autor[:250],
      "handle": handle[:150],
      "medio": medio[:250],
      "texto": texto,
  }


def normalizar_resultado_ia(item, ids_esperados):
  if not isinstance(item, dict):
    return None
  try:
    item_id = int(item.get("id"))
  except (TypeError, ValueError):
    return None
  if item_id not in ids_esperados:
    return None

  sentimiento = str(item.get("sentimiento", "")).strip().upper()
  if sentimiento not in {"POSITIVA", "NEGATIVA"}:
    return None

  try:
    confianza = float(item.get("confianza", 0.5))
  except (TypeError, ValueError):
    confianza = 0.5

  return {
      "sentimiento": sentimiento,
      "relevante": bool(item.get("relevante", True)),
      "confianza": min(1.0, max(0.0, confianza)),
      "motivo": str(item.get("motivo", "Decisión de la IA."))[:350],
      "origen": "IA",
  }


def clasificar_lote_con_ia(lista_notas, actor_nombre, max_intentos=1):
  if CLIENTE_GEMINI is None:
    raise RuntimeError(
        "No se encontró GEMINI_API_KEY. Configúrala en las variables del entorno"
        " o en .streamlit/secrets.toml antes de generar el reporte."
    )

  prompt = f"""
{SYSTEM_PROMPT_UNIVERSAL}

ACTOR POLÍTICO OBJETIVO: "{actor_nombre}"

PUBLICACIONES A EVALUAR:
{json.dumps(lista_notas, ensure_ascii=False)}
"""
  ids_esperados = {int(item["id"]) for item in lista_notas}
  ultimo_error = None

  for intento in range(max_intentos):
    try:
      interaction = CLIENTE_GEMINI.interactions.create(
          model=GEMINI_MODEL,
          input=prompt,
          response_format={
              "type": "text",
              "mime_type": "application/json",
              "schema": SCHEMA_CLASIFICACION,
          },
      )
      raw_txt = (interaction.output_text or "").strip()
      datos = json.loads(raw_txt)
      items = datos.get("resultados", [])
      if not isinstance(items, list):
        raise ValueError("La respuesta no contiene una lista de resultados.")

      res_map = {}
      for item in items:
        resultado = normalizar_resultado_ia(item, ids_esperados)
        if resultado is not None:
          res_map[int(item["id"])] = resultado

      if not res_map:
        raise ValueError("Gemini no devolvió decisiones válidas.")
      return res_map
    except Exception as exc:
      ultimo_error = exc
      if intento + 1 < max_intentos:
        time.sleep(1.2 * (intento + 1))

  raise RuntimeError(f"Gemini no respondió correctamente: {ultimo_error}")


def es_error_no_recuperable_gemini(exc):
  """Distingue cuota/credenciales de una demora que sí permite dividir lote."""
  mensaje = quitar_acentos(str(exc))
  indicadores = [
      "429",
      "quota",
      "api key",
      "api_key",
      "permission denied",
      "unauthenticated",
      "not found",
      "does not exist",
  ]
  return any(indicador in mensaje for indicador in indicadores)


def clasificar_lote_adaptativo(lista_notas, actor_nombre):
  """Divide lotes lentos sin repetir ids ni convertirlos en falsos positivos."""
  try:
    return clasificar_lote_con_ia(lista_notas, actor_nombre)
  except Exception as exc:
    if es_error_no_recuperable_gemini(exc):
      raise
    if len(lista_notas) <= 1:
      registro = lista_notas[0]
      return {
          int(registro["id"]): clasificar_respaldo_local(
              registro.get("texto", ""), actor_nombre
          )
      }

    mitad = max(1, len(lista_notas) // 2)
    resultado = {}
    resultado.update(
        clasificar_lote_adaptativo(lista_notas[:mitad], actor_nombre)
    )
    resultado.update(
        clasificar_lote_adaptativo(lista_notas[mitad:], actor_nombre)
    )
    return resultado


def clasificar_respaldo_local(texto, actor_nombre):
  """Respaldo visible y conservador; nunca convierte un error en todo positivo."""
  t = quitar_acentos(texto)
  actor_tokens = [
      token
      for token in re.findall(r"[a-z0-9]+", quitar_acentos(actor_nombre))
      if len(token) >= 4
  ]
  coincidencias_actor = sum(1 for token in set(actor_tokens) if token in t)

  indicadores_homonimo = [
      "soprano",
      "cantante",
      "feria del marisco",
      "originaria de",
      "concierto",
  ]
  if coincidencias_actor > 0 and any(x in t for x in indicadores_homonimo):
    return {
        "sentimiento": "POSITIVA",
        "relevante": False,
        "confianza": 0.75,
        "motivo": "Posible homónimo sin relación con el actor político.",
        "origen": "RESPALDO_LOCAL",
    }

  patrones_negativos = [
      r"viola(?:r|ndo|cion)?.{0,25}(?:la )?ley",
      r"promocion (?:personalizada|anticipada)",
      r"actos? anticipados?",
      r"origen de (?:los )?recursos",
      r"uso (?:indebido|ilegal).{0,20}recursos",
      r"recursos publicos",
      r"corrup",
      r"nepot",
      r"desvio",
      r"fraude",
      r"ilegal",
      r"irregular",
      r"incompet",
      r"opacidad",
      r"denuncia",
      r"acus[ao]",
      r"investigad[ao]",
      r"sancion",
      r"piso parejo",
      r"ventaja injusta",
      r"favoritismo",
      r"cuestion[ao]",
      r"critica",
      r"bardas?",
      r"lonas?",
  ]
  hay_ataque = any(re.search(patron, t) for patron in patrones_negativos)
  if hay_ataque and (coincidencias_actor > 0 or not actor_tokens):
    return {
        "sentimiento": "NEGATIVA",
        "relevante": True,
        "confianza": 0.62,
        "motivo": "Se detectó un señalamiento reputacional directo.",
        "origen": "RESPALDO_LOCAL",
    }

  return {
      "sentimiento": "POSITIVA",
      "relevante": True,
      "confianza": 0.35,
      "motivo": "Sin señalamiento negativo explícito en el respaldo local.",
      "origen": "RESPALDO_LOCAL",
  }


NOMBRES_COLUMNAS_SENTIMIENTO = {
    "sentimiento",
    "sentiment",
    "sentimiento de la nota",
    "tono",
    "sentimiento nota",
}


def buscar_columna_sentimiento(df_data):
  for col_name in df_data.columns:
    if quitar_acentos(str(col_name).strip()) in NOMBRES_COLUMNAS_SENTIMIENTO:
      return col_name
  return None


def normalizar_sentimiento_archivo(valor):
  """Convierte el tono del archivo al esquema informativa/negativa."""
  sent_raw = quitar_acentos(valor).strip()
  if not sent_raw or sent_raw in {"nan", "none", "null", "n/a", "na"}:
    return None
  if any(k in sent_raw for k in ["negat", "critica", "contra", "advers"]):
    return "NEGATIVA"
  if any(
      k in sent_raw
      for k in [
          "posit",
          "neutr",
          "inform",
          "institucional",
          "favor",
          "sin sentimiento",
      ]
  ):
    return "POSITIVA"
  return None


def resultado_desde_archivo(sentimiento):
  return {
      "sentimiento": sentimiento,
      "relevante": True,
      "confianza": 1.0,
      "motivo": (
          "La IA no emitió una decisión suficientemente confiable; se conservó"
          " el sentimiento incluido en el archivo."
      ),
      "origen": "ARCHIVO_RESPALDO",
  }


def determinar_sentimiento_df(df_data, actor_nombre_target, es_tradicionales):
  """Usa archivo en tradicionales; en redes, IA con respaldo del archivo."""
  sent_col_name = buscar_columna_sentimiento(df_data)
  sentimientos_archivo = [
      normalizar_sentimiento_archivo(valor)
      for valor in (
          df_data[sent_col_name].fillna("").astype(str)
          if sent_col_name is not None
          else [""] * len(df_data)
      )
  ]

  if es_tradicionales:
    if sent_col_name is None:
      raise RuntimeError(
          "El archivo de medios tradicionales no contiene una columna de"
          " sentimiento. Para evitar una evaluación incorrecta basada solo en"
          " el título, este reporte no se generó."
      )

    sentimientos_archivo = [
        sentimiento or "POSITIVA" for sentimiento in sentimientos_archivo
    ]

    return pd.DataFrame({
        "sentimiento_final": sentimientos_archivo,
        "relevante_ia": [True] * len(df_data),
        "confianza_ia": [1.0] * len(df_data),
        "motivo_ia": [
            "Clasificación tomada directamente del archivo original."
        ] * len(df_data),
        "origen_clasificacion": ["ARCHIVO_ORIGINAL"] * len(df_data),
    })

  df_eval = df_data.reset_index(drop=True)
  resultados_finales = [None] * len(df_eval)
  sentimientos_archivo = list(sentimientos_archivo)

  if CLIENTE_GEMINI is None:
    if not all(sentimientos_archivo):
      raise RuntimeError(
          "No se encontró GEMINI_API_KEY y algunas publicaciones tampoco tienen"
          " un sentimiento válido en el archivo. Configura la clave o completa"
          " esa columna antes de generar el reporte."
      )
    st.warning(
        "No se encontró GEMINI_API_KEY. Se utilizó el sentimiento incluido en"
        " el archivo para todas las publicaciones de redes sociales."
    )
    return pd.DataFrame({
        "sentimiento_final": sentimientos_archivo,
        "relevante_ia": [True] * len(df_eval),
        "confianza_ia": [1.0] * len(df_eval),
        "motivo_ia": [
            "Clasificación tomada del archivo porque la IA no estaba disponible."
        ] * len(df_eval),
        "origen_clasificacion": ["ARCHIVO_RESPALDO"] * len(df_eval),
    })

  # Se evalúa una sola vez cada texto idéntico, pero el resultado se vuelve a
  # colocar en todas sus filas. Así se acelera sin eliminar repeticiones ni
  # alterar los conteos del reporte.
  registros_originales = [
      construir_registro_para_ia(row, fila_id)
      for fila_id, (_, row) in enumerate(df_eval.iterrows())
  ]
  clave_por_fila = []
  registro_por_clave = {}
  filas_por_clave = {}
  for fila_id, registro in enumerate(registros_originales):
    texto_normalizado = re.sub(
        r"\s+", " ", quitar_acentos(registro.get("texto", ""))
    ).strip()
    clave = texto_normalizado or f"__fila_vacia_{fila_id}"
    clave_por_fila.append(clave)
    filas_por_clave.setdefault(clave, []).append(fila_id)
    if clave not in registro_por_clave:
      registro_por_clave[clave] = {**registro, "id": len(registro_por_clave)}

  claves_unicas = list(registro_por_clave)
  registros_unicos = [registro_por_clave[clave] for clave in claves_unicas]
  sentimiento_archivo_por_clave = {}
  for clave, filas in filas_por_clave.items():
    sentimiento_archivo_por_clave[clave] = next(
        (
            sentimientos_archivo[fila]
            for fila in filas
            if sentimientos_archivo[fila]
        ),
        None,
    )
  cache_ia = st.session_state.setdefault("cache_sentimiento_ia", {})
  prefijo_cache = quitar_acentos(actor_nombre_target).strip() + "|" + GEMINI_MODEL
  pendientes = []
  resultado_por_clave = {}
  for clave, registro in zip(claves_unicas, registros_unicos):
    cache_key = prefijo_cache + "|" + clave
    if cache_key in cache_ia:
      resultado_por_clave[clave] = cache_ia[cache_key]
    else:
      pendientes.append((clave, registro))

  lote_tamano = 8
  total_lotes = max(1, (len(pendientes) + lote_tamano - 1) // lote_tamano)
  progreso = st.progress(0)
  estado_progreso = st.empty()
  fallos_respaldo = 0

  for l_idx in range(total_lotes):
    inicio = l_idx * lote_tamano
    lote_pendiente = pendientes[inicio : inicio + lote_tamano]
    if not lote_pendiente:
      break
    estado_progreso.info(
        f"Analizando lote {l_idx + 1} de {total_lotes} con Gemini... "
        f"({len(registros_unicos)} textos únicos de {len(df_eval)} publicaciones)"
    )
    lista_lote = [{**registro, "id": local_id}
                  for local_id, (_, registro) in enumerate(lote_pendiente)]

    try:
      res_map = clasificar_lote_adaptativo(lista_lote, actor_nombre_target)
    except Exception as exc:
      # Un rechazo, demora o falta de cuota no debe impedir el reporte. Para
      # cada texto se conserva el sentimiento del archivo cuando existe.
      res_map = {}
      for local_id, (clave, registro_original) in enumerate(lote_pendiente):
        sentimiento_original = sentimiento_archivo_por_clave.get(clave)
        if sentimiento_original:
          res_map[local_id] = resultado_desde_archivo(
              sentimiento_original
          )
        else:
          res_map[local_id] = clasificar_respaldo_local(
              registro_original.get("texto", ""), actor_nombre_target
          )
      st.warning(
          "La IA no pudo completar un lote. Se utilizó el sentimiento incluido"
          " en el archivo como respaldo cuando estaba disponible."
      )

    for local_id, (clave, registro_original) in enumerate(lote_pendiente):
      resultado = res_map.get(local_id)
      if resultado is None:
        resultado = clasificar_respaldo_local(
            registro_original.get("texto", ""), actor_nombre_target
        )
      if resultado.get("origen") == "RESPALDO_LOCAL":
        fallos_respaldo += 1
      resultado_por_clave[clave] = resultado
      if resultado.get("origen") == "IA":
        cache_ia[prefijo_cache + "|" + clave] = resultado

    progreso.progress((l_idx + 1) / max(total_lotes, 1))

  progreso.empty()
  estado_progreso.empty()

  for fila_id, clave in enumerate(clave_por_fila):
    resultado = dict(resultado_por_clave[clave])
    sentimiento_original = sentimientos_archivo[fila_id]
    ia_no_concluyente = (
        resultado.get("origen") != "IA"
        or float(resultado.get("confianza", 0.0)) < UMBRAL_CONFIANZA_IA
    )
    if sentimiento_original and ia_no_concluyente:
      resultado = resultado_desde_archivo(sentimiento_original)
    resultados_finales[fila_id] = resultado

  if fallos_respaldo:
    st.warning(
        f"{fallos_respaldo} texto(s) único(s) no obtuvieron respuesta de Gemini y fueron"
        " resueltos con el sentimiento original del archivo cuando estaba"
        " disponible; en los demás casos se aplicó el respaldo local."
    )

  return pd.DataFrame({
      "sentimiento_final": [r["sentimiento"] for r in resultados_finales],
      "relevante_ia": [r["relevante"] for r in resultados_finales],
      "confianza_ia": [r["confianza"] for r in resultados_finales],
      "motivo_ia": [r["motivo"] for r in resultados_finales],
      "origen_clasificacion": [r["origen"] for r in resultados_finales],
  })


def limpiar_texto_para_resumen(texto):
  if not isinstance(texto, str) or texto == "nan":
    return ""
  t = re.sub(r"https?://\S+", "", texto)
  t = re.sub(r"---\s*transcripci[oó]n\s*---[\s\S]*", "", t, flags=re.I)
  t = re.sub(r"kind:\s*captions.*", "", t, flags=re.I)
  t = re.sub(r"#([a-zA-Z0-9_]+)", r"\1", t)
  t = re.sub(r"[\U00010000-\U0010ffff]", "", t)
  t = re.sub(r":[a-zA-Z0-9_\-|]+:", "", t)
  t = re.sub(r"\s+", " ", t).strip()
  return t


def extraer_resumen_temas_real(df_data, actor_nombre):
  pos_df = df_data[df_data["sentimiento_final"].isin(["POSITIVA", "NEUTRA"])]
  neg_df = df_data[df_data["sentimiento_final"] == "NEGATIVA"]

  columnas_posibles_texto = [
      "Titulo",
      "Título",
      "Contenido",
      "Detail",
      "Summary",
      "Síntesis",
      "Sintesis",
      "Encabezado",
      "Tema",
      "Nota",
      "Title",
  ]

  pos_textos = []
  for col in columnas_posibles_texto:
    if col in pos_df.columns or any(
        quitar_acentos(str(c)) == quitar_acentos(col) for c in pos_df.columns
    ):
      serie_t = obtener_columna_serie(pos_df, [col])
      raw_list = serie_t.dropna().astype(str).tolist()
      pos_textos = [
          limpiar_texto_para_resumen(t)
          for t in raw_list
          if len(limpiar_texto_para_resumen(t)) > 10 and "teaser" not in t.lower()
      ]
      if len(pos_textos) > 0:
        break

  neg_textos = []
  for col in columnas_posibles_texto:
    if col in neg_df.columns or any(
        quitar_acentos(str(c)) == quitar_acentos(col) for c in neg_df.columns
    ):
      serie_t = obtener_columna_serie(neg_df, [col])
      raw_list = serie_t.dropna().astype(str).tolist()
      neg_textos = [
          limpiar_texto_para_resumen(t)
          for t in raw_list
          if len(limpiar_texto_para_resumen(t)) > 10 and "teaser" not in t.lower()
      ]
      if len(neg_textos) > 0:
        break

  pos_unicos = list(dict.fromkeys(pos_textos))
  neg_unicos = list(dict.fromkeys(neg_textos))

  prompt = f"""
Eres un analista senior de comunicación política y redacción ejecutiva.
Redacta el "RESUMEN" ejecutivo oficial para el actor político: "{actor_nombre}".

REGLAS DE FORMATO Y ESTILO (OBLIGATORIAS):
1. CERO DUPLICADOS: Agrupa las notas por eje temático. Redacta UN SOLO punto que sintetice el caso globalmente.
2. REDACCIÓN EJECUTIVA FLUIDA: Cada punto debe ser un párrafo fluido, profesional y descriptivo (2 a 3 líneas), explicando hechos concretos.
3. PROHIBIDO COPIAR TITULARES O PEGAR FRAGMENTOS LITERALES. Redacta con tus propias palabras en tono institucional.
4. ESTRUCTURA EXACTA:
   Temas relevantes informativos
   1. [Título del Eje]: [Descripción ejecutiva redactada formalmente].
   
   Temas negativos
   1. [Título del Eje]: [Descripción ejecutiva redactada formalmente].
   (Si no hay notas negativas, escribe: 1. No se registraron temas negativos en el periodo analizado.)
5. Redacta como máximo tres puntos informativos y tres negativos. No repitas
   encabezados ni dejes una sección sin contenido.

NOTAS POSITIVAS DISPONIBLES:
{json.dumps(pos_unicos[:25], ensure_ascii=False)}

NOTAS NEGATIVAS DISPONIBLES:
{json.dumps(neg_unicos[:25], ensure_ascii=False)}
"""
  try:
    interaction = CLIENTE_GEMINI.interactions.create(
        model=GEMINI_MODEL,
        input=prompt,
    )
    res_ia = (interaction.output_text or "").strip()
    if res_ia and len(res_ia) > 30 and len(res_ia) < 1800:
      if not res_ia.startswith("Temas relevantes informativos"):
        res_ia = "Temas relevantes informativos\n" + res_ia

      lineas_limpias = []
      encabezados_vistos = set()
      for linea in res_ia.splitlines():
        linea = linea.strip()
        if not linea:
          continue
        linea_norm = quitar_acentos(linea).strip(": ")
        if linea_norm in {
            "temas relevantes informativos",
            "temas negativos",
        }:
          if linea_norm in encabezados_vistos:
            continue
          encabezados_vistos.add(linea_norm)
          linea = (
              "Temas relevantes informativos"
              if "relevantes" in linea_norm
              else "Temas negativos"
          )
        lineas_limpias.append(linea)

      if "temas negativos" not in encabezados_vistos:
        lineas_limpias.append("Temas negativos")
        encabezados_vistos.add("temas negativos")

      idx_neg = lineas_limpias.index("Temas negativos")
      if not neg_unicos:
        lineas_limpias = lineas_limpias[: idx_neg + 1]
        lineas_limpias.append(
            "1. No se registraron temas negativos en el periodo analizado."
        )
      elif not any(
          re.match(r"^\d+\.", linea)
          for linea in lineas_limpias[idx_neg + 1 :]
      ):
        raise ValueError("La IA no redactó los temas negativos detectados.")

      return "\n".join(lineas_limpias)
  except Exception:
    pass

  # Fallback limpio
  lineas_res = ["Temas relevantes informativos"]
  if len(pos_unicos) > 0:
    for i, t in enumerate(pos_unicos[:3], 1):
      t_clean = t.split(".")[0].strip()
      lineas_res.append(
          f"{i}. Gestión y Agenda Pública: Cobertura y seguimiento a"
          f" actividades relacionadas con {t_clean[:120]}."
      )
  else:
    lineas_res.append(
        "1. Agenda Institucional: Difusión de actividades públicas y agenda de"
        " trabajo."
    )

  lineas_res.append("\nTemas negativos")
  if len(neg_unicos) > 0:
    for i, t in enumerate(neg_unicos[:3], 1):
      t_clean = t.split(".")[0].strip()
      lineas_res.append(
          f"{i}. Cuestionamientos y Crítica Pública: Señalamientos referentes a"
          f" {t_clean[:120]}."
      )
  else:
    lineas_res.append(
        "1. No se registraron temas negativos en el periodo analizado."
    )

  return "\n".join(lineas_res)


def obtener_link_inteligente(row):
  urls_en_fila = []
  campos_prioridad = [
      "Link URL Medio",
      "Link de Nota",
      "URL",
      "Enlace",
      "Link",
      "Link Medio",
      "Alcance",
  ]
  for c in campos_prioridad:
    val = obtener_campo(row, [c])
    if (
        val
        and str(val).startswith("http")
        and str(val).strip() not in urls_en_fila
    ):
      urls_en_fila.append(str(val).strip())

  for val in row.values:
    if val is not None and not pd.isna(val):
      s_val = str(val).strip()
      if s_val.startswith("http") and s_val not in urls_en_fila:
        urls_en_fila.append(s_val)

  if not urls_en_fila:
    return ""

  urls_sin_testigo = [u for u in urls_en_fila if "hanakua.mx/Testigo" not in u]
  urls_medios_externos = [
      u for u in urls_sin_testigo if "hanakua.mx" not in u.lower()
  ]
  if urls_medios_externos:
    return urls_medios_externos[0]

  urls_hanakua_notas = [u for u in urls_sin_testigo if "hanakua.mx/Notas" in u]
  if urls_hanakua_notas:
    return urls_hanakua_notas[0]

  if urls_sin_testigo:
    return urls_sin_testigo[0]

  return ""


def aplicar_verdana_12_documento(doc):
  """Aplica Verdana 12 sin alterar negritas, colores, resaltados o enlaces."""
  doc.styles["Normal"].font.name = "Verdana"
  doc.styles["Normal"].font.size = Pt(12)

  elementos = [doc._element]
  for seccion in doc.sections:
    elementos.extend([
        seccion.header._element,
        seccion.first_page_header._element,
        seccion.even_page_header._element,
        seccion.footer._element,
        seccion.first_page_footer._element,
        seccion.even_page_footer._element,
    ])

  vistos = set()
  for elemento in elementos:
    identidad = id(elemento)
    if identidad in vistos:
      continue
    vistos.add(identidad)
    for run_xml in elemento.iter(qn("w:r")):
      rpr = run_xml.find(qn("w:rPr"))
      if rpr is None:
        rpr = OxmlElement("w:rPr")
        run_xml.insert(0, rpr)

      rfonts = rpr.find(qn("w:rFonts"))
      if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
      for atributo in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{atributo}"), "Verdana")

      for etiqueta in ("w:sz", "w:szCs"):
        tamano = rpr.find(qn(etiqueta))
        if tamano is None:
          tamano = OxmlElement(etiqueta)
          rpr.append(tamano)
        tamano.set(qn("w:val"), "24")


def crear_doc_desde_hoja(df_hoja, nombre_hoja, es_redes_sociales):
  if df_hoja is None or df_hoja.empty:
    return None

  df_hoja = reparar_desfase_columnas_excel(df_hoja)

  mask_sin_notas = df_hoja.apply(
      lambda r: any(
          k in str(v).lower()
          for v in r.values
          for k in ["sin notas", "si notas", "sin nota", "sin registro"]
      ),
      axis=1,
  )
  df_hoja = df_hoja[~mask_sin_notas].copy()
  if len(df_hoja) == 0:
    return None

  df_filtrado, total_descartadas = limpiar_dataframe_redes_automatico(
      df_hoja, nombre_hoja
  )
  if len(df_filtrado) == 0:
    return None

  serie_fechas_raw = obtener_columna_serie(
      df_filtrado,
      [
          "Publish date",
          "Fecha",
          "Date",
          "Fecha de publicación",
          "Fecha de publicacion",
      ],
  )
  df_filtrado["fecha_dt"] = serie_fechas_raw.apply(parsear_fecha_perfecta)
  df_filtrado = df_filtrado.dropna(subset=["fecha_dt"]).sort_values(
      by="fecha_dt", ascending=True
  )

  if len(df_filtrado) == 0:
    return None

  subset_dup = [
      c
      for c in ["ID Nota", "Titulo", "Link de Nota", "Link URL Medio"]
      if c in df_filtrado.columns
  ]
  if len(subset_dup) > 0:
    df_filtrado = df_filtrado.drop_duplicates(subset=subset_dup)

  df_filtrado = df_filtrado.reset_index(drop=True)
  resultados_ia = determinar_sentimiento_df(
      df_filtrado, nombre_hoja, es_tradicionales=not es_redes_sociales
  )
  for columna in resultados_ia.columns:
    df_filtrado[columna] = resultados_ia[columna].values

  total_irrelevantes = int((~df_filtrado["relevante_ia"]).sum())
  df_filtrado = df_filtrado[df_filtrado["relevante_ia"]].copy()
  if total_irrelevantes:
    st.info(
        f"La IA excluyó {total_irrelevantes} mención(es) por tratarse de"
        " homónimos o publicaciones ajenas al actor evaluado."
    )
  if len(df_filtrado) == 0:
    return None

  fechas_validas = df_filtrado["fecha_dt"]
  min_d = fechas_validas.min()
  max_d = fechas_validas.max()

  if min_d.strftime("%d.%m") == max_d.strftime("%d.%m"):
    periodo_texto = (
        f"{min_d.strftime('%d')} de {MESES_ES[max_d.month]} de {max_d.year}"
    )
  else:
    periodo_texto = (
        f"{min_d.strftime('%d')} al {max_d.strftime('%d')} de"
        f" {MESES_ES[max_d.month]} de {max_d.year}"
    )

  df_filtrado["fecha_str"] = df_filtrado["fecha_dt"].dt.strftime("%d.%m.%y")

  positivas_cnt = len(
      df_filtrado[df_filtrado["sentimiento_final"].isin(["POSITIVA", "NEUTRA"])]
  )
  negativas_cnt = len(
      df_filtrado[df_filtrado["sentimiento_final"] == "NEGATIVA"]
  )
  total_cnt = len(df_filtrado)

  serie_media_raw = obtener_columna_serie(
      df_filtrado,
      [
          "Tipo de Medio",
          "Fuente",
          "Media type",
          "Media Type",
          "Medio",
          "Nombre del Medio",
          "Canal",
          "Tipo de Nota",
      ],
  )
  df_filtrado["categoria_medio_std"] = (
      serie_media_raw.apply(estandarizar_categoria_medio)
      if not es_redes_sociales
      else "REDES SOCIALES"
  )

  tv_cnt = (
      (df_filtrado["categoria_medio_std"] == "TELEVISIÓN").sum()
      if not es_redes_sociales
      else 0
  )
  rad_cnt = (
      (df_filtrado["categoria_medio_std"] == "RADIO").sum()
      if not es_redes_sociales
      else 0
  )
  portales_cnt = (
      (df_filtrado["categoria_medio_std"] == "PORTALES DIGITALES").sum()
      if not es_redes_sociales
      else 0
  )
  prensa_cnt = (
      (df_filtrado["categoria_medio_std"] == "PRENSA LOCAL").sum()
      if not es_redes_sociales
      else 0
  )
  columnas_cnt = (
      (df_filtrado["categoria_medio_std"] == "COLUMNAS").sum()
      if not es_redes_sociales
      else 0
  )

  doc = Document()
  for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

  doc.styles["Normal"].font.name = "Verdana"
  doc.styles["Normal"].font.size = Pt(10)

  def add_run_verdana(
      p,
      text,
      bold=False,
      italic=False,
      size_pt=10,
      color_rgb=None,
      underline=False,
  ):
    run = p.add_run(text)
    run.font.name = "Verdana"
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size_pt)
    run.underline = underline
    if color_rgb:
      run.font.color.rgb = color_rgb
    return run

  def fondo_celda(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

  # 1. Encabezado
  p_title = doc.add_paragraph()
  add_run_verdana(
      p_title,
      nombre_hoja.upper(),
      bold=True,
      size_pt=12,
      color_rgb=RGBColor(0, 51, 102),
  )

  p_per = doc.add_paragraph()
  add_run_verdana(
      p_per, f"PERIODO DE MEDICIÓN: {periodo_texto}", bold=True, size_pt=10
  )

  p_can = doc.add_paragraph()
  p_can.paragraph_format.space_after = Pt(8)
  add_run_verdana(
      p_can,
      "CANALES: PRENSA, TV, RADIO, PORTALES, REDES SOCIALES Y COLUMNAS.",
      bold=True,
      size_pt=9.5,
  )

  # 2. Balance de Impactos
  p_bal = doc.add_paragraph()
  add_run_verdana(p_bal, "BALANCE DE IMPACTOS", bold=True, size_pt=10.5)

  t_imp = doc.add_table(rows=2, cols=3)
  t_imp.alignment = WD_TABLE_ALIGNMENT.CENTER

  headers = ["POSITIVA / INFORMATIVA", "NEGATIVA", "TOTAL DE IMPACTOS"]
  fills = ["E2EFDA", "FCE4D6", "D9E1F2"]

  for col_idx, (h_text, fill_color) in enumerate(zip(headers, fills)):
    cell = t_imp.cell(0, col_idx)
    fondo_celda(cell, fill_color)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_verdana(p, h_text, bold=True, size_pt=9.5)

  val_counts = [str(positivas_cnt), str(negativas_cnt), str(total_cnt)]
  for col_idx, val_text in enumerate(val_counts):
    cell = t_imp.cell(1, col_idx)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_verdana(p, val_text, bold=True, size_pt=11)

  p_tot = doc.add_paragraph()
  p_tot.paragraph_format.space_before = Pt(10)
  p_tot.paragraph_format.space_after = Pt(10)

  if es_redes_sociales:
    partes = [
        f"TOTAL NOTAS INFORMATIVAS: {total_cnt}",
        f"REDES SOCIALES: {total_cnt}",
        "PORTALES DIGITALES: 0",
        "PRENSA LOCAL: 0",
        "COLUMNAS: 0",
    ]
  else:
    partes = [f"TOTAL NOTAS INFORMATIVAS: {total_cnt}"]
    if tv_cnt > 0:
      partes.append(f"TELEVISIÓN: {tv_cnt}")
    if rad_cnt > 0:
      partes.append(f"RADIO: {rad_cnt}")
    if portales_cnt > 0:
      partes.append(f"PORTALES DIGITALES: {portales_cnt}")
    if prensa_cnt > 0:
      partes.append(f"PRENSA LOCAL: {prensa_cnt}")
    if columnas_cnt > 0:
      partes.append(f"COLUMNAS: {columnas_cnt}")
    if len(partes) == 1:
      partes.extend([
          "PORTALES DIGITALES: 0",
          "TELEVISIÓN: 0",
          "RADIO: 0",
          "PRENSA LOCAL: 0",
          "COLUMNAS: 0",
      ])

  texto_totales = "\n".join(partes)
  add_run_verdana(p_tot, texto_totales, bold=True, size_pt=10)

  # 3. Resumen
  p_res = doc.add_paragraph()
  p_res.paragraph_format.space_before = Pt(10)
  add_run_verdana(p_res, "RESUMEN", bold=True, size_pt=11)

  temas_texto = extraer_resumen_temas_real(df_filtrado, nombre_hoja)
  for linea in temas_texto.split("\n"):
    linea_clean = linea.strip()
    if linea_clean:
      p_t = doc.add_paragraph()
      p_t.paragraph_format.space_before = Pt(1)
      p_t.paragraph_format.space_after = Pt(2)
      if "TEMAS RELEVANTES INFORMATIVOS" in linea_clean.upper():
        add_run_verdana(
            p_t, "Temas relevantes informativos", bold=True, size_pt=10
        )
      elif "TEMAS NEGATIVOS" in linea_clean.upper():
        p_t.paragraph_format.space_before = Pt(4)
        add_run_verdana(
            p_t,
            "Temas negativos",
            bold=True,
            size_pt=10,
            color_rgb=RGBColor(180, 0, 0),
        )
      else:
        add_run_verdana(p_t, linea_clean, size_pt=9.5)

  # 4. Desglose
  p_des = doc.add_paragraph()
  p_des.paragraph_format.space_before = Pt(12)
  add_run_verdana(p_des, "DESGLOSE", bold=True, size_pt=11)

  orden_medios_oficial = [
      "TELEVISIÓN",
      "RADIO",
      "PORTALES DIGITALES",
      "PRENSA LOCAL",
      "COLUMNAS",
      "REDES SOCIALES",
  ]

  for fecha_dt_val, sub_df in df_filtrado.groupby("fecha_dt", sort=True):
    fecha_item = fecha_dt_val.strftime("%d.%m.%y")

    p_f = doc.add_paragraph()
    p_f.paragraph_format.space_before = Pt(10)
    p_f.paragraph_format.space_after = Pt(2)
    add_run_verdana(
        p_f,
        fecha_item,
        bold=True,
        size_pt=10.5,
        color_rgb=RGBColor(0, 51, 102),
    )

    pos_df = sub_df[sub_df["sentimiento_final"].isin(["POSITIVA", "NEUTRA"])]
    neg_df = sub_df[sub_df["sentimiento_final"] == "NEGATIVA"]

    if es_redes_sociales or (
        len(pos_df) > 0
        and pos_df["categoria_medio_std"].iloc[0] == "REDES SOCIALES"
    ):
      if len(pos_df) > 0:
        p_m = doc.add_paragraph()
        p_m.paragraph_format.space_before = Pt(4)
        p_m.paragraph_format.space_after = Pt(4)
        add_run_verdana(
            p_m, f"REDES SOCIALES: {len(pos_df)}", bold=True, size_pt=10
        )

        for _, row in pos_df.iterrows():
          autor = obtener_campo(
              row,
              ["Autor", "Author name", "Fuente", "Media name", "Programa"],
          )
          handle = obtener_campo(
              row, ["Author handle (@username)", "Handle", "Username"]
          )
          detalle = obtener_campo(
              row,
              [
                  "Contenido",
                  "Detail",
                  "Summary",
                  "Síntesis",
                  "Sintesis",
                  "Titulo",
                  "Título",
                  "Title",
                  "Encabezado",
              ],
          )
          link = obtener_link_inteligente(row)

          p_a = doc.add_paragraph()
          p_a.paragraph_format.space_before = Pt(4)
          p_a.paragraph_format.space_after = Pt(1)
          if handle and not handle.startswith("@"):
            handle = f"@{handle}"
          add_run_verdana(
              p_a,
              f"{autor} {handle}".strip() if handle else autor,
              bold=True,
              size_pt=10,
          )

          p_d = doc.add_paragraph()
          p_d.paragraph_format.space_after = Pt(2)
          add_run_verdana(
              p_d, limpiar_texto(detalle), bold=False, size_pt=9.5
          )

          if link:
            p_l = doc.add_paragraph()
            p_l.paragraph_format.space_after = Pt(6)
            add_run_verdana(
                p_l,
                link,
                bold=False,
                size_pt=9,
                color_rgb=RGBColor(0, 102, 204),
                underline=True,
            )

      if len(neg_df) > 0:
        p_neg_hdr = doc.add_paragraph()
        p_neg_hdr.paragraph_format.space_before = Pt(6)
        p_neg_hdr.paragraph_format.space_after = Pt(4)
        add_run_verdana(
            p_neg_hdr,
            f"NEGATIVAS: {len(neg_df)}",
            bold=True,
            size_pt=10,
            color_rgb=RGBColor(180, 0, 0),
        )

        for _, row in neg_df.iterrows():
          autor = obtener_campo(
              row,
              ["Autor", "Author name", "Fuente", "Media name", "Programa"],
          )
          handle = obtener_campo(
              row, ["Author handle (@username)", "Handle", "Username"]
          )
          detalle = obtener_campo(
              row,
              [
                  "Contenido",
                  "Detail",
                  "Summary",
                  "Síntesis",
                  "Sintesis",
                  "Titulo",
                  "Título",
                  "Title",
                  "Encabezado",
              ],
          )
          link = obtener_link_inteligente(row)

          p_a = doc.add_paragraph()
          p_a.paragraph_format.space_before = Pt(4)
          p_a.paragraph_format.space_after = Pt(1)
          if handle and not handle.startswith("@"):
            handle = f"@{handle}"
          add_run_verdana(
              p_a,
              f"{autor} {handle}".strip() if handle else autor,
              bold=True,
              size_pt=10,
          )

          p_d = doc.add_paragraph()
          p_d.paragraph_format.space_after = Pt(2)
          add_run_verdana(
              p_d, limpiar_texto(detalle), bold=False, size_pt=9.5
          )

          if link:
            p_l = doc.add_paragraph()
            p_l.paragraph_format.space_after = Pt(6)
            add_run_verdana(
                p_l,
                link,
                bold=False,
                size_pt=9,
                color_rgb=RGBColor(0, 102, 204),
                underline=True,
            )

    else:
      if len(pos_df) > 0:
        for cat_nombre in orden_medios_oficial:
          sub_pos_cat = pos_df[pos_df["categoria_medio_std"] == cat_nombre]
          if len(sub_pos_cat) > 0:
            p_m = doc.add_paragraph()
            p_m.paragraph_format.space_before = Pt(4)
            p_m.paragraph_format.space_after = Pt(4)
            add_run_verdana(
                p_m, f"{cat_nombre}: {len(sub_pos_cat)}", bold=True, size_pt=10
            )

            for _, row in sub_pos_cat.iterrows():
              medio = obtener_campo(
                  row, ["Nombre del Medio", "Fuente", "Media name", "Medio"]
              )
              autor = obtener_campo(
                  row, ["Autor", "Author name", "Programa", "Conductor"]
              )
              hora = obtener_campo(
                  row,
                  [
                      "Hora",
                      "Hour",
                      "Time",
                      "Hora de Transmisión",
                      "Hora de Transmision",
                  ],
              )
              titulo = obtener_campo(
                  row,
                  [
                      "Titulo",
                      "Título",
                      "Contenido",
                      "Detail",
                      "Summary",
                      "Síntesis",
                      "Sintesis",
                      "Encabezado",
                      "Nota",
                  ],
              )
              link = obtener_link_inteligente(row)

              p_a = doc.add_paragraph()
              p_a.paragraph_format.space_before = Pt(4)
              p_a.paragraph_format.space_after = Pt(1)
              cabecera = (
                  f"{medio} - {autor}"
                  if (
                      autor
                      and autor
                      not in ["Redacción", "Staff", "Online", medio]
                  )
                  else medio
              )
              add_run_verdana(p_a, cabecera, bold=True, size_pt=10)

              cuerpo_texto = limpiar_texto(titulo)
              if hora and not cuerpo_texto.lower().startswith(hora.lower()):
                cuerpo_texto = f"{hora} {cuerpo_texto}".strip()

              p_d = doc.add_paragraph()
              p_d.paragraph_format.space_after = Pt(2)
              add_run_verdana(
                  p_d, cuerpo_texto, bold=False, size_pt=9.5
              )

              if link:
                p_l = doc.add_paragraph()
                p_l.paragraph_format.space_after = Pt(6)
                add_run_verdana(
                    p_l,
                    link,
                    bold=False,
                    size_pt=9,
                    color_rgb=RGBColor(0, 102, 204),
                    underline=True,
                )

      if len(neg_df) > 0:
        p_neg_hdr = doc.add_paragraph()
        p_neg_hdr.paragraph_format.space_before = Pt(6)
        p_neg_hdr.paragraph_format.space_after = Pt(4)
        add_run_verdana(
            p_neg_hdr,
            f"NEGATIVAS: {len(neg_df)}",
            bold=True,
            size_pt=10,
            color_rgb=RGBColor(180, 0, 0),
        )

        for cat_nombre in orden_medios_oficial:
          sub_neg_cat = neg_df[neg_df["categoria_medio_std"] == cat_nombre]
          if len(sub_neg_cat) > 0:
            p_sub_neg = doc.add_paragraph()
            p_sub_neg.paragraph_format.space_before = Pt(4)
            p_sub_neg.paragraph_format.space_after = Pt(4)
            add_run_verdana(
                p_sub_neg,
                f"{cat_nombre}: {len(sub_neg_cat)}",
                bold=True,
                size_pt=10,
                color_rgb=RGBColor(180, 0, 0),
            )

            for _, row in sub_neg_cat.iterrows():
              medio = obtener_campo(
                  row, ["Nombre del Medio", "Fuente", "Media name", "Medio"]
              )
              autor = obtener_campo(
                  row, ["Autor", "Author name", "Programa", "Conductor"]
              )
              hora = obtener_campo(
                  row,
                  [
                      "Hora",
                      "Hour",
                      "Time",
                      "Hora de Transmisión",
                      "Hora de Transmision",
                  ],
              )
              titulo = obtener_campo(
                  row,
                  [
                      "Titulo",
                      "Título",
                      "Contenido",
                      "Detail",
                      "Summary",
                      "Síntesis",
                      "Sintesis",
                      "Encabezado",
                      "Nota",
                  ],
              )
              link = obtener_link_inteligente(row)

              p_a = doc.add_paragraph()
              p_a.paragraph_format.space_before = Pt(4)
              p_a.paragraph_format.space_after = Pt(1)
              cabecera = (
                  f"{medio} - {autor}"
                  if (
                      autor
                      and autor
                      not in ["Redacción", "Staff", "Online", medio]
                  )
                  else medio
              )
              add_run_verdana(p_a, cabecera, bold=True, size_pt=10)

              cuerpo_texto = limpiar_texto(titulo)
              if hora and not cuerpo_texto.lower().startswith(hora.lower()):
                cuerpo_texto = f"{hora} {cuerpo_texto}".strip()

              p_d = doc.add_paragraph()
              p_d.paragraph_format.space_after = Pt(2)
              add_run_verdana(
                  p_d, cuerpo_texto, bold=False, size_pt=9.5
              )

              if link:
                p_l = doc.add_paragraph()
                p_l.paragraph_format.space_after = Pt(6)
                add_run_verdana(
                    p_l,
                    link,
                    bold=False,
                    size_pt=9,
                    color_rgb=RGBColor(0, 102, 204),
                    underline=True,
                )

  buffer = io.BytesIO()
  aplicar_verdana_12_documento(doc)
  doc.save(buffer)
  buffer.seek(0)
  return buffer


# ==============================================================================
# UNIFICACIÓN DE REPORTES WORD
# Usa el primer archivo (medios tradicionales) como plantilla visual y combina
# en él las publicaciones del reporte de redes sociales.
# ==============================================================================

PATRON_FECHA_REPORTE = re.compile(
    r"^\s*(\d{2})\.(\d{2})\.(\d{2}|\d{4})\s*$"
)
ORDEN_CANALES_UNIFICADOS = [
    "ENTREVISTAS",
    "TELEVISIÓN",
    "RADIO",
    "PRENSA LOCAL",
    "PORTALES DIGITALES",
    "COLUMNAS",
    "REDES SOCIALES",
]


def _abrir_docx_desde_streamlit(archivo):
  """Abre bytes, BytesIO o UploadedFile sin depender de su posición actual."""
  if archivo is None:
    raise ValueError("Falta uno de los archivos Word.")
  if isinstance(archivo, (bytes, bytearray)):
    contenido = bytes(archivo)
  elif hasattr(archivo, "getvalue"):
    contenido = archivo.getvalue()
  else:
    posicion = archivo.tell() if hasattr(archivo, "tell") else None
    contenido = archivo.read()
    if posicion is not None and hasattr(archivo, "seek"):
      archivo.seek(posicion)
  if not contenido:
    raise ValueError("Uno de los archivos Word está vacío.")
  return Document(io.BytesIO(contenido))


def _texto_normalizado_docx(texto):
  return " ".join(str(texto or "").replace("\xa0", " ").split())


def _extraer_entero_docx(texto):
  coincidencia = re.search(r"\d[\d,.]*", str(texto or ""))
  if not coincidencia:
    return None
  solo_digitos = re.sub(r"\D", "", coincidencia.group(0))
  return int(solo_digitos) if solo_digitos else None


def _leer_balance_docx(doc):
  """Devuelve positiva, negativa y total desde la tabla de balance."""
  for tabla in doc.tables:
    for indice, fila in enumerate(tabla.rows[:-1]):
      celdas = [_texto_normalizado_docx(c.text).upper() for c in fila.cells]
      encabezado = " | ".join(celdas)
      if "POSITIVA" not in encabezado or "NEGATIVA" not in encabezado:
        continue
      fila_valores = tabla.rows[indice + 1]
      valores = [_extraer_entero_docx(c.text) for c in fila_valores.cells[:3]]
      if len(valores) >= 3 and all(v is not None for v in valores):
        return valores[0], valores[1], valores[2]
  return None, None, None


PATRON_ETIQUETA_CANAL = re.compile(
    r"(ENTREVISTAS?|TELEVISI[ÓO]N|TV|RADIO|PRENSA(?:\s+LOCAL)?|"
    r"PORTALES?(?:\s+(?:DIGITALES?|LOCALES?))?|LOCALES?|COLUMNAS?|"
    r"REDES(?:\s+SOCIALES)?)"
    r"(?:\s+(INFORMATIVAS?|NEGATIVAS?))?\s*:",
    re.IGNORECASE,
)


def _canal_unificado(etiqueta):
  limpio = quitar_acentos(etiqueta).upper()
  if limpio.startswith("ENTREVISTA"):
    return "ENTREVISTAS"
  if limpio.startswith("TV") or limpio.startswith("TELEVISION"):
    return "TELEVISIÓN"
  if limpio.startswith("RADIO"):
    return "RADIO"
  if limpio.startswith("PRENSA"):
    return "PRENSA LOCAL"
  if limpio.startswith("PORTAL") or limpio in {"LOCAL", "LOCALES"}:
    return "PORTALES DIGITALES"
  if limpio.startswith("COLUMNA"):
    return "COLUMNAS"
  if limpio.startswith("REDES"):
    return "REDES SOCIALES"
  return None


def _extraer_conteos_canales(texto):
  """
  Lee uno o varios encabezados dentro del mismo párrafo.

  Admite, entre otras variantes: PORTALES DIGITALES, PORTALES LOCALES,
  LOCALES, REDES y REDES SOCIALES. También resuelve párrafos concatenados como
  «COLUMNAS: PORTALES LOCALES: 178» sin atribuir 178 a COLUMNAS.
  """
  texto = str(texto or "").replace("\xa0", " ")
  coincidencias = list(PATRON_ETIQUETA_CANAL.finditer(texto))
  if not coincidencias:
    return []

  # Solo se consideran encabezados que comienzan el párrafo. Así una mención
  # casual a «redes sociales:» dentro del cuerpo de una nota no se contabiliza.
  prefijo = texto[:coincidencias[0].start()]
  if re.search(r"[A-Za-zÁÉÍÓÚÑáéíóúñ0-9]", prefijo):
    return []

  resultados = []
  for indice, coincidencia in enumerate(coincidencias):
    fin = (
        coincidencias[indice + 1].start()
        if indice + 1 < len(coincidencias)
        else len(texto)
    )
    segmento = texto[coincidencia.end():fin]
    valor = re.match(r"\s*\(?\s*(\d[\d,.]*)", segmento)
    cantidad = None
    if valor:
      solo_digitos = re.sub(r"\D", "", valor.group(1))
      if solo_digitos:
        cantidad = int(solo_digitos)
    resultados.append({
        "canal": _canal_unificado(coincidencia.group(1)),
        "modificador": quitar_acentos(
            coincidencia.group(2) or ""
        ).upper(),
        "cantidad": cantidad,
    })
  return resultados


def _buscar_indice_desglose(doc):
  for indice, parrafo in enumerate(doc.paragraphs):
    if _texto_normalizado_docx(parrafo.text).upper() == "DESGLOSE":
      return indice
  return None


def _contar_canales_en_desglose(doc):
  """Obtiene conteos por canal y sentimiento a partir de cada bloque diario."""
  positivos = {canal: 0 for canal in ORDEN_CANALES_UNIFICADOS}
  negativos = {canal: 0 for canal in ORDEN_CANALES_UNIFICADOS}
  indice_desglose = _buscar_indice_desglose(doc)
  if indice_desglose is None:
    return positivos, negativos

  sentimiento_actual = "POSITIVA"
  for parrafo in doc.paragraphs[indice_desglose + 1:]:
    texto = str(parrafo.text or "").strip()
    if not texto:
      continue
    texto_plano = _texto_normalizado_docx(texto)
    texto_mayus = quitar_acentos(texto_plano).upper()

    if PATRON_FECHA_REPORTE.match(texto_plano):
      sentimiento_actual = "POSITIVA"
      continue
    if texto_mayus.startswith("TOTAL DE IMPACTOS INFORMATIVOS"):
      sentimiento_actual = "POSITIVA"
      continue
    if (
        texto_mayus.startswith("TOTAL DE IMPACTOS NEGATIVOS")
        or texto_mayus.startswith("TOTAL NOTAS NEGATIVAS")
        or texto_mayus.startswith("NEGATIVAS:")
    ):
      sentimiento_actual = "NEGATIVA"
      continue

    conteos = _extraer_conteos_canales(texto_plano)
    if not conteos:
      continue
    for conteo in conteos:
      canal = conteo["canal"]
      cantidad = conteo["cantidad"]
      if not canal or cantidad is None:
        continue
      sentimiento = sentimiento_actual
      if conteo["modificador"].startswith("NEGATIVA"):
        sentimiento = "NEGATIVA"
      elif conteo["modificador"].startswith("INFORMATIVA"):
        sentimiento = "POSITIVA"
      destino = negativos if sentimiento == "NEGATIVA" else positivos
      destino[canal] += cantidad

  return positivos, negativos


def _leer_totales_canales_superiores(doc):
  """Lee los totales generales por canal ubicados antes de RESUMEN."""
  totales = {canal: 0 for canal in ORDEN_CANALES_UNIFICADOS}
  for parrafo in doc.paragraphs:
    texto = str(parrafo.text or "")
    if _texto_normalizado_docx(texto).upper() == "RESUMEN":
      break
    for linea in texto.splitlines():
      for conteo in _extraer_conteos_canales(linea):
        canal = conteo["canal"]
        cantidad = conteo["cantidad"]
        if canal and cantidad is not None:
          totales[canal] += cantidad
  return totales


def _obtener_metricas_docx(doc, es_reporte_redes=False):
  positivos, negativos = _contar_canales_en_desglose(doc)
  positiva_balance, negativa_balance, total_balance = _leer_balance_docx(doc)
  totales_superiores = _leer_totales_canales_superiores(doc)

  for canal, total_canal in totales_superiores.items():
    conocido = positivos[canal] + negativos[canal]
    if total_canal > conocido:
      positivos[canal] += total_canal - conocido

  if es_reporte_redes:
    # En reportes de redes, el resumen superior puede mostrar el total del
    # canal (informativas + negativas). La tabla de balance es la fuente que
    # separa correctamente ambos sentimientos, por lo que debe prevalecer.
    if positiva_balance is not None:
      positivos["REDES SOCIALES"] = positiva_balance
    if negativa_balance is not None:
      negativos["REDES SOCIALES"] = negativa_balance

  positiva_calculada = sum(positivos.values())
  negativa_calculada = sum(negativos.values())
  positiva = max(positiva_balance or 0, positiva_calculada)
  negativa = max(negativa_balance or 0, negativa_calculada)
  total = positiva + negativa
  return {
      "positiva": int(positiva),
      "negativa": int(negativa),
      "total": int(total),
      "positivos_canal": positivos,
      "negativos_canal": negativos,
  }


def _extraer_bloques_por_fecha(doc):
  """Conserva cada párrafo del desglose con su formato OOXML original."""
  indice_desglose = _buscar_indice_desglose(doc)
  if indice_desglose is None:
    raise ValueError("El archivo no contiene la sección DESGLOSE.")

  bloques = {}
  fecha_actual = None
  for parrafo in doc.paragraphs[indice_desglose + 1:]:
    texto = _texto_normalizado_docx(parrafo.text)
    coincidencia = PATRON_FECHA_REPORTE.match(texto)
    if coincidencia:
      dia, mes, anio = map(int, coincidencia.groups())
      fecha_actual = f"{dia:02d}.{mes:02d}.{anio % 100:02d}"
      # Algunos documentos repiten el encabezado de una misma fecha. Se
      # conserva un solo encabezado y se agregan debajo todas sus notas.
      if fecha_actual not in bloques:
        bloques[fecha_actual] = [parrafo._p]
    elif fecha_actual is not None:
      bloques[fecha_actual].append(parrafo._p)
  return bloques


def _fecha_desde_etiqueta(etiqueta):
  coincidencia = PATRON_FECHA_REPORTE.match(etiqueta)
  if not coincidencia:
    return datetime.max
  dia, mes, anio = map(int, coincidencia.groups())
  anio_completo = anio if anio >= 1000 else 2000 + anio
  return datetime(anio_completo, mes, dia)


def _titulo_seccion_temas_docx(texto):
  normalizado = quitar_acentos(
      _texto_normalizado_docx(texto)
  ).upper().rstrip(" :.-")
  if normalizado in {
      "TEMAS RELEVANTES INFORMATIVOS",
      "TEMAS INFORMATIVOS",
      "TEMAS POSITIVOS",
      "TEMAS POSITIVOS E INFORMATIVOS",
  }:
    return "INFORMATIVOS"
  if normalizado in {
      "TEMAS NEGATIVOS",
      "TEMAS RELEVANTES NEGATIVOS",
      "TEMAS NEGATIVOS RELEVANTES",
  }:
    return "NEGATIVOS"
  return None


def _extraer_temas_docx(doc):
  informativos = []
  negativos = []
  estado = None
  for parrafo in doc.paragraphs:
    texto = _texto_normalizado_docx(parrafo.text)
    seccion = _titulo_seccion_temas_docx(texto)
    if seccion:
      estado = seccion
      continue
    texto_norm = quitar_acentos(texto).upper().rstrip(" :.-")
    if texto_norm == "DESGLOSE":
      break
    if not texto or estado is None:
      continue
    tema = re.sub(
        r"^\s*(?:[•\-]|\d+\s*(?:[.)]\s*-?|[-]))\s*",
        "",
        texto,
    ).strip()
    if not tema or not re.search(r"[A-Za-zÁÉÍÓÚÑáéíóúñ0-9]", tema):
      continue
    (informativos if estado == "INFORMATIVOS" else negativos).append(tema)
  return informativos, negativos


def _deduplicar_temas(temas):
  resultado = []
  vistos = set()
  for tema in temas:
    clave = normalizar_cadena(tema)
    if clave and clave not in vistos:
      vistos.add(clave)
      resultado.append(tema)
  return resultado


def _reemplazar_texto_con_formato(parrafo, nuevo_texto):
  """Cambia el texto manteniendo el formato del primer run del párrafo."""
  p_xml = parrafo._p
  primer_rpr = None
  for run in p_xml.findall(qn("w:r")):
    rpr = run.find(qn("w:rPr"))
    if rpr is not None:
      primer_rpr = deepcopy(rpr)
      break
  for hijo in list(p_xml):
    if hijo.tag != qn("w:pPr"):
      p_xml.remove(hijo)
  run_xml = OxmlElement("w:r")
  if primer_rpr is not None:
    run_xml.append(primer_rpr)
  partes = str(nuevo_texto).split("\n")
  for indice, parte in enumerate(partes):
    if indice:
      run_xml.append(OxmlElement("w:br"))
    texto_xml = OxmlElement("w:t")
    texto_xml.set(qn("xml:space"), "preserve")
    texto_xml.text = parte
    run_xml.append(texto_xml)
  p_xml.append(run_xml)


def _clonar_parrafo_con_texto(parrafo_modelo, texto):
  clon = deepcopy(parrafo_modelo._p)
  primer_rpr = None
  for run in clon.findall(qn("w:r")):
    rpr = run.find(qn("w:rPr"))
    if rpr is not None:
      primer_rpr = deepcopy(rpr)
      break
  for hijo in list(clon):
    if hijo.tag != qn("w:pPr"):
      clon.remove(hijo)
  run_xml = OxmlElement("w:r")
  if primer_rpr is not None:
    run_xml.append(primer_rpr)
  texto_xml = OxmlElement("w:t")
  texto_xml.set(qn("xml:space"), "preserve")
  texto_xml.text = texto
  run_xml.append(texto_xml)
  clon.append(run_xml)
  return clon


def _copiar_elemento_con_relaciones(elemento, doc_origen, doc_destino):
  clon = deepcopy(elemento)
  if doc_origen.part is doc_destino.part:
    return clon

  atributos_relacion = [qn("r:id"), qn("r:embed"), qn("r:link")]
  relaciones_cache = {}
  for nodo in clon.iter():
    for atributo in atributos_relacion:
      rel_id = nodo.get(atributo)
      if not rel_id or rel_id not in doc_origen.part.rels:
        continue
      if rel_id not in relaciones_cache:
        relacion = doc_origen.part.rels[rel_id]
        if relacion.is_external:
          nuevo_rel_id = doc_destino.part.relate_to(
              relacion.target_ref,
              relacion.reltype,
              is_external=True,
          )
        else:
          nuevo_rel_id = doc_destino.part.relate_to(
              relacion.target_part,
              relacion.reltype,
          )
        relaciones_cache[rel_id] = nuevo_rel_id
      nodo.set(atributo, relaciones_cache[rel_id])

  # La configuración de secciones siempre debe seguir siendo la del segundo
  # archivo; por eso se retiran saltos de sección internos del archivo primero.
  for ppr in clon.iter(qn("w:pPr")):
    sect_pr = ppr.find(qn("w:sectPr"))
    if sect_pr is not None:
      ppr.remove(sect_pr)
  return clon


def _actualizar_periodo_desde_fechas(doc, fechas):
  if not fechas:
    return
  fechas_dt = sorted(_fecha_desde_etiqueta(fecha) for fecha in fechas)
  inicio, fin = fechas_dt[0], fechas_dt[-1]
  if inicio.date() == fin.date():
    periodo = f"{inicio.day:02d} de {MESES_ES[inicio.month]} de {inicio.year}"
  elif inicio.month == fin.month and inicio.year == fin.year:
    periodo = (
        f"{inicio.day:02d} al {fin.day:02d} de {MESES_ES[fin.month]}"
        f" de {fin.year}"
    )
  elif inicio.year == fin.year:
    periodo = (
        f"{inicio.day:02d} de {MESES_ES[inicio.month]} al {fin.day:02d} de"
        f" {MESES_ES[fin.month]} de {fin.year}"
    )
  else:
    periodo = (
        f"{inicio.day:02d} de {MESES_ES[inicio.month]} de {inicio.year} al"
        f" {fin.day:02d} de {MESES_ES[fin.month]} de {fin.year}"
    )
  parrafos = list(doc.paragraphs)
  for tabla in doc.tables:
    for fila in tabla.rows:
      for celda in fila.cells:
        parrafos.extend(celda.paragraphs)
  for parrafo in parrafos:
    if _texto_normalizado_docx(parrafo.text).upper().startswith(
        "PERIODO DE MEDICIÓN:"
    ):
      _reemplazar_texto_con_formato(
          parrafo, f"PERIODO DE MEDICIÓN: {periodo}"
      )
      break


def _actualizar_balance_unificado(doc, positiva, negativa, total):
  for tabla in doc.tables:
    for indice, fila in enumerate(tabla.rows[:-1]):
      encabezado = " | ".join(
          _texto_normalizado_docx(c.text).upper() for c in fila.cells
      )
      if "POSITIVA" in encabezado and "NEGATIVA" in encabezado:
        valores = [positiva, negativa, total]
        for celda, valor in zip(tabla.rows[indice + 1].cells[:3], valores):
          _reemplazar_texto_con_formato(celda.paragraphs[0], str(valor))
        return
  raise ValueError(
      "El segundo archivo no contiene una tabla de BALANCE DE IMPACTOS válida."
  )


def _actualizar_totales_unificados(doc, metricas):
  """Actualiza el resumen sin destruir los bloques verde y rojo del formato."""
  parrafos = doc.paragraphs
  indice_info = next(
      (
          i for i, p in enumerate(parrafos)
          if "TOTAL NOTAS INFORMATIVAS:" in p.text.upper()
      ),
      None,
  )
  indice_neg = next(
      (
          i for i, p in enumerate(parrafos)
          if "TOTAL NOTAS NEGATIVAS:" in p.text.upper()
      ),
      None,
  )

  # Formato editorial: un título verde y un título rojo con sus desgloses
  # separados. Se localizan por contenido para no depender de índices fijos.
  if indice_info is not None and indice_neg is not None:
    _reemplazar_texto_con_formato(
        parrafos[indice_info],
        f"TOTAL NOTAS INFORMATIVAS: {metricas['positiva']}",
    )
    if indice_info + 1 < indice_neg:
      positivos = metricas["positivos_canal"]
      _reemplazar_texto_con_formato(
          parrafos[indice_info + 1],
          "\n".join([
              f"TOTAL DE IMPACTOS: {metricas['positiva']}",
              f"ENTREVISTAS: {positivos['ENTREVISTAS']}",
              f"TV: {positivos['TELEVISIÓN']}",
              f"RADIO: {positivos['RADIO']}",
              f"PRENSA LOCAL: {positivos['PRENSA LOCAL']}",
              f"PORTALES DIGITALES: {positivos['PORTALES DIGITALES']}",
              f"COLUMNAS: {positivos['COLUMNAS']}",
              f"REDES SOCIALES: {positivos['REDES SOCIALES']}",
          ]),
      )

    _reemplazar_texto_con_formato(
        parrafos[indice_neg],
        f"TOTAL NOTAS NEGATIVAS: {metricas['negativa']}",
    )
    negativos = metricas["negativos_canal"]
    detalles_negativos = [
        f"TOTAL DE IMPACTOS: {metricas['negativa']}",
        "\n".join([
            f"ENTREVISTAS: {negativos['ENTREVISTAS']}",
            f"TV: {negativos['TELEVISIÓN']}",
            f"RADIO: {negativos['RADIO']}",
            f"PRENSA LOCAL: {negativos['PRENSA LOCAL']}",
            f"PORTALES DIGITALES: {negativos['PORTALES DIGITALES']}",
            f"COLUMNAS: {negativos['COLUMNAS']}",
        ]),
        f"REDES SOCIALES: {negativos['REDES SOCIALES']}",
    ]
    for desplazamiento, texto in enumerate(detalles_negativos, 1):
      indice = indice_neg + desplazamiento
      if indice < len(parrafos):
        _reemplazar_texto_con_formato(parrafos[indice], texto)
        parrafos[indice].alignment = WD_ALIGN_PARAGRAPH.LEFT
    return

  # Compatibilidad con plantillas antiguas que tenían un solo bloque.
  canales_totales = {
      canal: metricas["positivos_canal"][canal]
      + metricas["negativos_canal"][canal]
      for canal in ORDEN_CANALES_UNIFICADOS
  }
  lineas = [
      f"TOTAL NOTAS INFORMATIVAS: {metricas['positiva']}",
      f"TOTAL NOTAS NEGATIVAS: {metricas['negativa']}",
      f"TOTAL DE IMPACTOS: {metricas['total']}",
      f"ENTREVISTAS: {canales_totales['ENTREVISTAS']}",
      f"TV: {canales_totales['TELEVISIÓN']}",
      f"RADIO: {canales_totales['RADIO']}",
      f"PRENSA LOCAL: {canales_totales['PRENSA LOCAL']}",
      f"PORTALES DIGITALES: {canales_totales['PORTALES DIGITALES']}",
      f"COLUMNAS: {canales_totales['COLUMNAS']}",
      f"REDES SOCIALES: {canales_totales['REDES SOCIALES']}",
  ]
  if indice_info is not None:
    _reemplazar_texto_con_formato(parrafos[indice_info], "\n".join(lineas))
    return
  raise ValueError(
      "La plantilla no contiene el bloque de totales del reporte."
  )


def _texto_elemento_docx(elemento):
  return "".join(
      nodo.text or "" for nodo in elemento.iter(qn("w:t"))
  ).strip()


def _analizar_bloque_fecha(bloque):
  """Separa un día en nodos informativos/negativos y calcula sus cifras."""
  resultado = {
      "informativos": [],
      "negativos": [],
      "total_informativos": 0,
      "total_negativos": 0,
      "canales_informativos": {
          canal: 0 for canal in ORDEN_CANALES_UNIFICADOS
      },
      "canales_negativos": {
          canal: 0 for canal in ORDEN_CANALES_UNIFICADOS
      },
  }
  estado = "POSITIVA"
  total_info_explicito = None
  total_neg_explicito = None

  for indice, nodo in enumerate(bloque):
    texto = _texto_normalizado_docx(_texto_elemento_docx(nodo))
    texto_norm = quitar_acentos(texto).upper()
    if indice == 0 and PATRON_FECHA_REPORTE.match(texto):
      continue
    if texto_norm.startswith("TOTAL DE IMPACTOS INFORMATIVOS"):
      estado = "POSITIVA"
      valor = _extraer_entero_docx(texto.split(":", 1)[-1])
      if valor is not None:
        total_info_explicito = valor
      continue
    if (
        texto_norm.startswith("TOTAL DE IMPACTOS NEGATIVOS")
        or texto_norm.startswith("TOTAL NOTAS NEGATIVAS")
        or texto_norm.startswith("NEGATIVAS:")
    ):
      estado = "NEGATIVA"
      valor = _extraer_entero_docx(texto.split(":", 1)[-1])
      if valor is not None:
        total_neg_explicito = valor
      continue

    conteos = _extraer_conteos_canales(texto)
    if conteos:
      estado_nodo = estado
      for conteo in conteos:
        modificador = conteo["modificador"]
        if modificador.startswith("NEGATIVA"):
          sentimiento_conteo = "NEGATIVA"
        elif modificador.startswith("INFORMATIVA"):
          sentimiento_conteo = "POSITIVA"
        else:
          sentimiento_conteo = estado
        estado_nodo = sentimiento_conteo
        canal = conteo["canal"]
        cantidad = conteo["cantidad"]
        if canal and cantidad is not None:
          clave = (
              "canales_negativos"
              if sentimiento_conteo == "NEGATIVA"
              else "canales_informativos"
          )
          resultado[clave][canal] += cantidad
      destino = (
          resultado["negativos"]
          if estado_nodo == "NEGATIVA"
          else resultado["informativos"]
      )
      destino.append(nodo)
      continue

    resultado[
        "negativos" if estado == "NEGATIVA" else "informativos"
    ].append(nodo)

  calculado_info = sum(resultado["canales_informativos"].values())
  calculado_neg = sum(resultado["canales_negativos"].values())
  resultado["total_informativos"] = max(
      total_info_explicito or 0, calculado_info
  )
  resultado["total_negativos"] = max(
      total_neg_explicito or 0, calculado_neg
  )
  return resultado


def _es_encabezado_canal_elemento(nodo):
  return bool(_extraer_conteos_canales(_texto_elemento_docx(nodo)))


def _rpr_tema_modelo(parrafo_modelo, negrita):
  candidato = None
  for run_xml in parrafo_modelo._p.findall(qn("w:r")):
    rpr = run_xml.find(qn("w:rPr"))
    if rpr is None:
      continue
    es_negrita = rpr.find(qn("w:b")) is not None
    if es_negrita == negrita:
      return deepcopy(rpr)
    if candidato is None:
      candidato = deepcopy(rpr)

  rpr = candidato if candidato is not None else OxmlElement("w:rPr")
  for etiqueta in ("w:b", "w:bCs"):
    nodo = rpr.find(qn(etiqueta))
    if negrita and nodo is None:
      nodo = OxmlElement(etiqueta)
      nodo.set(qn("w:val"), "1")
      rpr.append(nodo)
    elif not negrita and nodo is not None:
      rpr.remove(nodo)
  return rpr


def _establecer_tema_en_elemento(p_xml, texto, parrafo_modelo):
  """Escribe un tema con su encabezado en negritas y explicación normal."""
  for hijo in list(p_xml):
    if hijo.tag != qn("w:pPr"):
      p_xml.remove(hijo)

  coincidencia = re.match(r"^(.*?:)(\s*)(.*)$", texto.strip())
  if coincidencia:
    titulo = coincidencia.group(1)
    separador = coincidencia.group(2) or " "
    cuerpo = coincidencia.group(3)
  else:
    titulo, separador, cuerpo = texto.strip(), "", ""

  for contenido, negrita in [
      (titulo + separador, True),
      (cuerpo, False),
  ]:
    if not contenido:
      continue
    run_xml = OxmlElement("w:r")
    run_xml.append(_rpr_tema_modelo(parrafo_modelo, negrita))
    texto_xml = OxmlElement("w:t")
    texto_xml.set(qn("xml:space"), "preserve")
    texto_xml.text = contenido
    run_xml.append(texto_xml)
    p_xml.append(run_xml)


def _rellenar_resumen_informativo_si_vacio(doc_base, doc_redes):
  """Completa los campos 1.-, 2.-, 3.- cuando la plantilla no trae temas."""
  informativos_base, _ = _extraer_temas_docx(doc_base)
  if informativos_base:
    return
  informativos_redes, _ = _extraer_temas_docx(doc_redes)
  informativos = _deduplicar_temas(informativos_redes)
  if not informativos:
    return

  parrafos = doc_base.paragraphs
  indice_info = next(
      (
          i for i, p in enumerate(parrafos)
          if _titulo_seccion_temas_docx(p.text) == "INFORMATIVOS"
      ),
      None,
  )
  indice_neg = next(
      (
          i for i, p in enumerate(parrafos)
          if _titulo_seccion_temas_docx(p.text) == "NEGATIVOS"
      ),
      None,
  )
  if indice_info is None or indice_neg is None or indice_info >= indice_neg:
    return

  espacios = parrafos[indice_info + 1:indice_neg]
  if not espacios:
    return
  modelo_tema = next(
      (p for p in espacios if _texto_normalizado_docx(p.text)),
      parrafos[indice_info],
  )
  for numero, tema in enumerate(informativos, 1):
    texto_tema = f"{numero}. {tema}"
    if numero <= len(espacios):
      _establecer_tema_en_elemento(
          espacios[numero - 1]._p, texto_tema, modelo_tema
      )
      continue
    clon = deepcopy(modelo_tema._p)
    _establecer_tema_en_elemento(clon, texto_tema, modelo_tema)
    parrafos[indice_neg]._p.addprevious(clon)

  for parrafo in espacios[len(informativos):]:
    texto = _texto_normalizado_docx(parrafo.text)
    if re.fullmatch(r"\d+\s*[.)]?\s*-?", texto):
      _reemplazar_texto_con_formato(parrafo, "")


def _reconstruir_resumen_negativo_unificado(doc_base, doc_redes):
  """Integra en la plantilla los temas negativos de ambos reportes."""
  _, negativos_base = _extraer_temas_docx(doc_base)
  _, negativos_redes = _extraer_temas_docx(doc_redes)
  negativos = _deduplicar_temas(negativos_base + negativos_redes)
  if not negativos:
    negativos = [
        "Sin incidencias negativas: No se registraron temas negativos en el"
        " periodo analizado."
    ]

  parrafos = doc_base.paragraphs
  indice_info = next(
      (
          i for i, p in enumerate(parrafos)
          if _titulo_seccion_temas_docx(p.text) == "INFORMATIVOS"
      ),
      None,
  )
  indice_neg = next(
      (
          i for i, p in enumerate(parrafos)
          if _titulo_seccion_temas_docx(p.text) == "NEGATIVOS"
      ),
      None,
  )
  indice_desglose = _buscar_indice_desglose(doc_base)
  if indice_neg is None or indice_desglose is None:
    raise ValueError(
        "La plantilla no contiene el encabezado de temas negativos o la"
        " sección DESGLOSE."
    )

  modelo_tema = next(
      (
          p for p in parrafos[indice_neg + 1:indice_desglose]
          if _texto_normalizado_docx(p.text)
      ),
      None,
  )
  if modelo_tema is None:
    if indice_info is not None:
      modelo_tema = next(
          (
              p for p in parrafos[indice_info + 1:indice_neg]
              if _texto_normalizado_docx(p.text)
          ),
          parrafos[indice_neg],
      )
    else:
      modelo_tema = parrafos[indice_neg]

  espacios = parrafos[indice_neg + 1:indice_desglose]
  for numero, tema in enumerate(negativos, 1):
    texto_tema = f"{numero}. {tema}"
    if numero <= len(espacios):
      _establecer_tema_en_elemento(
          espacios[numero - 1]._p, texto_tema, modelo_tema
      )
      continue
    clon = deepcopy(modelo_tema._p)
    _establecer_tema_en_elemento(clon, texto_tema, modelo_tema)
    parrafos[indice_desglose]._p.addprevious(clon)

  # Si la plantilla ya tenía más temas, se limpian los sobrantes sin eliminar
  # los párrafos de reserva que sostienen su distribución visual.
  for parrafo in espacios[len(negativos):]:
    if _texto_normalizado_docx(parrafo.text):
      _reemplazar_texto_con_formato(parrafo, "")


def _reconstruir_resumen_unificado(doc_base, doc_tradicional):
  info_base, neg_base = _extraer_temas_docx(doc_base)
  info_trad, neg_trad = _extraer_temas_docx(doc_tradicional)
  informativos = _deduplicar_temas(info_base + info_trad)
  negativos = _deduplicar_temas(neg_base + neg_trad)

  parrafos = doc_base.paragraphs
  indice_info = next(
      (
          i
          for i, p in enumerate(parrafos)
          if _titulo_seccion_temas_docx(p.text) == "INFORMATIVOS"
      ),
      None,
  )
  indice_neg = next(
      (
          i
          for i, p in enumerate(parrafos)
          if _titulo_seccion_temas_docx(p.text) == "NEGATIVOS"
      ),
      None,
  )
  indice_desglose = _buscar_indice_desglose(doc_base)
  if None in (indice_info, indice_neg, indice_desglose):
    return

  modelo_info = (
      parrafos[indice_info + 1]
      if indice_info + 1 < indice_neg
      else parrafos[indice_info]
  )
  modelo_neg = (
      parrafos[indice_neg + 1]
      if indice_neg + 1 < indice_desglose
      else modelo_info
  )
  nodo_info = parrafos[indice_info]._p
  nodo_neg = parrafos[indice_neg]._p
  nodo_desglose = parrafos[indice_desglose]._p
  cuerpo = doc_base._body._element

  elementos = list(cuerpo)
  inicio = elementos.index(nodo_info)
  fin = elementos.index(nodo_neg)
  for elemento in elementos[inicio + 1:fin]:
    cuerpo.remove(elemento)
  posicion = list(cuerpo).index(nodo_neg)
  for numero, tema in enumerate(informativos, 1):
    cuerpo.insert(
        posicion,
        _clonar_parrafo_con_texto(modelo_info, f"{numero}. {tema}"),
    )
    posicion += 1

  elementos = list(cuerpo)
  inicio = elementos.index(nodo_neg)
  fin = elementos.index(nodo_desglose)
  for elemento in elementos[inicio + 1:fin]:
    cuerpo.remove(elemento)
  posicion = list(cuerpo).index(nodo_desglose)
  if negativos:
    for numero, tema in enumerate(negativos, 1):
      cuerpo.insert(
          posicion,
          _clonar_parrafo_con_texto(modelo_neg, f"{numero}. {tema}"),
      )
      posicion += 1
  else:
    cuerpo.insert(
        posicion,
        _clonar_parrafo_con_texto(
            modelo_neg, "No se registraron temas negativos en el periodo."
        ),
    )


def _reconstruir_desglose_unificado(
    doc_base,
    doc_redes,
    bloques_tradicionales,
    bloques_redes,
):
  indice_desglose = _buscar_indice_desglose(doc_base)
  if indice_desglose is None:
    raise ValueError("La plantilla no contiene la sección DESGLOSE.")

  parrafos = doc_base.paragraphs
  modelo_fecha = next(
      (
          p for p in parrafos[indice_desglose + 1:]
          if PATRON_FECHA_REPORTE.match(_texto_normalizado_docx(p.text))
      ),
      None,
  )
  modelo_total_info = next(
      (
          p for p in parrafos[indice_desglose + 1:]
          if quitar_acentos(_texto_normalizado_docx(p.text)).upper().startswith(
              "TOTAL DE IMPACTOS INFORMATIVOS"
          )
      ),
      None,
  )
  modelo_canal = next(
      (
          p for p in parrafos[indice_desglose + 1:]
          if _extraer_conteos_canales(_texto_normalizado_docx(p.text))
      ),
      None,
  )
  modelo_total_neg = next(
      (
          p for p in parrafos[:indice_desglose]
          if "TOTAL NOTAS NEGATIVAS:" in p.text.upper()
      ),
      modelo_total_info,
  )
  if None in (modelo_fecha, modelo_total_info, modelo_canal, modelo_total_neg):
    raise ValueError(
        "La plantilla no contiene los estilos de fecha, total y canal"
        " necesarios para construir el desglose."
    )

  nodo_desglose = doc_base.paragraphs[indice_desglose]._p
  cuerpo = doc_base._body._element
  elementos = list(cuerpo)
  posicion_desglose = elementos.index(nodo_desglose)
  for elemento in elementos[posicion_desglose + 1:]:
    if elemento.tag != qn("w:sectPr"):
      cuerpo.remove(elemento)

  fechas = sorted(
      set(bloques_tradicionales) | set(bloques_redes),
      key=_fecha_desde_etiqueta,
  )
  sect_pr = cuerpo.sectPr
  posicion = list(cuerpo).index(sect_pr) if sect_pr is not None else len(cuerpo)

  for fecha in fechas:
    bloque_trad_original = bloques_tradicionales.get(fecha, [])
    bloque_redes_original = bloques_redes.get(fecha, [])
    analisis_trad = _analizar_bloque_fecha(
        bloque_trad_original
    )
    analisis_redes = _analizar_bloque_fecha(bloque_redes_original)
    total_info = (
        analisis_trad["total_informativos"]
        + analisis_redes["total_informativos"]
    )
    total_neg = (
        analisis_trad["total_negativos"]
        + analisis_redes["total_negativos"]
    )

    # Si la fecha ya aparece en tradicionales se conserva exactamente como fue
    # escrita allí (incluido año de dos o cuatro dígitos).
    nodo_fecha_origen = (
        bloque_trad_original[0]
        if bloque_trad_original
        else bloque_redes_original[0]
    )
    fecha_visible = _texto_normalizado_docx(
        _texto_elemento_docx(nodo_fecha_origen)
    )
    cuerpo.insert(
        posicion,
        _clonar_parrafo_con_texto(modelo_fecha, fecha_visible),
    )
    posicion += 1
    cuerpo.insert(
        posicion,
        _clonar_parrafo_con_texto(
            modelo_total_info,
            f"TOTAL DE IMPACTOS INFORMATIVOS: {total_info}",
        ),
    )
    posicion += 1

    # Orden institucional: medios tradicionales primero y redes al final.
    for nodo in analisis_trad["informativos"]:
      cuerpo.insert(
          posicion,
          _copiar_elemento_con_relaciones(nodo, doc_base, doc_base),
      )
      posicion += 1

    total_redes_info = analisis_redes["canales_informativos"][
        "REDES SOCIALES"
    ]
    if total_redes_info:
      cuerpo.insert(
          posicion,
          _clonar_parrafo_con_texto(
              modelo_canal, f"REDES SOCIALES: ({total_redes_info})"
          ),
      )
      posicion += 1
    for nodo in analisis_redes["informativos"]:
      if _es_encabezado_canal_elemento(nodo):
        continue
      cuerpo.insert(
          posicion,
          _copiar_elemento_con_relaciones(nodo, doc_redes, doc_base),
      )
      posicion += 1

    # El total negativo se imprime siempre, incluso si es cero, para que cada
    # fecha quede completa y sea auditable.
    cuerpo.insert(
        posicion,
        _clonar_parrafo_con_texto(
            modelo_total_neg,
            f"TOTAL DE IMPACTOS NEGATIVOS: {total_neg}",
        ),
    )
    posicion += 1

    for nodo in analisis_trad["negativos"]:
      cuerpo.insert(
          posicion,
          _copiar_elemento_con_relaciones(nodo, doc_base, doc_base),
      )
      posicion += 1

    total_redes_neg = analisis_redes["total_negativos"]
    if total_redes_neg:
      cuerpo.insert(
          posicion,
          _clonar_parrafo_con_texto(
              modelo_canal, f"REDES SOCIALES: ({total_redes_neg})"
          ),
      )
      posicion += 1
    for nodo in analisis_redes["negativos"]:
      if _es_encabezado_canal_elemento(nodo):
        continue
      cuerpo.insert(
          posicion,
          _copiar_elemento_con_relaciones(nodo, doc_redes, doc_base),
      )
      posicion += 1


def _obtener_actor_reporte_docx(doc):
  """Obtiene una etiqueta breve del actor para evitar mezclas accidentales."""
  etiquetas_omitidas = (
      "PERIODO DE MEDICIÓN:",
      "CANALES:",
      "BALANCE DE IMPACTOS",
      "RESUMEN",
  )
  for parrafo in doc.paragraphs[:15]:
    texto = _texto_normalizado_docx(parrafo.text)
    if texto.upper() == "RESUMEN":
      break
    if texto and not texto.upper().startswith(etiquetas_omitidas):
      return texto
  for tabla in doc.tables:
    for fila in tabla.rows[:2]:
      for celda in fila.cells:
        texto = _texto_normalizado_docx(celda.text)
        texto_mayus = texto.upper()
        if (
            texto
            and not texto_mayus.startswith(etiquetas_omitidas)
            and "POSITIVA" not in texto_mayus
            and "NEGATIVA" not in texto_mayus
            and "TOTAL DE IMPACTOS" not in texto_mayus
        ):
          return texto
  return ""


def _actores_compatibles_docx(actor_uno, actor_dos):
  ignorar = {
      "MORENA", "PARTIDO", "CANDIDATO", "CANDIDATA", "REPORTE",
      "PAN", "PRI", "PRD", "MOVIMIENTO", "CIUDADANO",
  }
  tokens_uno = {
      t for t in re.findall(r"[A-Z]{4,}", quitar_acentos(actor_uno).upper())
      if t not in ignorar
  }
  tokens_dos = {
      t for t in re.findall(r"[A-Z]{4,}", quitar_acentos(actor_dos).upper())
      if t not in ignorar
  }
  return not tokens_uno or not tokens_dos or bool(tokens_uno & tokens_dos)


def _validar_plantilla_visual_docx(doc):
  """Valida únicamente que el primer Word pueda actuar como plantilla."""
  textos = [
      quitar_acentos(_texto_normalizado_docx(p.text)).upper()
      for p in doc.paragraphs
  ]
  faltantes = []
  if "RESUMEN" not in textos:
    faltantes.append("RESUMEN")
  if not any("TOTAL NOTAS INFORMATIVAS:" in texto for texto in textos):
    faltantes.append("TOTAL NOTAS INFORMATIVAS")
  if not any("TOTAL NOTAS NEGATIVAS:" in texto for texto in textos):
    faltantes.append("TOTAL NOTAS NEGATIVAS")
  if "DESGLOSE" not in textos:
    faltantes.append("DESGLOSE")

  tiene_balance = False
  for tabla in doc.tables:
    contenido = " | ".join(
        quitar_acentos(_texto_normalizado_docx(celda.text)).upper()
        for fila in tabla.rows
        for celda in fila.cells
    )
    if "POSITIVA" in contenido and "NEGATIVA" in contenido:
      tiene_balance = True
      break
  if not tiene_balance:
    faltantes.append("tabla de BALANCE DE IMPACTOS")

  if faltantes:
    raise ValueError(
        "El primer archivo no contiene el entorno visual requerido: "
        + ", ".join(faltantes)
        + "."
    )


def unificar_reportes_word(archivo_tradicional, archivo_redes):
  """
  Genera un tercer Word unificado.

  El archivo tradicional es la plantilla visual: conserva sus estilos, tamaño
  de página, márgenes, encabezados, pies, tabla, colores e imágenes. Del archivo
  de redes se incorporan cifras y publicaciones. El periodo y los encabezados
  de fecha del archivo tradicional se conservan sin cambios.
  """
  doc_base = _abrir_docx_desde_streamlit(archivo_tradicional)
  doc_redes = _abrir_docx_desde_streamlit(archivo_redes)
  _validar_plantilla_visual_docx(doc_base)

  bloques_tradicionales = _extraer_bloques_por_fecha(doc_base)
  bloques_redes = _extraer_bloques_por_fecha(doc_redes)
  if not bloques_tradicionales:
    raise ValueError(
        "El primer archivo debe contener la sección DESGLOSE con al menos una"
        " fecha para funcionar como plantilla."
    )
  if not bloques_redes:
    raise ValueError(
        "El segundo archivo no contiene publicaciones fechadas dentro de la"
        " sección DESGLOSE."
    )

  metricas_trad = _obtener_metricas_docx(
      doc_base, es_reporte_redes=False
  )
  metricas_redes = _obtener_metricas_docx(doc_redes, es_reporte_redes=True)
  metricas = {
      "positiva": metricas_trad["positiva"] + metricas_redes["positiva"],
      "negativa": metricas_trad["negativa"] + metricas_redes["negativa"],
  }
  metricas["total"] = metricas["positiva"] + metricas["negativa"]
  metricas["positivos_canal"] = {
      canal: metricas_trad["positivos_canal"][canal]
      + metricas_redes["positivos_canal"][canal]
      for canal in ORDEN_CANALES_UNIFICADOS
  }
  metricas["negativos_canal"] = {
      canal: metricas_trad["negativos_canal"][canal]
      + metricas_redes["negativos_canal"][canal]
      for canal in ORDEN_CANALES_UNIFICADOS
  }

  _actualizar_balance_unificado(
      doc_base,
      metricas["positiva"],
      metricas["negativa"],
      metricas["total"],
  )
  _actualizar_totales_unificados(doc_base, metricas)
  # El periodo de medición de la portada pertenece a la plantilla tradicional
  # y no se reemplaza con el rango combinado.
  _rellenar_resumen_informativo_si_vacio(doc_base, doc_redes)
  _reconstruir_resumen_negativo_unificado(doc_base, doc_redes)
  _reconstruir_desglose_unificado(
      doc_base,
      doc_redes,
      bloques_tradicionales,
      bloques_redes,
  )

  salida = io.BytesIO()
  aplicar_verdana_12_documento(doc_base)
  doc_base.save(salida)
  salida.seek(0)
  return salida


def _extraer_temas_para_modelo(texto_temas):
  """Convierte el resumen devuelto por Gemini en dos listas de temas."""
  informativos = []
  negativos = []
  seccion = "INFORMATIVOS"
  for linea in str(texto_temas or "").splitlines():
    limpia = linea.strip()
    if not limpia:
      continue
    encabezado = quitar_acentos(limpia).upper()
    if "TEMAS RELEVANTES INFORMATIVOS" in encabezado:
      seccion = "INFORMATIVOS"
      continue
    if "TEMAS NEGATIVOS" in encabezado:
      seccion = "NEGATIVOS"
      continue
    limpia = re.sub(r"^\s*\d+\s*[.)-]\s*", "", limpia).strip()
    if not limpia:
      continue
    destino = negativos if seccion == "NEGATIVOS" else informativos
    destino.append(limpia)
  return informativos[:3], negativos[:3]


def _obtener_rango_fechas_dataframe(df_data):
  serie = obtener_columna_serie(
      df_data,
      [
          "Publish date", "Fecha", "Date", "Fecha de publicación",
          "Fecha de publicacion",
      ],
  ).apply(parsear_fecha_perfecta)
  serie = pd.to_datetime(serie, errors="coerce").dropna()
  if serie.empty:
    hoy = pd.Timestamp.today().normalize()
    return hoy.date(), hoy.date()
  return serie.min().date(), serie.max().date()


def construir_modelo_google_docs(
    df_hoja,
    nombre_actor,
    es_redes_sociales,
    fecha_inicio,
    fecha_fin,
):
  """Procesa una extracción y devuelve datos independientes del formato."""
  if df_hoja is None or df_hoja.empty:
    raise ValueError("El archivo no contiene filas para procesar.")
  df_trabajo = reparar_desfase_columnas_excel(df_hoja.copy())

  mask_sin_notas = df_trabajo.apply(
      lambda row: any(
          palabra in str(valor).lower()
          for valor in row.values
          for palabra in ["sin notas", "sin nota", "sin registro"]
      ),
      axis=1,
  )
  df_trabajo = df_trabajo[~mask_sin_notas].copy()
  if es_redes_sociales:
    df_trabajo, total_descartadas = limpiar_dataframe_redes_automatico(
        df_trabajo, nombre_actor
    )
    if total_descartadas:
      st.info(
          f"Se descartaron {int(total_descartadas)} publicaciones de cuentas"
          " propias o institucionales."
      )
  if df_trabajo.empty:
    raise ValueError("No quedaron notas válidas después de la limpieza.")

  serie_fechas = obtener_columna_serie(
      df_trabajo,
      [
          "Publish date", "Fecha", "Date", "Fecha de publicación",
          "Fecha de publicacion",
      ],
  )
  df_trabajo["fecha_dt"] = pd.to_datetime(
      serie_fechas.apply(parsear_fecha_perfecta), errors="coerce"
  )
  df_trabajo = df_trabajo.dropna(subset=["fecha_dt"])
  inicio = pd.Timestamp(fecha_inicio).normalize()
  fin = pd.Timestamp(fecha_fin).normalize()
  if inicio > fin:
    inicio, fin = fin, inicio
  df_trabajo = df_trabajo[
      (df_trabajo["fecha_dt"] >= inicio)
      & (df_trabajo["fecha_dt"] <= fin + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1))
  ].copy()
  if df_trabajo.empty:
    raise ValueError(
        "No hay notas dentro del periodo seleccionado. Revisa las fechas."
    )

  # Solo elimina una misma fila recargada; publicaciones semejantes de medios
  # distintos continúan contando como impactos separados.
  subset_dup = [
      col for col in ["ID Nota", "Link de Nota", "Link URL Medio", "URL"]
      if col in df_trabajo.columns
  ]
  if subset_dup:
    df_trabajo = df_trabajo.drop_duplicates(subset=subset_dup)
  df_trabajo = df_trabajo.sort_values("fecha_dt").reset_index(drop=True)

  resultados = determinar_sentimiento_df(
      df_trabajo,
      nombre_actor,
      es_tradicionales=not es_redes_sociales,
  )
  for columna in resultados.columns:
    df_trabajo[columna] = resultados[columna].values
  df_trabajo = df_trabajo[df_trabajo["relevante_ia"]].copy()
  if df_trabajo.empty:
    raise ValueError("La IA determinó que ninguna publicación correspondía al actor.")

  medios = obtener_columna_serie(
      df_trabajo,
      [
          "Tipo de Medio", "Fuente", "Media type", "Media Type", "Medio",
          "Nombre del Medio", "Canal", "Tipo de Nota",
      ],
  )
  df_trabajo["categoria_medio_std"] = (
      "REDES SOCIALES"
      if es_redes_sociales
      else medios.apply(estandarizar_categoria_medio)
  )

  notas = []
  for _, row in df_trabajo.iterrows():
    if es_redes_sociales:
      autor = obtener_campo(
          row,
          ["Autor", "Author name", "Fuente", "Media name", "Programa"],
      )
      handle = obtener_campo(
          row,
          ["Author handle (@username)", "Handle", "Username"],
      )
      if handle and not handle.startswith("@"):
        handle = "@" + handle
      fuente = " ".join(parte for parte in [autor, handle] if parte).strip()
      contenido = obtener_campo(
          row,
          [
              "Contenido", "Detail", "Summary", "Síntesis", "Sintesis",
              "Titulo", "Título", "Title", "Encabezado",
          ],
      )
    else:
      medio = obtener_campo(
          row,
          ["Nombre del Medio", "Fuente", "Media name", "Medio"],
      )
      autor = obtener_campo(
          row,
          ["Autor", "Author name", "Programa", "Conductor"],
      )
      fuente = (
          f"{medio} - {autor}"
          if autor and autor not in ["Redacción", "Staff", "Online", medio]
          else medio
      )
      contenido = obtener_campo(
          row,
          [
              "Titulo", "Título", "Contenido", "Detail", "Summary",
              "Síntesis", "Sintesis", "Encabezado", "Nota",
          ],
      )
      hora = obtener_campo(
          row,
          ["Hora", "Hour", "Time", "Hora de Transmisión", "Hora de Transmision"],
      )
      if hora and contenido and not contenido.lower().startswith(hora.lower()):
        contenido = f"{hora} {contenido}"

    notas.append({
        "fecha": pd.Timestamp(row["fecha_dt"]).date().isoformat(),
        "canal": row["categoria_medio_std"],
        "sentimiento": row["sentimiento_final"],
        "fuente": fuente or "Sin fuente",
        "contenido": re.sub(r"\s+", " ", limpiar_texto(contenido)).strip(),
        "url": obtener_link_inteligente(row),
        "origen": "REDES" if es_redes_sociales else "TRADICIONALES",
    })

  resumen = extraer_resumen_temas_real(df_trabajo, nombre_actor)
  informativos, negativos = _extraer_temas_para_modelo(resumen)
  return {
      "actor": nombre_actor,
      "fecha_inicio": inicio.date().isoformat(),
      "fecha_fin": fin.date().isoformat(),
      "temas_informativos": informativos,
      "temas_negativos": negativos,
      "notas": notas,
  }


def _agrupar_archivos_por_actor(archivos):
  candidatos = {}
  for archivo in archivos:
    hojas = cargar_archivo_seguro(archivo)
    for nombre_hoja, dataframe in hojas.items():
      if dataframe is None or dataframe.empty:
        continue
      if "Menu" in dataframe.columns and len(dataframe["Menu"].dropna()) > 0:
        nombre_raw = str(dataframe["Menu"].dropna().iloc[0]).strip()
      else:
        nombre_raw = nombre_hoja
      actor = normalizar_nombre_candidato(nombre_raw)
      candidatos.setdefault(actor, []).append(dataframe)
  return candidatos


# --- INTERFAZ STREAMLIT ---

tipo_analisis = st.radio(
    "¿Qué tipo de archivo vas a analizar?",
    [
        "Redes Sociales",
        "Medios Tradicionales / Portales / TV y Radio (Multi-Archivo)",
        "Unificar reportes Word (Tradicionales + Redes Sociales)",
        "Google Docs automático (Primera extracción + Redes)",
    ],
    index=0,
)

if tipo_analisis == "Redes Sociales":
  uploaded_file = st.file_uploader(
      "Sube tu archivo Excel o CSV de Redes Sociales",
      type=["xlsx", "xls", "csv"],
  )
  actor_nombre_in = st.text_input(
      "Nombre y Cargo / Partido del Actor Político",
      placeholder=(
          "ej. GABRIELA SÁNCHEZ SAAVEDRA, ALEJANDRO ARMENTA, PEPE CHEDRAUI, etc."
      ),
  ).strip().upper()

  if uploaded_file and actor_nombre_in:
    if st.button("Generar Reporte Oficial", type="primary"):
      with st.spinner(
          "Limpiando cuentas oficiales y evaluando sentimiento con IA desde la"
          " perspectiva del actor..."
      ):
        try:
          dict_h = cargar_archivo_seguro(uploaded_file)
          df_redes = list(dict_h.values())[0]
          buf = crear_doc_desde_hoja(
              df_redes, actor_nombre_in, es_redes_sociales=True
          )
          if buf is not None:
            st.success(f"¡Reporte generado exitosamente para '{actor_nombre_in}'!")
            st.download_button(
                label=f"📥 Descargar Reporte Word de {actor_nombre_in}",
                data=buf,
                file_name=f"Reporte_{actor_nombre_in.replace(' ', '_')}.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                ),
            )
          else:
            st.warning(
                "El archivo no contiene notas válidas tras la limpieza"
                " automática."
            )
        except Exception as e:
          st.error(f"Error procesando el archivo: {str(e)}")

elif tipo_analisis == (
    "Medios Tradicionales / Portales / TV y Radio (Multi-Archivo)"
):
  uploaded_files = st.file_uploader(
      "Sube uno o varios archivos Excel (ej. Archivo de TV/Radio y Archivo de"
      " Portales Web)",
      type=["xlsx", "xls", "csv"],
      accept_multiple_files=True,
  )

  if uploaded_files and len(uploaded_files) > 0:
    candidatos_unificados = {}

    for file_item in uploaded_files:
      try:
        dict_hojas = cargar_archivo_seguro(file_item)
        for h_name, df_h in dict_hojas.items():
          if df_h is None or df_h.empty:
            continue

          if "Menu" in df_h.columns and len(df_h["Menu"].dropna()) > 0:
            nombre_raw = str(df_h["Menu"].dropna().iloc[0]).strip()
          else:
            nombre_raw = h_name

          candidato_canon = normalizar_nombre_candidato(nombre_raw)

          if candidato_canon not in candidatos_unificados:
            candidatos_unificados[candidato_canon] = []
          candidatos_unificados[candidato_canon].append(df_h)
      except Exception as e:
        st.warning(
            "No se pudo procesar una de las hojas de"
            f" {file_item.name}: {str(e)}"
        )

    candidatos_disponibles = sorted(
        [c for c in candidatos_unificados.keys() if c and "SIN NOTAS" not in c]
    )

    if len(candidatos_disponibles) > 0:
      st.success(
          f"✅ Se cargaron exitosamente {len(uploaded_files)} archivo(s) y se"
          f" detectaron {len(candidatos_disponibles)} candidatos."
      )
      st.markdown(
          f"**Candidatos detectados:** {', '.join(candidatos_disponibles)}"
      )

      st.write("---")
      st.subheader("Opciones de Descarga:")

      col1, col2 = st.columns(2)

      with col1:
        st.markdown("### 📦 Descarga Masiva")
        if st.button(
            "Generar y Descargar TODOS los Reportes en .ZIP",
            type="primary",
            use_container_width=True,
        ):
          with st.spinner(
              "Generando reportes consolidados para todos los candidatos..."
          ):
            zip_buffer = io.BytesIO()
            cnt_generados = 0

            with zipfile.ZipFile(
                zip_buffer, "w", zipfile.ZIP_DEFLATED
            ) as zip_file:
              for cand_name in candidatos_disponibles:
                lista_dfs = candidatos_unificados[cand_name]
                df_total_candidato = pd.concat(lista_dfs, ignore_index=True)

                buf = crear_doc_desde_hoja(
                    df_total_candidato, cand_name, es_redes_sociales=False
                )
                if buf is not None:
                  doc_bytes = buf.getvalue()
                  fname = f"Reporte_{cand_name.replace(' ', '_')}.docx"
                  zip_file.writestr(fname, doc_bytes)
                  cnt_generados += 1

            zip_buffer.seek(0)
            if cnt_generados > 0:
              st.success(
                  f"¡Se generaron con éxito {cnt_generados} reportes"
                  " consolidados!"
              )
              st.download_button(
                  label="📥 Descargar Archivo .ZIP",
                  data=zip_buffer,
                  file_name="Reportes_Monitoreo_Consolidados_Completos.zip",
                  mime="application/zip",
                  use_container_width=True,
              )
            else:
              st.warning(
                "No se encontraron candidatos con notas activas."
            )

      with col2:
        st.markdown("### 📄 Descarga Individual")
        cand_sel = st.selectbox(
            "Selecciona un candidato:", candidatos_disponibles
        )
        if st.button(
            f"Generar Reporte de {cand_sel}", use_container_width=True
        ):
          with st.spinner(f"Consolidando notas para {cand_sel}..."):
            lista_dfs = candidatos_unificados[cand_sel]
            df_total_candidato = pd.concat(lista_dfs, ignore_index=True)

            buf = crear_doc_desde_hoja(
                df_total_candidato, cand_sel, es_redes_sociales=False
            )
            if buf is not None:
              st.success(f"¡Reporte consolidado listo para '{cand_sel}'!")
              st.download_button(
                  label=f"📥 Descargar Word de {cand_sel}",
                  data=buf,
                  file_name=f"Reporte_{cand_sel.replace(' ', '_')}.docx",
                  mime=(
                      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  ),
                  use_container_width=True,
              )
            else:
              st.warning(
                  f"El candidato '{cand_sel}' no contiene notas registradas."
              )
    else:
      st.warning(
          "No se detectaron candidatos con notas válidas en los archivos"
          " seleccionados."
      )

elif tipo_analisis == "Unificar reportes Word (Tradicionales + Redes Sociales)":
  st.subheader("🔗 Unificar dos reportes Word")
  st.info(
      "Sube primero el reporte de medios tradicionales y después el reporte de"
      " redes sociales. El primer archivo se utilizará como plantilla para"
      " conservar su formato, estilos, colores, tabla, imágenes, encabezados y"
      " pies de página. El desglose incluirá los totales informativos y"
      " negativos de cada día. Las fechas de ambos archivos se unirán aunque"
      " no coincidan entre sí, pero el periodo y las fechas del primer archivo"
      " no se modificarán. Todo el documento se generará en Verdana 12."
  )

  col_trad, col_redes = st.columns(2)
  with col_trad:
    reporte_tradicional = st.file_uploader(
        "1. Reporte Word de medios tradicionales (plantilla visual)",
        type=["docx"],
        key="reporte_tradicional_unificar",
    )
  with col_redes:
    reporte_redes = st.file_uploader(
        "2. Reporte Word de redes sociales",
        type=["docx"],
        key="reporte_redes_unificar",
    )

  if reporte_tradicional and reporte_redes:
    if st.button(
        "Generar tercer reporte unificado",
        type="primary",
        use_container_width=True,
    ):
      with st.spinner(
          "Sumando cifras y ordenando tradicionales y redes por fecha..."
      ):
        try:
          reporte_unificado = unificar_reportes_word(
              reporte_tradicional, reporte_redes
          )
          st.session_state["reporte_word_unificado"] = (
              reporte_unificado.getvalue()
          )
          st.success(
              "El tercer reporte quedó unificado y conserva como base el"
              " formato del archivo de medios tradicionales."
          )
        except Exception as e:
          st.session_state.pop("reporte_word_unificado", None)
          st.error(f"No fue posible unificar los reportes: {str(e)}")

  if st.session_state.get("reporte_word_unificado"):
    st.download_button(
        label="📥 Descargar tercer reporte Word unificado",
        data=st.session_state["reporte_word_unificado"],
        file_name="Reporte_Unificado_Tradicionales_y_Redes.docx",
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        use_container_width=True,
    )

else:
  st.subheader("☁️ Reporte automático en Google Docs")
  st.info(
      "Inicia sesión con Google. En la primera extracción, el sistema hará una"
      " copia del Doc que contiene la fotografía y construirá el reporte. Más"
      " adelante podrás agregar redes sociales al mismo documento sin volver a"
      " subir tradicionales."
  )

  try:
    oauth_config = dict(st.secrets.get("google_oauth", {}))
    google_oauth = StreamlitGoogleOAuth(st, oauth_config)
    if google_oauth.procesar_retorno():
      st.success("Autorización de Google completada.")
      st.rerun()
    google_credentials = google_oauth.credentials()
  except Exception as exc:
    google_credentials = None
    st.error(f"No fue posible iniciar la conexión con Google: {exc}")
    st.caption(
        "Completa la sección [google_oauth] de .streamlit/secrets.toml usando"
        " el archivo de ejemplo incluido en el proyecto."
    )

  if google_credentials is None and "google_oauth" in locals():
    login_url = google_oauth.authorization_url()
    st.link_button(
        "🔐 Iniciar sesión con Google",
        login_url,
        type="primary",
        use_container_width=True,
    )
    st.stop()

  if google_credentials is not None:
    google_service = GoogleReportService(google_credentials)
    col_status, col_logout = st.columns([3, 1])
    with col_status:
      st.success("Google Drive y Google Docs autorizados para esta sesión.")
    with col_logout:
      if st.button("Cerrar sesión", use_container_width=True):
        google_oauth.logout()
        st.rerun()

    tab_primera, tab_redes = st.tabs([
        "1. Primera extracción",
        "2. Agregar redes sociales",
    ])

    with tab_primera:
      st.markdown("### Crear el reporte a partir del Doc con fotografía")
      plantilla_google = st.text_input(
          "Liga del Google Doc inicial",
          placeholder="https://docs.google.com/document/d/.../edit",
          key="plantilla_google_primera",
      )
      carpeta_google = st.text_input(
          "Liga de la carpeta de destino (opcional)",
          placeholder="https://drive.google.com/drive/folders/...",
          key="carpeta_google_destino",
          help=(
              "Si se deja vacía, la copia se guardará junto a la plantilla."
              " Para trabajo en equipo conviene usar una carpeta compartida."
          ),
      )
      archivos_trad_google = st.file_uploader(
          "Sube los archivos tradicionales",
          type=["xlsx", "xls", "csv"],
          accept_multiple_files=True,
          key="archivos_trad_google",
      )

      if archivos_trad_google:
        try:
          actores_google = _agrupar_archivos_por_actor(archivos_trad_google)
          actores_validos = sorted(
              actor for actor in actores_google if actor and "SIN NOTAS" not in actor
          )
          if not actores_validos:
            st.warning("No se detectaron hojas con notas válidas.")
          else:
            actor_detectado = st.selectbox(
                "Actor detectado en las hojas",
                actores_validos,
                key="actor_detectado_google",
            )
            actor_portada = st.text_input(
                "Nombre que aparecerá en el reporte",
                value=actor_detectado,
                key=f"actor_portada_google_{actor_detectado}",
                help=(
                    "Puedes corregirlo cuando el nombre de la pestaña no"
                    " corresponda al candidato de la fotografía."
                ),
            ).strip().upper()
            df_actor_google = pd.concat(
                actores_google[actor_detectado], ignore_index=True
            )
            inicio_sugerido, fin_sugerido = _obtener_rango_fechas_dataframe(
                df_actor_google
            )
            periodo_google = st.date_input(
                "Periodo de medición",
                value=(inicio_sugerido, fin_sugerido),
                key=f"periodo_google_{actor_detectado}",
                format="DD/MM/YYYY",
            )
            if isinstance(periodo_google, (tuple, list)) and len(periodo_google) == 2:
              fecha_inicio_google, fecha_fin_google = periodo_google
            else:
              fecha_inicio_google = fecha_fin_google = periodo_google

            if st.button(
                "Generar reporte en Google Docs",
                type="primary",
                use_container_width=True,
                disabled=not bool(plantilla_google and actor_portada),
            ):
              with st.spinner(
                  "Analizando notas, creando la copia y construyendo el reporte..."
              ):
                try:
                  modelo_google = construir_modelo_google_docs(
                      df_actor_google,
                      actor_portada,
                      es_redes_sociales=False,
                      fecha_inicio=fecha_inicio_google,
                      fecha_fin=fecha_fin_google,
                  )
                  resultado_google = google_service.crear_desde_plantilla(
                      plantilla_google,
                      modelo_google,
                      folder_url=carpeta_google or None,
                  )
                  st.session_state["ultimo_reporte_google"] = {
                      "url": resultado_google.document_url,
                      "title": resultado_google.title,
                  }
                  st.success(
                      "El reporte tradicional se creó sin modificar la plantilla"
                      " original."
                  )
                except Exception as exc:
                  st.error(f"No fue posible generar el Google Doc: {exc}")
        except Exception as exc:
          st.error(f"No fue posible leer los archivos tradicionales: {exc}")

      ultimo_google = st.session_state.get("ultimo_reporte_google")
      if ultimo_google:
        st.link_button(
            f"📄 Abrir {ultimo_google['title']}",
            ultimo_google["url"],
            use_container_width=True,
        )

    with tab_redes:
      st.markdown("### Agregar redes al reporte ya generado")
      reporte_google_existente = st.text_input(
          "Liga del reporte generado",
          value=st.session_state.get("ultimo_reporte_google", {}).get("url", ""),
          placeholder="https://docs.google.com/document/d/.../edit",
          key="reporte_google_existente",
      )
      archivo_redes_google = st.file_uploader(
          "Sube el Excel o CSV de redes sociales",
          type=["xlsx", "xls", "csv"],
          key="archivo_redes_google",
      )
      if reporte_google_existente and archivo_redes_google:
        if st.button(
            "Agregar redes y recalcular el reporte",
            type="primary",
            use_container_width=True,
        ):
          with st.spinner(
              "Evaluando sentimiento, integrando redes y recalculando cifras..."
          ):
            try:
              estado_google = google_service.cargar_estado(
                  reporte_google_existente
              )
              hojas_redes = cargar_archivo_seguro(archivo_redes_google)
              dataframes_redes = [
                  dataframe for dataframe in hojas_redes.values()
                  if dataframe is not None and not dataframe.empty
              ]
              if not dataframes_redes:
                raise ValueError("El archivo de redes no contiene publicaciones.")
              df_redes_google = pd.concat(dataframes_redes, ignore_index=True)
              modelo_redes_google = construir_modelo_google_docs(
                  df_redes_google,
                  estado_google["actor"],
                  es_redes_sociales=True,
                  fecha_inicio=estado_google["fecha_inicio"],
                  fecha_fin=estado_google["fecha_fin"],
              )
              resultado_actualizado = google_service.actualizar_con_redes(
                  reporte_google_existente,
                  modelo_redes_google,
              )
              st.session_state["ultimo_reporte_google"] = {
                  "url": resultado_actualizado.document_url,
                  "title": resultado_actualizado.title,
              }
              st.success(
                  "Redes sociales se integraron al mismo Google Doc y las"
                  " cifras fueron recalculadas."
              )
              st.link_button(
                  "📄 Abrir reporte actualizado",
                  resultado_actualizado.document_url,
                  use_container_width=True,
              )
            except Exception as exc:
              st.error(f"No fue posible actualizar el Google Doc: {exc}")
