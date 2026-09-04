# Monitor político con Google Docs

Esta versión conserva la generación Word existente y agrega una cuarta opción:
**Google Docs automático (Primera extracción + Redes)**.

## Qué hace

1. Cada integrante inicia sesión con su cuenta de Google.
2. La aplicación copia un Google Doc inicial que puede contener únicamente la
   fotografía del candidato.
3. Con los archivos tradicionales construye nombre, periodo, tablas, resumen,
   temas y desglose diario.
4. Guarda un archivo JSON de control junto al reporte para no exigir nuevamente
   los tradicionales.
5. Al recibir redes sociales, carga ese estado, añade las publicaciones,
   conserva el periodo y recalcula todas las cifras en el mismo Google Doc.

El documento original nunca se modifica: cada reporte comienza desde una copia.

## 1. Crear el proyecto en Google Cloud

1. Entra a [Google Cloud Console](https://console.cloud.google.com/).
2. Crea un proyecto, por ejemplo `Monitor político`.
3. En **APIs y servicios > Biblioteca**, habilita:
   - Google Drive API
   - Google Docs API
4. En **APIs y servicios > Pantalla de consentimiento OAuth**:
   - Selecciona `Externo` si usarán cuentas personales, o `Interno` si todas las
     cuentas pertenecen al mismo Google Workspace.
   - Registra el nombre de la aplicación y los correos de contacto.
   - Mientras la aplicación esté en modo de prueba, agrega como usuarios de
     prueba las cuentas del equipo.
5. En **Credenciales**, crea un **ID de cliente OAuth > Aplicación web**.
6. Agrega una URI de redirección autorizada:
   - Local: `http://localhost:8501`
   - Publicada: la URL completa de la aplicación Streamlit, sin rutas extras.
7. Copia el `client_id` y el `client_secret` en los secretos de Streamlit.

La aplicación solicita acceso a Drive y Docs para copiar la plantilla, crear el
reporte, guardar el estado y actualizarlo. Google mostrará estos permisos antes
de que cada usuario decida autorizar.

Documentación oficial:

- [OAuth 2.0 para aplicaciones web](https://developers.google.com/identity/protocols/oauth2/web-server)
- [Combinar datos con una plantilla de Google Docs](https://developers.google.com/workspace/docs/api/how-tos/merge)
- [Google Docs batchUpdate](https://developers.google.com/workspace/docs/api/reference/rest/v1/documents/batchUpdate)

## 2. Configurar los secretos

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Edita `.streamlit/secrets.toml` y coloca:

```toml
GEMINI_API_KEY = "..."

[google_oauth]
client_id = "...apps.googleusercontent.com"
client_secret = "..."
redirect_uri = "http://localhost:8501"
```

En Streamlit Community Cloud, crea los mismos valores en **App settings >
Secrets** y cambia `redirect_uri` por la URL pública exacta de la aplicación.

## 3. Instalar y ejecutar

Windows PowerShell:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

Windows CMD:

```bat
py -m venv .venv
.venv\Scripts\activate.bat
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

## 4. Primera extracción

1. Selecciona **Google Docs automático**.
2. Pulsa **Iniciar sesión con Google** y acepta los permisos.
3. Pega la liga del Doc que contiene la fotografía.
4. Opcionalmente pega una carpeta compartida de destino.
5. Sube los tradicionales.
6. Confirma el nombre del candidato y el periodo.
7. Pulsa **Generar reporte en Google Docs**.

El nombre del reporte se construye con candidato y periodo. La plantilla con la
fotografía permanece intacta.

## 5. Agregar redes sociales

1. Abre la pestaña **Agregar redes sociales**.
2. Pega la liga del reporte creado en la primera extracción.
3. Sube el CSV o Excel de redes.
4. Pulsa **Agregar redes y recalcular el reporte**.

Tradicionales conservan el sentimiento del archivo. En redes, Gemini analiza
cada publicación; si no responde o la confianza es insuficiente, se conserva el
sentimiento que venía en el archivo. Solo las ligas quedan subrayadas.

## Formato generado

- Verdana 12.
- Periodo y tabla de balance centrados.
- Fechas con fondo turquesa, negritas y sin subrayado.
- Totales informativos en verde/amarillo.
- Totales negativos en rojo.
- Fuentes en negritas.
- Ligas azules y subrayadas.
- `local` y `portales locales` se cuentan como Portales digitales.
- `redes` y `redes sociales` se cuentan como Redes sociales.
- `prensa local` se conserva como Prensa local.

## Seguridad

- No coloques el `client_secret`, la clave de Gemini ni tokens dentro de
  `app.py`.
- No publiques `.streamlit/secrets.toml`.
- Mantén la aplicación OAuth en modo de prueba hasta completar las pruebas con
  el equipo.
- Para que varios integrantes actualicen el mismo reporte, usa una carpeta de
  Drive compartida con esas personas; el JSON de control se guarda junto al
  documento y no debe borrarse.

