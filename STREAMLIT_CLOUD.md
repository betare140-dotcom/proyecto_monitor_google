# Publicación en Streamlit Community Cloud

## Antes de desplegar

1. Sube estos archivos a un repositorio privado de GitHub.
2. No subas `.streamlit/secrets.toml`; ya está excluido mediante `.gitignore`.
3. En Streamlit Cloud selecciona `app.py` como archivo principal.
4. Copia la URL pública definitiva de la aplicación, por ejemplo:
   `https://monitor-politico.streamlit.app`.

## Configuración de Google Cloud

1. Habilita Google Drive API y Google Docs API.
2. Configura la pantalla de consentimiento OAuth.
3. Mientras esté en pruebas, agrega los correos del equipo como usuarios de
   prueba.
4. Crea un cliente OAuth de tipo **Aplicación web**.
5. En **URI de redireccionamiento autorizados**, registra la URL pública exacta
   de Streamlit, sin rutas adicionales.

## Secrets de Streamlit

En la aplicación publicada abre **Settings > Secrets** y pega:

```toml
GEMINI_API_KEY = "TU_CLAVE"

[google_oauth]
client_id = "TU_CLIENT_ID.apps.googleusercontent.com"
client_secret = "TU_CLIENT_SECRET"
redirect_uri = "https://TU-APLICACION.streamlit.app"
```

La URL de `redirect_uri` debe coincidir carácter por carácter con la registrada
en Google Cloud. No compartas públicamente el `client_secret` ni la clave de
Gemini.

## Primera prueba

1. Reinicia la aplicación desde Streamlit Cloud.
2. Selecciona **Google Docs automático**.
3. Pulsa **Iniciar sesión con Google**.
4. Autoriza Drive y Docs con uno de los usuarios de prueba.
5. Usa un Google Doc de prueba que contenga solamente la fotografía.
6. Sube una extracción tradicional pequeña y genera el reporte.
7. Verifica que se creó una copia y que la plantilla original no cambió.
