import sys
import types
import unittest


def _install_google_stubs():
    modules = {
        "google": types.ModuleType("google"),
        "google.auth": types.ModuleType("google.auth"),
        "google.auth.transport": types.ModuleType("google.auth.transport"),
        "google.auth.transport.requests": types.ModuleType("google.auth.transport.requests"),
        "google.oauth2": types.ModuleType("google.oauth2"),
        "google.oauth2.credentials": types.ModuleType("google.oauth2.credentials"),
        "google_auth_oauthlib": types.ModuleType("google_auth_oauthlib"),
        "google_auth_oauthlib.flow": types.ModuleType("google_auth_oauthlib.flow"),
        "googleapiclient": types.ModuleType("googleapiclient"),
        "googleapiclient.discovery": types.ModuleType("googleapiclient.discovery"),
        "googleapiclient.http": types.ModuleType("googleapiclient.http"),
    }
    modules["google.auth.transport.requests"].Request = object
    modules["google.oauth2.credentials"].Credentials = object
    modules["google_auth_oauthlib.flow"].Flow = object
    modules["googleapiclient.discovery"].build = lambda *a, **k: None
    modules["googleapiclient.http"].MediaIoBaseDownload = object
    modules["googleapiclient.http"].MediaIoBaseUpload = object
    sys.modules.update(modules)


try:
    import google.oauth2.credentials  # noqa: F401
except ModuleNotFoundError:
    _install_google_stubs()

from google_workspace_report import (  # noqa: E402
    AUTO_END,
    AUTO_START,
    BALANCE_TABLE,
    GoogleReportService,
    PERIOD_TABLE,
    SOURCE_MARKER,
    calcular_metricas,
    combinar_modelos,
    extraer_id_google,
    normalizar_modelo,
    periodo_espanol,
)


class GoogleReportCoreTests(unittest.TestCase):
    def sample_model(self):
        return {
            "actor": "Persona de prueba",
            "fecha_inicio": "2026-09-04",
            "fecha_fin": "2026-09-10",
            "temas_informativos": ["Tema uno"],
            "temas_negativos": ["Tema negativo"],
            "notas": [
                {
                    "fecha": "2026-09-04",
                    "canal": "PORTALES DIGITALES",
                    "sentimiento": "POSITIVA",
                    "fuente": "Medio A",
                    "contenido": "Nota informativa",
                    "url": "https://example.com/a",
                    "origen": "TRADICIONALES",
                },
                {
                    "fecha": "2026-09-04",
                    "canal": "REDES SOCIALES",
                    "sentimiento": "NEGATIVA",
                    "fuente": "Cuenta B",
                    "contenido": "Crítica directa",
                    "url": "https://example.com/b",
                    "origen": "REDES",
                },
            ],
        }

    def test_extract_document_and_folder_ids(self):
        doc_id = "1AbCdEfGhIjKlMnOpQrStUvWxYz12"
        self.assertEqual(
            extraer_id_google(f"https://docs.google.com/document/d/{doc_id}/edit"),
            doc_id,
        )
        self.assertEqual(
            extraer_id_google(f"https://drive.google.com/drive/folders/{doc_id}"),
            doc_id,
        )

    def test_metrics_are_split_by_sentiment_and_channel(self):
        model = normalizar_modelo(self.sample_model())
        metrics = calcular_metricas(model)
        self.assertEqual(metrics["positiva"], 1)
        self.assertEqual(metrics["negativa"], 1)
        self.assertEqual(metrics["total"], 2)
        self.assertEqual(metrics["positivos_canal"]["PORTALES DIGITALES"], 1)
        self.assertEqual(metrics["negativos_canal"]["REDES SOCIALES"], 1)

    def test_merge_keeps_original_period_and_avoids_same_url_twice(self):
        base = self.sample_model()
        social = {
            **self.sample_model(),
            "fecha_inicio": "2026-09-01",
            "fecha_fin": "2026-09-30",
            "notas": [base["notas"][1]],
        }
        merged = combinar_modelos(base, social)
        self.assertEqual(merged["fecha_inicio"], "2026-09-04")
        self.assertEqual(merged["fecha_fin"], "2026-09-10")
        self.assertEqual(len(merged["notas"]), 2)

    def test_spanish_period(self):
        self.assertEqual(
            periodo_espanol("2026-09-04", "2026-09-10"),
            "4 al 10 de septiembre de 2026",
        )

    def test_document_skeleton_contains_sections_and_daily_totals(self):
        service = object.__new__(GoogleReportService)
        skeleton = service._build_skeleton(normalizar_modelo(self.sample_model()))
        self.assertIn(AUTO_START, skeleton)
        self.assertIn(AUTO_END, skeleton)
        self.assertIn(PERIOD_TABLE, skeleton)
        self.assertIn(BALANCE_TABLE, skeleton)
        self.assertIn("RESUMEN", skeleton)
        self.assertIn("DESGLOSE", skeleton)
        self.assertIn("TOTAL DE IMPACTOS INFORMATIVOS: 1", skeleton)
        self.assertIn("TOTAL DE IMPACTOS NEGATIVOS: 1", skeleton)
        self.assertEqual(skeleton.count(SOURCE_MARKER), 2)


if __name__ == "__main__":
    unittest.main()
