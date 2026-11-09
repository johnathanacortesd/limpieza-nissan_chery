"""Detección de duplicados: el enlace de "Link (Streaming - Imagen)" manda.

app.py levanta la interfaz de Streamlit al importarse, así que aquí se cargan
solo las funciones de detección de duplicados (y sus dependencias) vía ast.
"""
from __future__ import annotations

import ast
import html
import re
import unittest
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd

_APP = Path(__file__).resolve().parent.parent / "app.py"
_NOMBRES = {
    "SIMILARITY_THRESHOLD_TITULOS", "_TRACKING_PARAMS", "convert_html_entities",
    "clean_title_for_output", "normalize_title_for_comparison", "normalize_url",
    "calculate_title_quality_score", "are_duplicates", "detect_duplicates",
}


def _cargar_funciones() -> dict:
    tree = ast.parse(_APP.read_text(encoding="utf-8"))
    nodos = []
    for nodo in tree.body:
        if isinstance(nodo, ast.FunctionDef) and nodo.name in _NOMBRES:
            nodos.append(nodo)
        elif isinstance(nodo, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id in _NOMBRES for t in nodo.targets
        ):
            nodos.append(nodo)
    ns = {"pd": pd, "re": re, "html": html, "SequenceMatcher": SequenceMatcher}
    exec(compile(ast.Module(body=nodos, type_ignores=[]), str(_APP), "exec"), ns)
    faltan = _NOMBRES - ns.keys()
    assert not faltan, f"No se encontraron en app.py: {faltan}"
    return ns


APP = _cargar_funciones()
normalize_url = APP["normalize_url"]
are_duplicates = APP["are_duplicates"]
detect_duplicates = APP["detect_duplicates"]

LINK_BMW = ("https://paisminero.com/index.php?option=com_content&view=article"
            "&id=32884:bmw-ofrece-carga-gratuita-de-vehiculos-electricos-y-aun-no-has-visto-nada"
            "-20261108060306&catid=131&Itemid=928")
LINK_GM = ("https://paisminero.com/index.php?option=com_content&view=article"
           "&id=32886:gm-pisa-los-pies-de-tesla-en-el-sector-de-vehiculos-electricos-con-la"
           "-celebracio-20261108060450&catid=192&Itemid=10003")


def _fila(id_noticia, titulo, link_stream, **extra):
    fila = {
        "ID Noticia": id_noticia,
        "Fecha": datetime(2026, 11, 8),
        "Hora": "",
        "Medio": "Paisminero.com",
        "Tipo de Medio": "Internet",
        "Título": titulo,
        "Link Nota": ("http://news.globalnews.com.co/Validar.aspx?n="
                      f"{id_noticia}&u=50e96988-ec1d-4356-a82a-d70bf7d74619&c=16642&m=i"),
        "Link (Streaming - Imagen)": link_stream,
        "Menciones - Empresa": "Nissan",
    }
    fila.update(extra)
    return pd.Series(fila)


class NormalizeUrlTests(unittest.TestCase):
    def test_conserva_query_string_que_identifica_la_nota(self):
        self.assertNotEqual(normalize_url(LINK_BMW), normalize_url(LINK_GM))

    def test_ignora_esquema_www_fragmento_y_rastreo(self):
        a = "https://www.medio.com/nota/123/?utm_source=x&fbclid=abc#comentarios"
        b = "http://medio.com/nota/123"
        self.assertEqual(normalize_url(a), normalize_url(b))

    def test_orden_de_parametros_no_importa(self):
        self.assertEqual(normalize_url("https://m.com/a?id=1&cat=2"),
                         normalize_url("https://m.com/a?cat=2&id=1"))

    def test_valor_no_url(self):
        self.assertEqual(normalize_url("Link"), "")
        self.assertEqual(normalize_url(None), "")


class AreDuplicatesTests(unittest.TestCase):
    def test_caso_reportado_paisminero_no_es_duplicada(self):
        bmw = _fila(61836227, "BMW Ofrece Carga Gratuita de Vehículos Eléctricos y Aún No Has Visto Nada", LINK_BMW)
        gm = _fila(61836232, "GM pisa los pies de Tesla en el sector de vehículos eléctricos con la celebración del EV1", LINK_GM)
        self.assertFalse(are_duplicates(bmw, gm))

    def test_sin_hipervinculo_en_stream_link_nota_distinto_no_es_duplicada(self):
        # Segundo archivo reportado: "Link (Streaming - Imagen)" quedó como el
        # texto "Link" (sin hipervínculo) y los Link Nota solo difieren en ?n=...
        def nota(n):
            return ("http://news.globalnews.com.co/Validar.aspx?n="
                    f"{n}&u=50e96988-ec1d-4356-a82a-d70bf7d74619&m=audiovisual&lang=es")
        bmw = _fila(61836227, "BMW Ofrece Carga Gratuita de Vehículos Eléctricos y Aún No Has Visto Nada",
                    "Link", **{"Link Nota": nota(11987851)})
        gm = _fila(61836232, "GM pisa los pies de Tesla en el sector de vehículos eléctricos con la celebración del EV1",
                   "Link", **{"Link Nota": nota(11992065)})
        aura = _fila(61841785, "Nuevo ID. AURA es el vehículo eléctrico que Volkswagen necesita, pero su software falla",
                     "Link", **{"Link Nota": nota(12000001)})
        self.assertFalse(are_duplicates(bmw, gm))
        self.assertFalse(are_duplicates(bmw, aura))
        out = detect_duplicates(pd.DataFrame([bmw, gm, aura]))
        self.assertFalse(out["is_duplicate"].any())

    def test_mismo_titulo_distinto_link_no_es_duplicada(self):
        a = _fila(1, "Nissan lanza nuevo modelo", "https://m.com/nota?id=1")
        b = _fila(2, "Nissan lanza nuevo modelo", "https://m.com/nota?id=2")
        self.assertFalse(are_duplicates(a, b))

    def test_distinto_titulo_mismo_link_es_duplicada(self):
        a = _fila(1, "Nissan lanza nuevo modelo", "https://m.com/nota?id=1")
        b = _fila(2, "Otro título completamente diferente", "https://www.m.com/nota?id=1")
        self.assertTrue(are_duplicates(a, b))

    def test_mismo_link_distinta_mencion_no_es_duplicada(self):
        a = _fila(1, "Nota", "https://m.com/nota?id=1")
        b = _fila(2, "Nota", "https://m.com/nota?id=1", **{"Menciones - Empresa": "Chery"})
        self.assertFalse(are_duplicates(a, b))

    def test_mismo_link_distinto_medio_no_es_duplicada(self):
        a = _fila(1, "Nota", "https://m.com/nota?id=1")
        b = _fila(2, "Nota", "https://m.com/nota?id=1", Medio="Otro medio")
        self.assertFalse(are_duplicates(a, b))

    def test_sin_link_sigue_comparando_titulo(self):
        a = _fila(1, "Nissan lanza nuevo modelo en Colombia", None)
        b = _fila(2, "Nissan lanza nuevo modelo en Colombia", None)
        self.assertTrue(are_duplicates(a, b))


class DetectDuplicatesTests(unittest.TestCase):
    def test_caso_reportado_ambas_quedan_unicas(self):
        df = pd.DataFrame([
            _fila(61836227, "BMW Ofrece Carga Gratuita de Vehículos Eléctricos y Aún No Has Visto Nada", LINK_BMW),
            _fila(61836232, "GM pisa los pies de Tesla en el sector de vehículos eléctricos con la celebración del EV1", LINK_GM),
        ])
        out = detect_duplicates(df)
        self.assertEqual(out["is_duplicate"].tolist(), [False, False])
        self.assertTrue(out["ID Original Retenido"].isna().all())


if __name__ == "__main__":
    unittest.main()
