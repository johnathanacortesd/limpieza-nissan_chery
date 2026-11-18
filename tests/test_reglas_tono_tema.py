"""Reglas de negocio sobre Tono / Tema / Subtema (accidentes, delitos, Juan Carlos López)."""
from __future__ import annotations

import ast
import re
import unicodedata
import unittest
from pathlib import Path

import pandas as pd

_APP = Path(__file__).resolve().parent.parent / "app.py"


def _cargar() -> dict:
    """Carga solo la sección REGLAS DE NEGOCIO de app.py (que importa Streamlit)."""
    fuente = _APP.read_text(encoding="utf-8")
    inicio = fuente.index("# REGLAS DE NEGOCIO SOBRE TONO / TEMA")
    fin = fuente.index("# GENERACIÓN DEL EXCEL DE SALIDA")
    seccion = ast.parse(fuente[inicio:fin])
    ns = {"re": re, "unicodedata": unicodedata, "pd": pd}
    exec(compile(seccion, str(_APP), "exec"), ns)
    return ns


R = _cargar()
es_accidente, es_delito = R["es_accidente"], R["es_delito"]
aplicar = R["aplicar_reglas_tono_tema"]
SUB = "Temas Generales - Tema"


def _df(*filas):
    base = {"Título": "", "Resumen - Aclaracion": "", "Autor - Conductor": "",
            "Menciones - Empresa": "Nissan", "Tono": "Neutro", "Tema": "Tecnológico",
            SUB: "Electrificación", "is_duplicate": False}
    return pd.DataFrame([{**base, **f} for f in filas])


class DeteccionTests(unittest.TestCase):
    def test_accidentes_en_titulo(self):
        for t in ("Grave accidente de tránsito en la vía al Llano deja dos heridos",
                  "Camioneta Nissan se volcó en la autopista Norte",
                  "Motociclista fue atropellado por un bus en Medellín",
                  "Choque múltiple en la Calle 80"):
            self.assertTrue(es_accidente(t, ""), t)

    def test_accidente_en_resumen_necesita_dos_senales(self):
        self.assertTrue(es_accidente("Tragedia en Boyacá", "Un bus chocó contra un camión y dos personas murieron."))
        self.assertFalse(es_accidente("Nuevo Nissan Kicks", "El SUV se presentó tras el accidente comercial de 2020."))

    def test_producto_no_es_accidente(self):
        self.assertFalse(es_accidente("Chery Tiggo 7 obtiene 5 estrellas en pruebas de choque de Latin NCAP", ""))
        self.assertFalse(es_accidente(
            "Nissan presenta su nuevo sistema de seguridad",
            "El asistente de frenado ayuda a evitar accidentes y la alerta de colisión frontal avisa al conductor."))
        self.assertFalse(es_accidente("Murió el fundador de una marca de carros", ""))

    def test_delitos(self):
        for t in ("Capturan banda dedicada al hurto de camionetas en Bogotá",
                  "Aumenta el robo de vehículos en Colombia",
                  "Desmantelan red de delincuentes que vendía autopartes robadas"):
            self.assertTrue(es_delito(t, ""), t)
        self.assertTrue(es_delito("Operativo en Cali", "La Policía recuperó tres carros hurtados."))

    def test_producto_no_es_delito(self):
        self.assertFalse(es_delito("Nissan estrena sistema antirrobo en la Frontier", ""))
        self.assertFalse(es_delito("Tesla capturó el 30 % del mercado", "La marca lidera ventas."))


class AplicarReglasTests(unittest.TestCase):
    def _valores(self, df):
        return list(zip(df["Tono"], df["Tema"], df[SUB]))

    def test_accidente(self):
        out = aplicar(_df({"Título": "Accidente de tránsito en la vía Bogotá-Girardot"}))
        self.assertEqual(self._valores(out), [("No aplica", "Accidentes", "-")])

    def test_delito(self):
        out = aplicar(_df({"Título": "Aumenta el hurto de motos en Medellín"}))
        self.assertEqual(self._valores(out), [("No aplica", "Orden público", "-")])

    def test_juan_carlos_lopez_por_marca(self):
        resumen = "Juan Carlos López, gerente de la marca, habló de las ventas del año."
        out = aplicar(_df(
            {"Resumen - Aclaracion": resumen, "Menciones - Empresa": "Nissan"},
            {"Resumen - Aclaracion": resumen, "Menciones - Empresa": "Chery"},
            {"Resumen - Aclaracion": resumen, "Menciones - Empresa": "AVATR"},
            {"Resumen - Aclaracion": resumen, "Menciones - Empresa": "Changan"},
            {"Resumen - Aclaracion": resumen, "Menciones - Empresa": "Deepal"},
            {"Resumen - Aclaracion": resumen, "Menciones - Empresa": "Toyota"},
        ))
        self.assertEqual(self._valores(out), [
            ("No aplica", "-", "-"),
            ("No aplica", "-", "-"),
            ("Positivo", "Indefinido", "Proyecciones - Balances"),
            ("Positivo", "Indefinido", "Proyecciones - Balances"),
            ("Positivo", "Indefinido", "Proyecciones - Balances"),
            ("Neutro", "Tecnológico", "Electrificación"),
        ])

    def test_juan_carlos_lopez_sin_tilde_ni_mayusculas(self):
        out = aplicar(_df({"Título": "JUAN CARLOS LOPEZ habla del mercado", "Menciones - Empresa": "Deepal"}))
        self.assertEqual(self._valores(out), [("Positivo", "Indefinido", "Proyecciones - Balances")])

    def test_duplicadas_y_notas_normales_no_cambian(self):
        out = aplicar(_df(
            {"Título": "Accidente en la vía", "is_duplicate": True, "Tono": "Duplicada", "Tema": "-", SUB: "-"},
            {"Título": "Nissan lanza la nueva Frontier"},
        ))
        self.assertEqual(self._valores(out), [("Duplicada", "-", "-"), ("Neutro", "Tecnológico", "Electrificación")])


if __name__ == "__main__":
    unittest.main()
