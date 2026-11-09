"""Resumen - Aclaracion: siempre termina en un solo punto, nunca en "..."."""
from __future__ import annotations

import ast
import html
import re
import unittest
from pathlib import Path

_APP = Path(__file__).resolve().parent.parent / "app.py"
_NOMBRES = {"convert_html_entities", "corregir_resumen"}


def _cargar():
    tree = ast.parse(_APP.read_text(encoding="utf-8"))
    nodos = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in _NOMBRES]
    ns = {"re": re, "html": html}
    exec(compile(ast.Module(body=nodos, type_ignores=[]), str(_APP), "exec"), ns)
    return ns["corregir_resumen"]


corregir_resumen = _cargar()


class CorregirResumenTests(unittest.TestCase):
    def test_tres_puntos_pasa_a_un_punto(self):
        self.assertEqual(corregir_resumen("Nissan lanza el Kicks..."), "Nissan lanza el Kicks.")

    def test_caracter_elipsis_pasa_a_un_punto(self):
        self.assertEqual(corregir_resumen("Nissan lanza el Kicks…"), "Nissan lanza el Kicks.")

    def test_sin_punto_final_se_agrega(self):
        self.assertEqual(corregir_resumen("Nissan lanza el Kicks"), "Nissan lanza el Kicks.")

    def test_ya_termina_en_punto_no_se_duplica(self):
        self.assertEqual(corregir_resumen("Nissan lanza el Kicks."), "Nissan lanza el Kicks.")

    def test_corchetes_y_signos_sueltos_al_final(self):
        self.assertEqual(corregir_resumen("Nissan lanza el Kicks, [...]"), "Nissan lanza el Kicks.")
        self.assertEqual(corregir_resumen("Nissan lanza el Kicks; ... "), "Nissan lanza el Kicks.")

    def test_caso_real_social2s(self):
        texto = "Fuente: Cleantechnica Detalles Visitas: 28 powered by social2s..."
        self.assertEqual(corregir_resumen(texto), "Fuente: Cleantechnica Detalles Visitas: 28 powered by social2s.")

    def test_vacio_y_no_texto(self):
        self.assertEqual(corregir_resumen(""), "")
        self.assertEqual(corregir_resumen("..."), "")
        self.assertIsNone(corregir_resumen(None))


if __name__ == "__main__":
    unittest.main()
