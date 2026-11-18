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
detecta_vocero = R["detecta_vocero"]
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
        self.assertTrue(es_accidente("Tragedia en Boyacá",
                                     "Grave accidente: un bus chocó contra un camión y dos personas murieron."))
        # Una sola palabra fuerte en el resumen ya no basta (calibrado con la base de nov. 2026).
        self.assertFalse(es_accidente("Tragedia en Boyacá", "Un bus chocó contra un camión y dos personas murieron."))
        self.assertFalse(es_accidente("Volvo reafirma su promesa de seguridad",
                                      "Volvo busca reducir los heridos en accidentes de tránsito en la carretera."))
        self.assertFalse(es_accidente("Nuevo Nissan Kicks", "El SUV se presentó tras el accidente comercial de 2020."))

    def test_producto_no_es_accidente(self):
        self.assertFalse(es_accidente("Chery Tiggo 7 obtiene 5 estrellas en pruebas de choque de Latin NCAP", ""))
        self.assertFalse(es_accidente(
            "Nissan presenta su nuevo sistema de seguridad",
            "El asistente de frenado ayuda a evitar accidentes y la alerta de colisión frontal avisa al conductor."))
        self.assertFalse(es_accidente("Murió el fundador de una marca de carros", ""))
        self.assertFalse(es_accidente("Cupra León VZ y Golf GTI suman sistemas que evitan choques", ""))
        self.assertFalse(es_accidente("Volvo exhibe un vehículo accidentado en el Salón del Automóvil", ""))

    def test_delitos(self):
        for t in ("Capturan banda dedicada al hurto de camionetas en Bogotá",
                  "Aumenta el robo de vehículos en Colombia",
                  "Desmantelan red de delincuentes que vendía autopartes robadas"):
            self.assertTrue(es_delito(t, ""), t)
        for t in ("Buscan a dos mujeres desaparecidas en Barranquilla",
                  "Investigan homicidio de un hombre dentro de un vehículo en Medellín",
                  "Secuestrada la gestora social de Uribia",
                  "Incendio en establecimiento comercial de Cartagena",
                  "Delincuentes armados perpetraron millonario robo en el oriente de Cali"):
            self.assertTrue(es_delito(t, ""), t)
        # En el resumen hace falta mucha evidencia: una mención suelta no basta.
        self.assertFalse(es_delito("Operativo en Cali", "La Policía recuperó tres carros hurtados."))
        self.assertTrue(es_delito("Resumen de noticias", "Robo, hurto de motos, un secuestro y un homicidio: "
                                  "la Policía y la Fiscalía capturaron a los delincuentes."))

    def test_producto_no_es_delito(self):
        self.assertFalse(es_delito("Nissan estrena sistema antirrobo en la Frontier", ""))
        self.assertFalse(es_delito("Tesla capturó el 30 % del mercado", "La marca lidera ventas."))
        self.assertFalse(es_delito("Con 57 años y 1,7 millones de carros armados, Renault usa su planta", ""))


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
            ("No aplica", "Juan Carlos López", "-"),
            ("No aplica", "Juan Carlos López", "-"),
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


class VocerosTests(unittest.TestCase):
    RESUMEN_LARGO = ("La marca presentó su portafolio para 2027 en el Salón del Automóvil. " * 30)

    def test_voceros_nissan_por_nombre_aunque_esten_al_final(self):
        for nombre in ("Leonardo Gutiérrez", "LEONARDO GUTIERREZ", "Carlos Vargas"):
            resumen = self.RESUMEN_LARGO + f"Así lo explicó {nombre} durante el evento."
            self.assertTrue(detecta_vocero("Nissan", "Nissan en el Salón", resumen), nombre)

    def test_nombre_de_vocero_de_nissan_no_marca_otra_marca(self):
        self.assertFalse(detecta_vocero("Chery", "", "Leonardo Gutiérrez habló del mercado."))

    def test_cargo_ligado_a_la_marca_de_la_fila(self):
        casos = [
            ("Nissan", "Ana Pérez, gerente general de Nissan Colombia, habló de las ventas."),
            ("Nissan", "El director de mercadeo y ventas de Nissan dio un balance del año."),
            ("Nissan", "Nissan Colombia, a través de su presidente, anunció nuevas metas."),
            ("Chery", "El CEO de Chery visitó Bogotá."),
            ("Changan", "Según el vocero de la marca Changan, las ventas crecieron."),
            ("Chevrolet", "La presidenta de General Motors Colombia habló del mercado."),
            ("Deepal", "Pedro Gómez, country manager de Deepal, explicó la estrategia."),
        ]
        for marca, texto in casos:
            self.assertTrue(detecta_vocero(marca, "", self.RESUMEN_LARGO + texto), texto)

    def test_cargo_de_otra_entidad_no_cuenta(self):
        casos = [
            ("Nissan", "Elon Musk, CEO de Tesla, habló de los carros eléctricos de Nissan."),
            ("Nissan", "El presidente de Fenalco, en la presentación, mencionó a Nissan."),
            ("Nissan", "El exgerente de Nissan abrió su propia empresa."),
            ("Nissan", "El gerente de Toyota destacó la Hilux; Nissan también exhibió la Frontier."),
            ("Chery", "Nissan Colombia, a través de su presidente, anunció nuevas metas."),
        ]
        for marca, texto in casos:
            self.assertFalse(detecta_vocero(marca, "", texto), texto)

    def test_voceros_desde_hoja(self):
        self.assertFalse(detecta_vocero("Chery", "", "María Ruiz presentó el Tiggo 4."))
        voceros = R["preparar_voceros"]({"Chery": ["María Ruiz"]})
        self.assertTrue(detecta_vocero("Chery", "", "María Ruiz presentó el Tiggo 4.", voceros))

    def test_aplicar_conserva_tono_y_usa_mapa_de_temas(self):
        out = aplicar(
            _df({"Resumen - Aclaracion": self.RESUMEN_LARGO + "Lo dijo Carlos Vargas.", "Tono": "Positivo"},
                {"Título": "Nissan lanza la nueva Frontier"}),
            final_topic_map={"Proyecciones - Balances": "Indefinido"},
        )
        self.assertEqual(list(zip(out["Tono"], out["Tema"], out[SUB])), [
            ("Positivo", "Indefinido", "Proyecciones - Balances"),
            ("Neutro", "Tecnológico", "Electrificación"),
        ])

    def test_nota_con_varias_menciones_solo_cambia_la_marca_del_vocero(self):
        # Misma nota expandida en una fila por mención (expand_menciones):
        # habla el CEO de Renault -> solo la fila de Renault va a
        # Proyecciones - Balances; Nissan y Chery quedan como dijo el PKL.
        resumen = (self.RESUMEN_LARGO + "Luca de Meo, CEO de Renault, habló de las ventas. "
                   "Nissan y Chery también exhibieron sus modelos.")
        comun = {"Título": "Así fue el Salón del Automóvil", "Resumen - Aclaracion": resumen}
        out = aplicar(
            _df({**comun, "Menciones - Empresa": "Renault", "Tono": "Neutro", "Tema": "Indefinido", SUB: "Eventos"},
                {**comun, "Menciones - Empresa": "Nissan", "Tono": "Positivo", "Tema": "Indefinido", SUB: "Lanzamientos"},
                {**comun, "Menciones - Empresa": "Chery", "Tono": "Neutro", "Tema": "Tecnológico", SUB: "Electrificación"}),
            final_topic_map={"Proyecciones - Balances": "Indefinido"},
        )
        self.assertEqual(list(zip(out["Menciones - Empresa"], out["Tono"], out["Tema"], out[SUB])), [
            ("Renault", "Neutro", "Indefinido", "Proyecciones - Balances"),
            ("Nissan", "Positivo", "Indefinido", "Lanzamientos"),
            ("Chery", "Neutro", "Tecnológico", "Electrificación"),
        ])

    def test_prioridades_accidente_y_juan_carlos_lopez_ganan(self):
        out = aplicar(_df(
            {"Título": "Accidente en la vía", "Resumen - Aclaracion": "Habló Carlos Vargas."},
            {"Resumen - Aclaracion": "Juan Carlos López y el gerente general de Nissan hablaron."},
        ))
        self.assertEqual(list(zip(out["Tono"], out["Tema"], out[SUB])), [
            ("No aplica", "Accidentes", "-"),
            ("No aplica", "Juan Carlos López", "-"),
        ])


class DescubrirVocerosTests(unittest.TestCase):
    def test_descubre_nombre_con_cargo_y_lo_usa_en_otra_nota(self):
        # Nota 1 trae nombre + cargo; nota 2 solo el nombre, al final de un resumen largo.
        largo = "La marca mostró su portafolio para 2027. " * 30
        df = _df(
            {"Resumen - Aclaracion": "Laura Méndez, gerente comercial de Nissan Colombia, habló del Salón."},
            {"Resumen - Aclaracion": largo + "Así lo explicó Laura Méndez en entrevista."},
            {"Resumen - Aclaracion": largo + "Así lo explicó Laura Méndez en entrevista.", "Menciones - Empresa": "Chery"},
        )
        out = aplicar(df)
        self.assertEqual(out.attrs["voceros_descubiertos"], {"Nissan": ["Laura Mendez"]})
        self.assertEqual(out[SUB].tolist(), ["Proyecciones - Balances", "Proyecciones - Balances", "Electrificación"])

    def test_formas_de_presentar_al_vocero(self):
        casos = {
            "El gerente general de Nissan Colombia, Pedro Ramírez, dio el balance.": "Pedro Ramirez",
            "En Bogotá Ana María Ruiz, directora de mercadeo de Nissan, habló.": "Ana Maria Ruiz",
            "Según Carlos de la Torre, vocero de la marca Nissan, las ventas crecieron.": "Carlos de la Torre",
        }
        for texto, esperado in casos.items():
            hallados = R["descubrir_voceros"](_df({"Resumen - Aclaracion": texto}))
            self.assertIn(esperado, hallados.get("Nissan", []), texto)

    def test_quita_rotulos_pegados_al_nombre(self):
        casos = {
            "Declaración Clave Andrea Saldívar, gerente de Mercadeo de Hyundai Colombia, habló.": ("Hyundai", "Andrea Saldivar"),
            "Oficial Juan Pérez, gerente general de Chery, habló.": ("Chery", "Juan Perez"),
            "Aa Felipe Negret, gerente general de Farizon en Grupo Vardí.": ("Farizon", "Felipe Negret"),
        }
        for texto, (marca, esperado) in casos.items():
            hallados = R["descubrir_voceros"](_df({"Resumen - Aclaracion": texto, "Menciones - Empresa": marca}))
            self.assertEqual(hallados.get(marca), [esperado], texto)
        for texto, marca in (("Ma López, vicepresidente del Centro de Diseño de XPENG.", "Xpeng"),
                             ("La directora de mercadeo de iCAUR para Colombia Datos Clave", "iCAUR")):
            self.assertEqual(R["descubrir_voceros"](_df({"Resumen - Aclaracion": texto, "Menciones - Empresa": marca})), {}, texto)

    def test_no_descubre_lugares_ni_cargos_de_otra_marca(self):
        textos = [
            "Bogotá Colombia, gerente de Nissan, no es un nombre.",
            "Elon Musk, CEO de Tesla, habló de Nissan.",
        ]
        for t in textos:
            self.assertEqual(R["descubrir_voceros"](_df({"Resumen - Aclaracion": t})), {}, t)

    def test_distribuidor_de_una_sola_marca(self):
        self.assertTrue(detecta_vocero("Nissan", "", "El gerente general de Dinissan habló del mercado."))
        self.assertTrue(detecta_vocero("Renault", "", "El presidente de Sofasa anunció exportaciones."))
        self.assertTrue(detecta_vocero("Chery", "", "La gerente general de Auto Vardí explicó la meta."))

    def test_voceros_investigados(self):
        for marca, nombre in (("Renault", "Juan Camilo Vélez"), ("Toyota", "Hiroshi Yonenaga"),
                              ("Hyundai", "Adriana Casadiego"), ("JMC", "Noel Ardila Rey"),
                              ("iCAUR", None), ("Nissan", "Ricardo Toro")):
            if nombre:
                self.assertTrue(detecta_vocero(marca, "", f"Habló {nombre} sobre el año."), marca)


class CeoGlobalesTests(unittest.TestCase):
    def test_musk_en_nota_de_tesla_va_a_proyecciones(self):
        for texto in ("Elon Musk anunció la producción del Cybercab.",
                      "Tesla presentó resultados. " * 20 + "Según Musk, el robotaxi llegará en 2027."):
            self.assertTrue(detecta_vocero("Tesla", "", texto), texto)
        out = aplicar(_df({"Título": "Musk habla del futuro de la marca", "Menciones - Empresa": "Tesla",
                           "Tono": "Neutro", SUB: "Lanzamientos"}))
        self.assertEqual(out[SUB].tolist(), ["Proyecciones - Balances"])
        self.assertEqual(out["Tono"].tolist(), ["Neutro"])

    def test_musk_no_cuenta_para_otra_marca(self):
        self.assertFalse(detecta_vocero("Nissan", "", "Elon Musk, CEO de Tesla, criticó a Nissan."))

    def test_ceo_global_por_marca(self):
        for marca, nombre in (("Toyota", "Kenta Kon"), ("BYD", "Wang Chuanfu"), ("Volkswagen", "Oliver Blume"),
                              ("Chevrolet", "Mary Barra"), ("Ford", "Jim Farley"), ("Volvo", "Håkan Samuelsson"),
                              ("Hyundai", "Chung Eui-sun"), ("Xpeng", "He Xiaopeng"), ("Mazda", "Masahiro Moro")):
            self.assertTrue(detecta_vocero(marca, "", f"{nombre} habló de las metas del grupo."), marca)
        self.assertFalse(detecta_vocero("Kia", "", "Kenta Kon habló de las metas de Toyota."))


class AliasMarcaTests(unittest.TestCase):
    def test_jetour_nissan_pasa_a_jetour(self):
        fuente = _APP.read_text(encoding="utf-8")
        self.assertIn("'jetour - nissan': 'Jetour'", fuente)


if __name__ == "__main__":
    unittest.main()
