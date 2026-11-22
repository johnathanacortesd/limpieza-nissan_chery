"""Reglas de negocio sobre Tono / Tema / Subtema (accidentes, delitos, Juan Carlos López)."""
from __future__ import annotations

import ast
import re
import unicodedata
import unittest
from functools import lru_cache
from pathlib import Path

import pandas as pd

_APP = Path(__file__).resolve().parent.parent / "app.py"


def _cargar() -> dict:
    """Carga solo la sección REGLAS DE NEGOCIO de app.py (que importa Streamlit)."""
    fuente = _APP.read_text(encoding="utf-8")
    inicio = fuente.index("# REGLAS DE NEGOCIO SOBRE TONO / TEMA")
    fin = fuente.index("# GENERACIÓN DEL EXCEL DE SALIDA")
    seccion = ast.parse(fuente[inicio:fin])
    ns = {"re": re, "unicodedata": unicodedata, "pd": pd, "lru_cache": lru_cache}
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
    def test_casos_reales_de_la_base_competencia_nov_2026(self):
        self.assertTrue(es_accidente(
            "Lluvias en Sabaneta y balance de daños materiales en vehículos y parqueaderos tras inundaciones en El Poblado",
            ""))
        self.assertTrue(es_accidente("Lluvias en Sabaneta", (
            "En el balance de afectaciones materiales ocasionadas por el desbordamiento de la quebrada La Presidenta "
            "en el sector de El Poblado en Medellín, se reportaron vehículos inundados entre ellos un automóvil "
            "Renault Twingo y un Ford Fiesta.")))
        self.assertTrue(es_accidente("En llamas por un Tesla", (
            "Mientras transitaba por una autopista de California (Estados Unidos), un motociclista sufrió un grave "
            "siniestro luego de que un vehículo autónomo de Tesla, que circulaba sin conductor, invadiera "
            "inesperadamente el carril. Tras el impacto, el tanque se rompió y lo dejó envuelto en llamas.")))

    def test_emergencias_y_siniestros_que_no_son_accidente(self):
        for titulo, resumen in (
            ("Cómo saber si un carro usado estuvo inundado", ""),
            ("La Nissan Patrol resiste inundaciones y lluvias", ""),
            ("Avalancha de reservas para el nuevo BYD", ""),
            ("Así quedó la movilidad en Bogotá tras las lluvias", ""),
            ("Cómo reclamar un siniestro a la aseguradora", "El seguro cubre el siniestro si el carro tiene póliza."),
            ("Lynk & Co 07 GT se lanzó de un barranco para demostrar su seguridad",
             "La prueba buscaba mostrar la protección ante un accidente de gran impacto, con choques simulados."),
        ):
            self.assertFalse(es_accidente(titulo, resumen), titulo)


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
            {"Resumen - Aclaracion": largo + "Chery también habló. Así lo explicó Laura Méndez en entrevista.",
             "Menciones - Empresa": "Chery"},
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
                              ("iCAUR", None), ("Nissan", "Ricardo Toro"),
                              ("iCAUR", "Hernán Novoa"), ("Xpeng", "Zhao Deli")):
            if nombre:
                self.assertTrue(detecta_vocero(marca, "", f"Habló {nombre} sobre el año."), marca)
        # Solo por nombre, sin cargo: cuentan para su marca y no para otra.
        self.assertFalse(detecta_vocero("Chery", "", "Habló Hernán Novoa sobre el año."))
        self.assertFalse(detecta_vocero("Tesla", "", "Habló Zhao Deli sobre el año."))


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


class VocerosSinErroresTests(unittest.TestCase):
    """Calibración del 2026-11-22: cargo sin persona, empresas y huecos largos."""

    def test_representante_sin_persona_no_cuenta(self):
        casos = [
            ("Chery", "Chery, a través de su representante en Colombia, presentó el Tiggo 8."),
            ("Chery", "Auto Vardí, representante de Chery en Colombia, presentó el Tiggo 8."),
            ("Chery", "El representante de la marca Chery en el país anunció precios."),
            ("Nissan", "Los representantes de Nissan estuvieron en el evento."),
            ("Nissan", "Nissan y sus directivos celebraron los 10 años de la planta."),
            ("Toyota", "La Hilux, el líder de ventas de Toyota, se renueva."),
            ("Kia", "Los voceros de la marca Kia no respondieron."),
            ("Renault", "El jefe de taller de Renault revisó el carro."),
        ]
        for marca, texto in casos:
            self.assertFalse(detecta_vocero(marca, "", texto), texto)
            self.assertEqual(R["descubrir_voceros"](_df({"Resumen - Aclaracion": texto, "Menciones - Empresa": marca})),
                             {}, texto)

    def test_representante_con_persona_si_cuenta(self):
        texto = "Ana Ruiz, representante de Chery en Colombia, explicó la meta de ventas."
        self.assertTrue(detecta_vocero("Chery", "", texto))
        self.assertEqual(R["descubrir_voceros"](_df({"Resumen - Aclaracion": texto, "Menciones - Empresa": "Chery"})),
                         {"Chery": ["Ana Ruiz"]})
        self.assertTrue(detecta_vocero("Nissan", "", "Nissan, a través de su líder de producto, Pedro Gil, presentó la Frontier."))

    def test_cargo_lejos_de_la_marca_no_cuenta(self):
        casos = [
            ("Renault", "El presidente de Colombia visitó la planta de Renault en Envigado."),
            ("Renault", "El presidente de Colombia en la planta de Renault habló de empleo."),
            ("Nissan", "El presidente de Fenalco en el evento de Nissan habló de ventas."),
            ("Toyota", "El gerente del concesionario de Toyota fue capturado."),
            ("Nissan", "El director del noticiero habló del nuevo Nissan Kicks."),
        ]
        for marca, texto in casos:
            self.assertFalse(detecta_vocero(marca, "", texto), texto)

    def test_ex_cargos_no_cuentan(self):
        casos = [
            "Carlos Ghosn, expresidente de Nissan, habló desde Líbano.",
            "El ex presidente de Nissan Carlos Ghosn publicó un libro.",
            "El entonces presidente de Nissan fue arrestado en 2018.",
            "El antiguo CEO de Nissan criticó la fusión.",
        ]
        for texto in casos:
            self.assertFalse(detecta_vocero("Nissan", "", texto), texto)
            self.assertEqual(R["descubrir_voceros"](_df({"Resumen - Aclaracion": texto})), {}, texto)

    def test_vocero_conocido_con_cargo_en_otra_marca_no_cuenta(self):
        self.assertFalse(detecta_vocero("Nissan", "", "Carlos Vargas, gerente de mercadeo de Toyota, habló."))
        self.assertFalse(detecta_vocero("Nissan", "", "Carlos Vargas, exgerente de Nissan, abrió su empresa."))
        self.assertTrue(detecta_vocero("Nissan", "", "Carlos Vargas, gerente de mercadeo de Nissan, habló."))
        self.assertTrue(detecta_vocero("Nissan", "", "Carlos Vargas, gerente de la Fundación, habló de Nissan."))

    def test_cargos_largos_de_la_marca_si_cuentan(self):
        casos = [
            ("Renault", "Juan Camilo Vélez, presidente y director general de Renault-Sofasa, dio el balance."),
            ("Renault", "El presidente para Colombia y la región Andina de Renault habló del mercado."),
            ("Xpeng", "El vicepresidente del centro de diseño de Xpeng presentó el G6."),
            ("Nissan", "El director de producto Nissan Colombia explicó la e-Power."),
            ("Toyota", "El presidente de la automotriz japonesa Toyota anunció inversiones."),
            ("Toyota", "El presidente del consejo de administración de Toyota habló."),
            ("Kia", "Kia y su nuevo gerente general presentaron metas para 2027."),
            ("Hyundai", "LA GERENTE DE MERCADEO DE HYUNDAI COLOMBIA HABLÓ DEL AÑO."),
        ]
        for marca, texto in casos:
            self.assertTrue(detecta_vocero(marca, "", texto), texto)

    def test_no_descubre_empresas_ni_modelos(self):
        textos = {
            "Chery": "Grupo Vardi, representante de Chery, y Auto Vardí, distribuidor de Chery, crecen.",
            "Toyota": "Toyota Hilux, líder del segmento de Toyota, se renueva.",
            "Nissan": "Automotores Andinos, representante de Nissan en Nariño, abrió vitrina.",
        }
        for marca, texto in textos.items():
            self.assertEqual(R["descubrir_voceros"](_df({"Resumen - Aclaracion": texto, "Menciones - Empresa": marca})),
                             {}, texto)

    def test_casos_reales_de_la_base_competencia_nov_2026(self):
        # Sí son voceros:
        for marca, texto in (
            ("Chevrolet", "afirma Hermann Mahnke, director de experiencias y servicios digitales de GM Sudamérica."),
            ("iCAUR", "señaló Hernán Novoa, director de marca iCAUR en Colombia."),
            ("BYD", "su estilo de vida, Leidy Silva - Jefe de Planeación y Desarrollo deProducto de BYD - Motorysa."),
            ("Xpeng", "El fundador de Xpeng, Zhao Deli, ha explicado que ya se hicieron demostraciones."),
        ):
            self.assertTrue(detecta_vocero(marca, "", texto), texto)
        # No son voceros (empresa distribuidora o grupo sin persona):
        for marca, texto in (
            ("BYD", "Motorysa, que ya completa seis años como representante de BYD en Colombia, estará a cargo."),
            ("Jetour", "bajo la marca Karry a cargo de Autocom, también representante de Jetour, tiene un motor."),
            ("Renault", "Según los representantes de Renault-Sofasa, el modelo ya recibió adaptaciones."),
            ("Renault", "Al evento asistieron representantes del Gobierno Nacional y directivos de Sofasa."),
        ):
            self.assertFalse(detecta_vocero(marca, "", texto), texto)
            self.assertEqual(R["descubrir_voceros"](_df({"Resumen - Aclaracion": texto, "Menciones - Empresa": marca})),
                             {}, texto)

    def test_pule_nombres_descubiertos(self):
        df = _df(
            {"Resumen - Aclaracion": "Trump Elon Musk, CEO de Tesla, habló.", "Menciones - Empresa": "Tesla"},
            {"Resumen - Aclaracion": "Dato relevante Rafael Melo, gerente de ventas de Toyota, habló.",
             "Menciones - Empresa": "Toyota"},
            {"Resumen - Aclaracion": "Toyota Financial Services, representante de Toyota, ofrece créditos.",
             "Menciones - Empresa": "Toyota"},
        )
        self.assertEqual(R["descubrir_voceros"](df), {"Tesla": ["Elon Musk"], "Toyota": ["Rafael Melo"]})

    def test_directivo_en_singular_cuenta_sin_nombre(self):
        # Base Competencia 22 nov. 2026: "Colombia y Japón reactivan su acuerdo comercial".
        self.assertTrue(detecta_vocero("Toyota", "", "La otra reunión estratégica fue con un directivo de Toyota, "
                                                     "una compañía con cinco décadas en el país."))
        self.assertTrue(detecta_vocero("Nissan", "", "El directivo de Nissan habló de las ventas."))
        self.assertFalse(detecta_vocero("Nissan", "", "Nissan y sus directivos celebraron los 10 años."))
        self.assertFalse(detecta_vocero("Toyota", "", "La junta directiva de Toyota se reunió en Tokio."))
        self.assertFalse(detecta_vocero("Toyota", "", "El exdirectivo de Toyota abrió su empresa."))


class CherySinMencionTests(unittest.TestCase):
    def _valores(self, df):
        return list(zip(df["Tono"], df["Tema"], df[SUB]))

    def test_chery_sin_mencion_va_a_changan_deepal(self):
        out = aplicar(_df(
            {"Título": "Chery en el Salón", "Resumen - Aclaracion": "Deepal presentó el S07 en Bogotá.",
             "Menciones - Empresa": "Chery"},
            {"Resumen - Aclaracion": "Changan lanzó la CS55 Plus.", "Menciones - Empresa": "Chery"},
        ))
        self.assertEqual(self._valores(out), [("No aplica", "Changan/Deepal", "-")] * 2)

    def test_chery_con_mencion_no_cambia(self):
        out = aplicar(_df(
            {"Resumen - Aclaracion": "CHERY presentó su portafolio.", "Menciones - Empresa": "Chery"},
            {"Resumen - Aclaracion": "Cherry presentó su portafolio.", "Menciones - Empresa": "Chery"},
            {"Resumen - Aclaracion": "La Tiggo 7 Pro llega en versión híbrida.", "Menciones - Empresa": "Chery"},
            {"Resumen - Aclaracion": "El Omoda 5 y el Jaecoo 7 llegan a España.", "Menciones - Empresa": "Chery"},
            {"Resumen - Aclaracion": "JAECOO presentó su nueva SUV.", "Menciones - Empresa": "Chery"},
            {"Título": "Chery Tiggo 4", "Resumen - Aclaracion": "", "Menciones - Empresa": "Chery"},
            {"Título": "", "Resumen - Aclaracion": "", "Menciones - Empresa": "Chery"},
            {"Resumen - Aclaracion": "Deepal presentó el S07.", "Menciones - Empresa": "Nissan"},
        ))
        self.assertEqual(self._valores(out), [("Neutro", "Tecnológico", "Electrificación")] * 8)

    def test_chery_sin_mencion_gana_a_vocero_pero_no_a_accidentes_ni_jcl(self):
        out = aplicar(_df(
            {"Resumen - Aclaracion": "El CEO de Changan, Zhu Huarong, habló de metas.", "Menciones - Empresa": "Chery"},
            {"Título": "Accidente de tránsito en la vía al Llano", "Resumen - Aclaracion": "Un carro se volcó.",
             "Menciones - Empresa": "Chery"},
            {"Resumen - Aclaracion": "Juan Carlos López presentó el Deepal S05.", "Menciones - Empresa": "Chery"},
            {"Resumen - Aclaracion": "Deepal presentó el S07.", "Menciones - Empresa": "Chery",
             "is_duplicate": True, "Tono": "Duplicada", "Tema": "-", SUB: "-"},
        ))
        self.assertEqual(self._valores(out), [
            ("No aplica", "Changan/Deepal", "-"),
            ("No aplica", "Accidentes", "-"),
            ("No aplica", "Juan Carlos López", "-"),
            ("Duplicada", "-", "-"),
        ])


class ProyeccionesDelModeloTests(unittest.TestCase):
    """El PKL pone Proyecciones - Balances; sin vocero de la marca no vale (2026-11-22)."""
    AUTOFINANCIERA = (
        "Autofinanciera es una compañía con 55 años de trayectoria en Colombia que ofrece un modelo de ahorro "
        "programado como alternativa al crédito tradicional para adquirir vehículos, según explicó su "
        "vicepresidente, Wilman Martínez. La empresa cuenta con alianzas con marcas como Chevrolet, Renault y "
        "Kia, dispone de 42 sedes a nivel nacional y actualmente genera 800 empleos directos en el país.")

    def _fila(self, **extra):
        base = {"Resumen - Aclaracion": self.AUTOFINANCIERA, "Tono": "Positivo", "Tema": "Indefinido",
                SUB: "Proyecciones - Balances", "_subtema_alternativo": "Estadísticas - Mercado"}
        return {**base, **extra}

    def test_sin_vocero_toma_la_segunda_opcion_del_modelo(self):
        out = aplicar(_df(*[self._fila(**{"Menciones - Empresa": m}) for m in ("Chevrolet", "Renault", "Kia")]),
                      final_topic_map={"Estadísticas - Mercado": "Indefinido", "Electrificación": "Tecnológico"})
        self.assertEqual(list(zip(out["Tono"], out["Tema"], out[SUB])),
                         [("Positivo", "Indefinido", "Estadísticas - Mercado")] * 3)
        for m in ("Chevrolet", "Renault", "Kia"):
            self.assertFalse(detecta_vocero(m, "", self.AUTOFINANCIERA), m)

    def test_tema_sale_del_mapa_y_sin_alternativa_queda_mencion(self):
        out = aplicar(_df(
            self._fila(**{"Menciones - Empresa": "Kia", "_subtema_alternativo": "Electrificación"}),
            self._fila(**{"Menciones - Empresa": "Kia", "_subtema_alternativo": ""}),
            {**self._fila(**{"Menciones - Empresa": "Kia"}), SUB: "Proyecciones -Balances"},
        ), final_topic_map={"Electrificación": "Tecnológico", "Estadísticas - Mercado": "Indefinido"})
        self.assertEqual(list(zip(out["Tema"], out[SUB])), [
            ("Tecnológico", "Electrificación"), ("Indefinido", "Mención"), ("Indefinido", "Estadísticas - Mercado")])

    def test_con_vocero_de_la_marca_se_queda(self):
        resumen = "Jorge Neira, director general de Kia Colombia, presentó el balance del año."
        out = aplicar(_df(self._fila(**{"Menciones - Empresa": "Kia", "Resumen - Aclaracion": resumen})))
        self.assertEqual(out[SUB].tolist(), ["Proyecciones - Balances"])

    def test_juan_carlos_lopez_y_duplicadas_no_cambian(self):
        out = aplicar(_df(
            self._fila(**{"Menciones - Empresa": "Changan", "Resumen - Aclaracion": "Juan Carlos López habló."}),
            {**self._fila(**{"Menciones - Empresa": "Kia"}), "is_duplicate": True, "Tono": "Duplicada",
             "Tema": "-", SUB: "-"},
        ))
        self.assertEqual(list(zip(out["Tono"], out[SUB])),
                         [("Positivo", "Proyecciones - Balances"), ("Duplicada", "-")])

    def test_segunda_opcion_tema(self):
        import numpy as np
        fuente = _APP.read_text(encoding="utf-8")
        arbol = ast.parse(fuente)
        fn = next(n for n in arbol.body if isinstance(n, ast.FunctionDef) and n.name == "segunda_opcion_tema")
        ns = {"np": np, "pd": pd, "SUBTEMAS_PROYECCIONES": R["SUBTEMAS_PROYECCIONES"]}
        exec(compile(ast.Module([fn], []), str(_APP), "exec"), ns)

        class Modelo:
            classes_ = np.array(["Eventos", "Proyecciones - Balances", "Electrificación"])
            def decision_function(self, textos):
                return np.array([[0.1, 0.9, 0.3], [0.5, 0.8, 0.2]])[:len(textos)]

        out = ns["segunda_opcion_tema"](Modelo(), pd.Series(["a", "b"], index=[7, 9]))
        self.assertEqual(out.to_dict(), {7: "Electrificación", 9: "Eventos"})
        self.assertEqual(ns["segunda_opcion_tema"](object(), pd.Series(["a"])).tolist(), ["Mención"])


class ListaDeMarcasTests(unittest.TestCase):
    def test_cargo_en_varias_marcas(self):
        texto = "afirmó Cristhian Camelo, director de Mercadeo de AVATR, Deepal, NEVO y Changan Colombia."
        for marca in ("AVATR", "Deepal", "Changan"):
            self.assertTrue(detecta_vocero(marca, "", texto), marca)
        self.assertFalse(detecta_vocero("Nissan", "", texto))
        otra = "afirmó Ana Ruiz, directora de Mercadeo de AVATR, Deepal, NEVO y Changan Colombia."
        for marca in ("AVATR", "Deepal", "Changan"):
            self.assertTrue(detecta_vocero(marca, "", otra), marca)
        self.assertFalse(detecta_vocero("Nissan", "", "El CEO de Tesla y Nissan anunciaron una alianza."))

class AliasMarcaTests(unittest.TestCase):
    def test_jetour_nissan_pasa_a_jetour(self):
        fuente = _APP.read_text(encoding="utf-8")
        self.assertIn("'jetour - nissan': 'Jetour'", fuente)


if __name__ == "__main__":
    unittest.main()
