# Dossier Intelligence · Procesador

App Streamlit para limpiar y clasificar dossiers de prensa (Nissan / Chery).

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud

En **Advanced settings** de la app, usa **Python 3.11 o 3.12**. Evita **Python 3.14**: el runtime aún es inestable (openpyxl / defusedxml y otras dependencias).

Algunos `.xlsx` de Brandwatch u otros exports traen XML de relaciones no estándar. La app intenta abrirlos con openpyxl, sanitiza el paquete si hace falta y, si sigue fallando, lee los valores con `python-calamine`.
