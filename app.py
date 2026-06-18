import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
import datetime
import io
import re
import html
import joblib
import numpy as np
import nltk
from pathlib import Path
from difflib import SequenceMatcher
from collections import defaultdict

# ──────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dossier Intelligence · Procesador",
    page_icon="📰",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
    .stApp { background-color: #F7F9F8; }
    #MainMenu, footer, header { visibility: hidden; }
    .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1100px; }

    .app-header {
        background: linear-gradient(135deg, #0D4F3C 0%, #1A7A5E 60%, #22A37A 100%);
        border-radius: 12px; padding: 1.1rem 1.8rem; margin-bottom: 1.2rem;
        position: relative; overflow: hidden; display: flex; align-items: center; gap: 1.2rem;
    }
    .app-header::before {
        content: ''; position: absolute; top: -30px; right: -30px;
        width: 120px; height: 120px; background: rgba(255,255,255,0.06); border-radius: 50%;
    }
    .app-header p { color: rgba(255,255,255,0.9); font-size: 0.92rem; font-weight: 500; margin: 0; }
    .app-header .badge {
        display: inline-block; background: rgba(255,255,255,0.15); color: #fff;
        font-size: 0.7rem; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase;
        padding: 0.2rem 0.6rem; border-radius: 20px; border: 1px solid rgba(255,255,255,0.25);
        white-space: nowrap; flex-shrink: 0;
    }

    .card {
        background: #FFFFFF; border: 1px solid #E8EFEC; border-radius: 12px;
        padding: 1.6rem 1.8rem; margin-bottom: 1.2rem;
        box-shadow: 0 1px 4px rgba(13,79,60,0.06);
    }
    .card-title {
        font-size: 0.8rem; font-weight: 600; letter-spacing: 0.1em; text-transform: uppercase;
        color: #1A7A5E; margin-bottom: 1rem;
    }

    .metrics-row { display: flex; gap: 1rem; margin: 1.5rem 0; }
    .metric-card {
        flex: 1; background: #FFFFFF; border: 1px solid #E8EFEC; border-radius: 12px;
        padding: 1.4rem 1.6rem; text-align: center; box-shadow: 0 1px 4px rgba(13,79,60,0.06);
    }
    .metric-card .metric-value {
        font-size: 2.2rem; font-weight: 700; color: #0D4F3C; line-height: 1;
        margin-bottom: 0.4rem; font-family: 'DM Mono', monospace;
    }
    .metric-card .metric-label { font-size: 0.78rem; color: #6B8F82; font-weight: 500; letter-spacing: 0.03em; }
    .metric-card.accent .metric-value { color: #1A7A5E; }
    .metric-card.muted  .metric-value { color: #9BB5AC; }

    .success-banner {
        background: linear-gradient(135deg, #0D4F3C, #1A7A5E); border-radius: 12px;
        padding: 1.2rem 1.8rem; display: flex; align-items: center; gap: 1rem;
        margin: 1.5rem 0; box-shadow: 0 2px 12px rgba(13,79,60,0.2);
    }
    .success-banner .icon { font-size: 1.6rem; line-height: 1; }
    .success-banner .text strong { display: block; color: #FFFFFF; font-size: 1rem; font-weight: 700; margin-bottom: 0.15rem; }
    .success-banner .text span { color: rgba(255,255,255,0.72); font-size: 0.85rem; }

    .step { display: flex; align-items: flex-start; gap: 1rem; margin-bottom: 0.8rem; }
    .step-num {
        min-width: 28px; height: 28px; background: #1A7A5E; color: white; border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        font-size: 0.78rem; font-weight: 700; margin-top: 1px;
    }
    .step-text { color: #2D5A4A; font-size: 0.9rem; line-height: 1.6; }

    [data-testid="stFileUploader"] {
        background: #FFFFFF; border: 2px dashed #B2D4C8; border-radius: 12px;
        padding: 0.5rem; transition: border-color 0.2s;
    }
    [data-testid="stFileUploader"]:hover { border-color: #1A7A5E; }

    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #0D4F3C, #1A7A5E); color: white;
        border: none; border-radius: 10px; padding: 0.7rem 2rem;
        font-family: 'DM Sans', sans-serif; font-size: 0.95rem; font-weight: 600;
        letter-spacing: 0.01em; transition: all 0.2s;
        box-shadow: 0 2px 8px rgba(13,79,60,0.25); width: 100%; height: 48px;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #0A3D2E, #156A50);
        box-shadow: 0 4px 14px rgba(13,79,60,0.35); transform: translateY(-1px);
    }
    .stButton > button[kind="primary"]:disabled { background: #C5D9D4; box-shadow: none; transform: none; }

    .stDownloadButton > button {
        background: #FFFFFF; color: #0D4F3C; border: 2px solid #1A7A5E;
        border-radius: 10px; padding: 0.65rem 1.8rem;
        font-family: 'DM Sans', sans-serif; font-size: 0.9rem; font-weight: 600;
        transition: all 0.2s; width: 100%; height: 48px;
    }
    .stDownloadButton > button:hover {
        background: #F0F7F4; box-shadow: 0 2px 10px rgba(13,79,60,0.15); transform: translateY(-1px);
    }

    .stWarning { background: #FFF8ED; border: 1px solid #F5C97A; border-radius: 10px; }
    .stError   { background: #FEF2F2; border: 1px solid #FCA5A5; border-radius: 10px; }
    .stProgress > div > div > div { background: linear-gradient(90deg, #1A7A5E, #22A37A); border-radius: 4px; }
    .streamlit-expanderHeader { background: #F7F9F8; border-radius: 8px; font-weight: 500; color: #0D4F3C; }
    [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; border: 1px solid #E8EFEC; }
    hr { border: none; border-top: 1px solid #E8EFEC; margin: 1.5rem 0; }
    .stSpinner > div { border-top-color: #1A7A5E !important; }

    .file-status {
        display: flex; align-items: center; gap: 0.6rem; padding: 0.7rem 1rem;
        border-radius: 8px; font-size: 0.87rem; font-weight: 500; margin-top: 0.5rem;
    }
    .file-status.ok      { background: #F0F7F4; color: #0D4F3C; border: 1px solid #B2D4C8; }
    .file-status.missing { background: #FFF8ED; color: #92400E; border: 1px solid #F5C97A; }

    .results-header {
        font-size: 1.1rem; font-weight: 700; color: #0D4F3C;
        margin: 1.8rem 0 1rem 0; padding-bottom: 0.5rem; border-bottom: 2px solid #E8EFEC;
    }
</style>
""", unsafe_allow_html=True)

# NLTK stopwords
try:
    nltk.data.find('corpora/stopwords')
except LookupError:
    with st.spinner("Descargando recursos de lenguaje por primera vez..."):
        nltk.download('stopwords')

# ──────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ──────────────────────────────────────────────────────────────────────────────
SIMILARITY_THRESHOLD_TITULOS = 0.93

FINAL_ORDER = [
    "ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa",
    "Región", "Título", "Autor - Conductor", "Nro. Pagina", "Dimensión",
    "Duración - Nro. Caracteres", "CPE", "Tier", "Audiencia", "Tono", "Tema",
    "Temas Generales - Tema", "Resumen - Aclaracion", "Link Nota",
    "Link (Streaming - Imagen)", "Menciones - Empresa",
]

TIPO_MEDIO_MAP = {
    'online': 'Internet', 'internet': 'Internet',
    'diario': 'Prensa',
    'am': 'Radio', 'fm': 'Radio', 'radio': 'Radio',
    'aire': 'Televisión', 'cable': 'Televisión', 'tv': 'Televisión',
    'television': 'Televisión', 'televisión': 'Televisión',
    'revista': 'Revistas', 'revistas': 'Revistas',
}

# ──────────────────────────────────────────────────────────────────────────────
# CARGA DE MODELOS PKL
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_resource
def load_ml_models():
    try:
        sentiment_pipeline = joblib.load('pipeline_sentimiento.pkl')
        topic_pipeline     = joblib.load('pipeline_tema.pkl')
        return sentiment_pipeline, topic_pipeline
    except FileNotFoundError as e:
        st.error(
            f"**Error Crítico:** No se encontró `{e.filename}`. "
            "Asegúrate de que `pipeline_sentimiento.pkl` y `pipeline_tema.pkl` "
            "estén en la misma carpeta que esta app."
        )
        st.stop()

# ──────────────────────────────────────────────────────────────────────────────
# CARGA AUTOMÁTICA DE Configuracion.xlsx
# ──────────────────────────────────────────────────────────────────────────────
def _find_config_file():
    candidates = ["Configuracion.xlsx", "configuracion.xlsx", "Config.xlsx", "config.xlsx"]
    for name in candidates:
        p = Path(name)
        if p.exists():
            return p
    base = Path(__file__).parent
    for f in base.iterdir():
        if f.suffix.lower() == '.xlsx' and 'config' in f.stem.lower():
            return f
    return None

@st.cache_data
def load_config(_path):
    sheets = pd.read_excel(_path, sheet_name=None, engine='openpyxl')
    region_map = pd.Series(
        sheets['Regiones'].iloc[:, 1].values,
        index=sheets['Regiones'].iloc[:, 0].astype(str).str.lower().str.strip()
    ).to_dict()
    internet_map = pd.Series(
        sheets['Internet'].iloc[:, 1].values,
        index=sheets['Internet'].iloc[:, 0].astype(str).str.lower().str.strip()
    ).to_dict()
    mention_map = pd.Series(
        sheets['Menciones'].iloc[:, 1].values,
        index=sheets['Menciones'].iloc[:, 0].astype(str).str.strip()
    ).to_dict()
    final_topic_map = pd.Series(
        sheets['Mapa_Temas'].iloc[:, 1].values,
        index=sheets['Mapa_Temas'].iloc[:, 0].astype(str).str.strip()
    ).to_dict()
    return region_map, internet_map, mention_map, final_topic_map

# ──────────────────────────────────────────────────────────────────────────────
# UTILIDADES DE TEXTO
# ──────────────────────────────────────────────────────────────────────────────
def convert_html_entities(text):
    if not isinstance(text, str):
        return text
    text = html.unescape(text)
    for bad, good in {
        '\u201c': '"', '\u201d': '"', '\u2018': "'", '\u2019': "'",
        'Â': '', 'â': '', '€': '', '™': '', '\x9d': '', '\xa0': ' ',
    }.items():
        text = text.replace(bad, good)
    return text

def clean_text(text):
    if not isinstance(text, str):
        return text
    return convert_html_entities(text).strip()

def clean_title(title):
    if not isinstance(title, str):
        return ""
    return convert_html_entities(title)

def clean_title_for_output(title):
    if not isinstance(title, str):
        return ""
    title = convert_html_entities(title)
    title = title.replace('\n', ' ').replace('\r', ' ')
    title = re.sub(r'\s+\|.*$', '', title)
    title = re.sub(r'\|\s+.*$', '', title)
    title = re.sub(r'\s+-\s+.*$', '', title)
    return title.strip()

def normalize_title_for_comparison(title):
    if not isinstance(title, str):
        return ""
    tmp = re.split(r'\s*[:|-]\s*', title, 1)
    cleaned = re.sub(r'\W+', ' ', tmp[0]).lower().strip()
    return cleaned

def corregir_resumen(text):
    if not isinstance(text, str):
        return text
    text = convert_html_entities(text)
    text = re.sub(r'(<br\s*/?>|\[\.\.\.\]|\s+)', ' ', text).strip()
    m = re.search(r'[A-Z]', text)
    if m:
        text = text[m.start():]
    if text and not text.endswith('...'):
        text = text.rstrip('.') + '...'
    return text

def clean_cuerpo(text):
    if not isinstance(text, str) or not text.strip():
        return text
    text = convert_html_entities(text)
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    return text.strip()

def preprocess_text_for_topic(text):
    if not isinstance(text, str):
        return ""
    try:
        from nltk.corpus import stopwords
        stop_words = set(stopwords.words('spanish'))
    except Exception:
        stop_words = set([
            'de','la','que','el','en','y','a','los','del','se','las','por',
            'un','para','con','no','una','su','al','lo'
        ])
    tokens = re.findall(r'\b\w+\b', text.lower())
    return " ".join(t for t in tokens if t not in stop_words)

def parse_numeric(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        if isinstance(val, float) and val.is_integer():
            return int(val)
        return val
    s = str(val).strip()
    if not s:
        return None
    try:
        if ',' in s and '.' not in s:
            parts = s.split(',')
            s = s.replace(',', '') if len(parts[-1]) == 3 else s.replace(',', '.')
        f = float(s.replace(',', ''))
        return int(f) if f.is_integer() else f
    except ValueError:
        return None

def _normalizar_url(url):
    if not url:
        return ""
    url = str(url).strip().lower()
    url = re.sub(r'^https?://', '', url)
    url = re.sub(r'^www\.', '', url)
    url = url.rstrip('/')
    return url

# ──────────────────────────────────────────────────────────────────────────────
# LECTURA Y NORMALIZACIÓN DEL DOSSIER (formato nuevo - documento 3)
# ──────────────────────────────────────────────────────────────────────────────
def _extract_cell_link(cell):
    """Extrae URL del hipervínculo de una celda openpyxl."""
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    return None

def read_and_normalize_dossier(wb_sheet, region_map, internet_map):
    """
    Lee la hoja activa con openpyxl preservando hipervínculos,
    normaliza tipos de medio, mapea región e internet, y
    devuelve una lista de dicts (una por fila, sin expansión por `;` aún).
    """
    headers = [cell.value for cell in wb_sheet[1] if cell.value is not None]
    raw_rows = []

    for row in wb_sheet.iter_rows(min_row=2):
        if all(c.value is None for c in row):
            continue
        row_data = {}
        for i, h in enumerate(headers):
            if i >= len(row):
                row_data[h] = None
                continue
            cell = row[i]
            url = _extract_cell_link(cell)
            row_data[h] = {"value": cell.value or "Link", "url": url} if url else cell.value
        raw_rows.append(row_data)

    df = pd.DataFrame(raw_rows)

    # ── Tipo de Medio ──────────────────────────────────────────────────────────
    if 'Tipo de Medio' in df.columns:
        df['Tipo de Medio'] = (
            df['Tipo de Medio'].astype(str).str.lower().str.strip()
            .map(TIPO_MEDIO_MAP)
            .fillna(df['Tipo de Medio'].astype(str).str.strip())
        )
    else:
        df['Tipo de Medio'] = 'Otro'

    is_av       = df['Tipo de Medio'].isin(['Radio', 'Televisión'])
    is_grafica  = df['Tipo de Medio'].isin(['Prensa', 'Internet', 'Revistas'])
    is_internet = df['Tipo de Medio'] == 'Internet'
    is_print    = df['Tipo de Medio'].isin(['Prensa', 'Revistas'])
    is_bcast    = df['Tipo de Medio'].isin(['Radio', 'Televisión'])

    # ── Región (antes de cambiar Medio) ───────────────────────────────────────
    if 'Medio' in df.columns:
        df['Región'] = (
            df['Medio'].astype(str).str.lower().str.strip()
            .map(region_map)
            .fillna('N/A')
        )
    else:
        df['Medio'] = 'N/A'
        df['Región'] = 'N/A'

    # ── Mapeo internet sobre Medio ─────────────────────────────────────────────
    if 'Medio' in df.columns:
        df.loc[is_internet, 'Medio'] = (
            df.loc[is_internet, 'Medio']
            .astype(str).str.lower().str.strip()
            .map(internet_map)
            .fillna(df.loc[is_internet, 'Medio'])
        )

    # ── Campos de identidad / texto ────────────────────────────────────────────
    df['ID Noticia']         = df.get('NoticiaId', df.get('ID Noticia', pd.Series(dtype=str)))
    df['Fecha']              = pd.to_datetime(df.get('Fecha', pd.Series(dtype=str)), dayfirst=True, errors='coerce').dt.normalize()
    df['Hora']               = df.get('Hora', pd.Series(dtype=str))
    df['Sección - Programa'] = df.get('Sección - Programa', pd.Series(dtype=str)).astype(str).apply(clean_text)

    titulo_col = 'Título' if 'Título' in df.columns else 'Titulo'
    df['Título'] = df.get(titulo_col, pd.Series(dtype=str)).astype(str).apply(clean_title)

    df['Autor - Conductor']       = df.get('Autor - Conductor', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Nro. Pagina']             = df.get('Nro. Pagina', pd.Series(dtype=str))

    dim_col = 'Dimensioncm2' if 'Dimensioncm2' in df.columns else 'Dimensión'
    df['Dimensión']               = df.get(dim_col, pd.Series(dtype=str))
    df['Duración - Nro. Caracteres'] = df.get('Duración - Nro. Caracteres', pd.Series(dtype=str))

    # ── Dimensión ↔ Duración para AV ──────────────────────────────────────────
    df.loc[is_av, 'Dimensión']                 = df.loc[is_av, 'Duración - Nro. Caracteres']
    df.loc[is_av, 'Duración - Nro. Caracteres'] = 0

    # ── CPE ────────────────────────────────────────────────────────────────────
    cpe_av      = df.get('CPE',          pd.Series([np.nan] * len(df)))
    cpe_grafica = df.get('Valor de Nota', pd.Series([np.nan] * len(df)))
    df['CPE']   = np.where(is_av, cpe_av, np.where(is_grafica, cpe_grafica, np.nan))

    df['Tier']     = df.get('Tier',     pd.Series(dtype=str))
    df['Audiencia'] = df.get('Audiencia', pd.Series(dtype=str))
    df['Tono']     = ''
    df['Tema']     = ''
    df['Temas Generales - Tema'] = ''

    # ── Resumen ────────────────────────────────────────────────────────────────
    cuerpo_col = 'CuerpoEs' if 'CuerpoEs' in df.columns else 'Resumen - Aclaracion'
    cuerpo_cleaned = df.get(cuerpo_col, pd.Series([''] * len(df))).astype(str).apply(clean_cuerpo)
    df['Resumen - Aclaracion'] = cuerpo_cleaned.apply(corregir_resumen)

    # ── Links ──────────────────────────────────────────────────────────────────
    url_nota_av    = df.get('URL Nota AV',            df.get('Link Nota AV',   pd.Series([''] * len(df))))
    url_streaming  = df.get('URL (Streaming - Imagen)', pd.Series([''] * len(df)))
    url_nota_raw   = df.get('URL Nota',                pd.Series([''] * len(df)))

    link_nota_final    = []
    link_stream_final  = []

    for val_av, val_str, val_url_nota, av_row, int_row, print_row, bcast_row in zip(
        url_nota_av, url_streaming, url_nota_raw, is_av, is_internet, is_print, is_bcast
    ):
        # Link Nota
        if av_row:
            raw_url = val_av.get("url", "") if isinstance(val_av, dict) else str(val_av or "")
            link_nota_final.append({"value": "Link", "url": raw_url or None})
        else:
            if isinstance(val_str, dict):
                link_nota_final.append(val_str)
            else:
                link_nota_final.append({"value": "Link", "url": str(val_str) if val_str else None})

        # Link (Streaming - Imagen) — solo Internet
        if int_row:
            if isinstance(val_url_nota, dict):
                link_stream_final.append(val_url_nota)
            else:
                link_stream_final.append({"value": "Link", "url": str(val_url_nota) if val_url_nota else None})
        else:
            link_stream_final.append(None)

    df['Link Nota']                = link_nota_final
    df['Link (Streaming - Imagen)'] = link_stream_final

    # ── Menciones ──────────────────────────────────────────────────────────────
    menciones_av      = df.get('Menciones - Empresa', pd.Series([''] * len(df))).fillna('').astype(str).apply(clean_text)
    menciones_grafica = df.get('Empresa rel.',        pd.Series([''] * len(df))).fillna('').astype(str).apply(clean_text)
    df['Menciones - Empresa'] = np.where(is_av, menciones_av, np.where(is_grafica, menciones_grafica, menciones_av))

    return df

# ──────────────────────────────────────────────────────────────────────────────
# EXPANSIÓN POR ; EN MENCIONES
# ──────────────────────────────────────────────────────────────────────────────
def expand_menciones(df):
    rows_expanded = []
    for _, row in df.iterrows():
        menciones = [m.strip() for m in str(row['Menciones - Empresa']).split(';') if m.strip()]
        if not menciones:
            d = row.to_dict()
            d['Menciones - Empresa'] = ''
            rows_expanded.append(d)
        else:
            for m in menciones:
                d = row.to_dict()
                d['Menciones - Empresa'] = m
                rows_expanded.append(d)
    return pd.DataFrame(rows_expanded).reset_index(drop=True)

# ──────────────────────────────────────────────────────────────────────────────
# MAPEO DE MENCIONES (Configuracion.xlsx)
# ──────────────────────────────────────────────────────────────────────────────
def apply_mention_map(df, mention_map):
    if 'Menciones - Empresa' in df.columns:
        df['Menciones - Empresa'] = (
            df['Menciones - Empresa'].astype(str).str.strip()
            .map(mention_map)
            .fillna(df['Menciones - Empresa'])
        )
    return df

# ──────────────────────────────────────────────────────────────────────────────
# DETECCIÓN DE DUPLICADOS (fusión de ambas apps)
# ──────────────────────────────────────────────────────────────────────────────
def _normalize_url_for_dedup(url):
    if not isinstance(url, str):
        return ""
    url = url.strip().lower().rstrip('/')
    if not url.startswith('http'):
        return ""
    url = re.sub(r'^(https?://)www\.', r'\1', url)
    url = re.split(r'[?#]', url)[0].rstrip('/')
    return url

def _get_cell_url(val):
    if isinstance(val, dict):
        return val.get('url') or ''
    return str(val) if val else ''

def _title_quality(title):
    if not isinstance(title, str):
        return -999
    score = 100
    score -= len(re.findall(r'&[#\w]+;', title)) * 10
    score -= title.count('??') * 5
    if len(title) > 250: score -= 5
    if len(title) < 15:  score -= 20
    if '\n' in title:    score -= 15
    if '|' in title:     score -= 5
    return score

def detect_duplicates(df):
    """
    Marca duplicados en columna 'is_duplicate'.
    Reglas por tipo de medio (igual que dossier_utils original).
    """
    df = df.copy().reset_index(drop=True)
    df['_orig_idx']     = df.index
    df['_title_quality'] = df['Título'].apply(_title_quality)
    df['is_duplicate']  = False

    # Ordenar: mejor calidad primero → queda como original
    df.sort_values(
        by=['_title_quality', 'Fecha', '_orig_idx'],
        ascending=[False, True, True],
        inplace=True, na_position='last'
    )

    seen_url      = {}   # (url_norm, mencion) → orig_idx
    seen_stream   = {}   # (stream_url_norm, mencion) → orig_idx
    seen_bcast    = {}   # (mencion, medio, hora) → orig_idx
    title_buckets = defaultdict(list)  # (medio, mencion) → [idx]

    for pos, row in df.iterrows():
        if df.at[pos, 'is_duplicate']:
            continue

        tipo    = str(row.get('Tipo de Medio', ''))
        mencion = str(row.get('Menciones - Empresa', '')).strip()
        medio   = str(row.get('Medio', '')).strip().lower()

        # ── URL Streaming duplicada ────────────────────────────────────────────
        stream_url = _get_cell_url(row.get('Link (Streaming - Imagen)'))
        if stream_url and mencion:
            sn = _normalizar_url(stream_url)
            if sn:
                sk = (sn, mencion)
                if sk in seen_stream:
                    df.at[pos, 'is_duplicate'] = True
                    continue
                seen_stream[sk] = pos

        if tipo == 'Internet':
            link_url = _get_cell_url(row.get('Link Nota'))
            if link_url and mencion:
                un = _normalize_url_for_dedup(link_url)
                k  = (un, mencion)
                if k in seen_url:
                    df.at[pos, 'is_duplicate'] = True
                    continue
                seen_url[k] = pos
            title_buckets[(medio, mencion)].append(pos)

        elif tipo in ('Radio', 'Televisión'):
            hora = str(row.get('Hora', '')).strip()
            if mencion and medio and hora:
                k = (mencion, medio, hora)
                if k in seen_bcast:
                    df.at[pos, 'is_duplicate'] = True
                    continue
                seen_bcast[k] = pos
            else:
                title_buckets[(medio, mencion)].append(pos)

        else:
            title_buckets[(medio, mencion)].append(pos)

    # ── Comparación de títulos dentro de cada (medio, mención) ────────────────
    for idxs in title_buckets.values():
        if len(idxs) < 2:
            continue
        for i in range(len(idxs)):
            for j in range(i + 1, len(idxs)):
                a, b = idxs[i], idxs[j]
                if df.at[a, 'is_duplicate'] or df.at[b, 'is_duplicate']:
                    continue
                ta = normalize_title_for_comparison(df.at[a, 'Título'])
                tb = normalize_title_for_comparison(df.at[b, 'Título'])
                if not ta or not tb:
                    continue
                # Ventana de fecha para Internet
                tipo_a = str(df.at[a, 'Tipo de Medio'])
                if tipo_a == 'Internet':
                    fa = df.at[a, 'Fecha']
                    fb = df.at[b, 'Fecha']
                    if pd.notna(fa) and pd.notna(fb) and abs((fa - fb).days) > 1:
                        continue
                else:
                    fa = df.at[a, 'Fecha']
                    fb = df.at[b, 'Fecha']
                    if pd.notna(fa) and pd.notna(fb) and fa.date() != fb.date():
                        continue

                sim = SequenceMatcher(None, ta, tb).ratio()
                if ta in tb or tb in ta or sim >= SIMILARITY_THRESHOLD_TITULOS:
                    # Marca el de menor calidad como duplicado
                    qa = df.at[a, '_title_quality']
                    qb = df.at[b, '_title_quality']
                    if qa >= qb:
                        df.at[b, 'is_duplicate'] = True
                    else:
                        df.at[a, 'is_duplicate'] = True

    # Restaurar orden original
    df.sort_values('_orig_idx', inplace=True)
    df.drop(columns=['_orig_idx', '_title_quality'], inplace=True)
    return df.reset_index(drop=True)

# ──────────────────────────────────────────────────────────────────────────────
# CLASIFICACIÓN PKL (Tono + Tema)
# ──────────────────────────────────────────────────────────────────────────────
LABEL_MAP_TONO = {1: 'Positivo', 0: 'Neutro', -1: 'Negativo',
                  '1': 'Positivo', '0': 'Neutro', '-1': 'Negativo'}

def classify_with_pkl(df, sentiment_pipeline, topic_pipeline, final_topic_map, progress_bar):
    """
    Aplica los pipelines PKL solo a noticias únicas.
    Homogeneiza temas por título similar y aplica Mapa_Temas.
    """
    df = df.copy()
    mask_unique = ~df['is_duplicate']
    df_valid = df[mask_unique].copy()

    if df_valid.empty:
        df['Tono'] = ''
        df['Temas Generales - Tema'] = ''
        df['Tema'] = ''
        return df

    df_valid['_texto_ia'] = (
        df_valid['Título'].fillna('') + ' ' + df_valid['Resumen - Aclaracion'].fillna('')
    )

    progress_bar.progress(60, text="Paso 5 / 7 — Clasificando tono con modelo PKL...")
    preds_sent = sentiment_pipeline.predict(df_valid['_texto_ia'])
    df_valid['Tono'] = [LABEL_MAP_TONO.get(p, str(p).title()) for p in preds_sent]

    progress_bar.progress(72, text="Paso 6 / 7 — Clasificando temas con modelo PKL...")
    df_valid['_resumen_procesado'] = df_valid['_texto_ia'].apply(preprocess_text_for_topic)
    df_valid['Temas Generales - Tema'] = topic_pipeline.predict(df_valid['_resumen_procesado'])

    # ── Homogeneización de temas por título normalizado ────────────────────────
    df_valid['_titulo_norm'] = df_valid['Título'].apply(normalize_title_for_comparison)
    homo = df_valid.groupby('_titulo_norm')['Temas Generales - Tema'].transform(
        lambda x: x.mode()[0] if not x.mode().empty else x
    )
    df_valid['Temas Generales - Tema'] = homo

    # ── Mapa_Temas ─────────────────────────────────────────────────────────────
    df_valid['Tema'] = (
        df_valid['Temas Generales - Tema'].astype(str).str.strip()
        .map(final_topic_map)
        .fillna('Indefinido')
    )

    df_valid.drop(columns=['_texto_ia', '_resumen_procesado', '_titulo_norm'], inplace=True)

    # Volcar a df original
    df.update(df_valid[['Tono', 'Temas Generales - Tema', 'Tema']])

    # Marcar duplicadas
    mask_dup = df['is_duplicate']
    if mask_dup.any():
        df.loc[mask_dup, 'Tono']                    = 'Duplicada'
        df.loc[mask_dup, 'Temas Generales - Tema']  = '-'
        df.loc[mask_dup, 'Tema']                    = '-'

    return df

# ──────────────────────────────────────────────────────────────────────────────
# GENERACIÓN DEL EXCEL DE SALIDA
# ──────────────────────────────────────────────────────────────────────────────
def to_excel(df):
    output = io.BytesIO()
    cols   = [c for c in FINAL_ORDER if c in df.columns]
    df_out = df[cols].copy()

    with pd.ExcelWriter(output, engine='xlsxwriter',
                        datetime_format='dd/mm/yyyy', date_format='dd/mm/yyyy') as writer:
        # Escribir con openpyxl para conservar hipervínculos correctamente
        pass  # usaremos openpyxl directamente abajo

    # ── Construir con openpyxl para hipervínculos ──────────────────────────────
    wb  = Workbook()
    ws  = wb.active
    ws.title = 'Resultado'

    font_link   = Font(color='0563C1', underline='single')
    font_header = Font(bold=True)
    align_left  = Alignment(horizontal='left')
    date_fmt    = 'DD/MM/YYYY'
    num_fmt_int = '#,##0'

    NUM_COLS = {'ID Noticia', 'Nro. Pagina', 'Dimensión', 'Duración - Nro. Caracteres',
                'CPE', 'Tier', 'Audiencia'}

    ws.append(cols)
    for i, col_name in enumerate(cols, start=1):
        ws.cell(row=1, column=i).font = font_header

    for _, row in df_out.iterrows():
        out_row = []
        link_map = {}  # col_idx → url

        for ci, col_name in enumerate(cols, start=1):
            val = row[col_name]

            if col_name == 'Fecha':
                if pd.notna(val):
                    cv = val.to_pydatetime() if isinstance(val, pd.Timestamp) else val
                else:
                    cv = None

            elif col_name in NUM_COLS:
                cv = parse_numeric(val)

            elif isinstance(val, dict) and 'url' in val:
                url = val.get('url')
                cv  = val.get('value', 'Link')
                if url:
                    link_map[ci] = url

            elif isinstance(val, str) and val.startswith('http'):
                cv = 'Link'
                link_map[ci] = val

            else:
                cv = None if (isinstance(val, float) and np.isnan(val)) else val
                if cv is not None and not isinstance(cv, (int, float, datetime.datetime, datetime.date)):
                    cv = str(cv)

            out_row.append(cv)

        ws.append(out_row)
        cur_row = ws.max_row

        # Fecha
        date_idx = cols.index('Fecha') + 1 if 'Fecha' in cols else None
        if date_idx:
            cell = ws.cell(row=cur_row, column=date_idx)
            if isinstance(cell.value, (datetime.datetime, datetime.date)):
                cell.number_format = date_fmt

        # CPE sin notación científica para AV
        if 'CPE' in cols and 'Tipo de Medio' in cols:
            cpe_idx  = cols.index('CPE') + 1
            tipo_idx = cols.index('Tipo de Medio') + 1
            tipo_val = ws.cell(row=cur_row, column=tipo_idx).value
            cpe_cell = ws.cell(row=cur_row, column=cpe_idx)
            if tipo_val in ('Radio', 'Televisión') and isinstance(cpe_cell.value, (int, float)):
                cpe_cell.number_format = num_fmt_int

        # Hipervínculos
        for ci, url in link_map.items():
            cell = ws.cell(row=cur_row, column=ci)
            cell.hyperlink = url
            cell.font      = font_link
            cell.alignment = align_left

    # Anchos de columna
    for i, col_name in enumerate(cols, start=1):
        letter = ws.cell(row=1, column=i).column_letter
        if col_name in ('Título', 'Resumen - Aclaracion'):
            ws.column_dimensions[letter].width = 50
        elif 'Link' in col_name:
            ws.column_dimensions[letter].width = 15
        else:
            ws.column_dimensions[letter].width = 20

    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue()

# ──────────────────────────────────────────────────────────────────────────────
# PROCESO COMPLETO
# ──────────────────────────────────────────────────────────────────────────────
def run_full_process(dossier_file, download_placeholder):
    st.markdown("<hr>", unsafe_allow_html=True)
    progress_bar = st.progress(0, text="Iniciando proceso...")

    # 1. Modelos y configuración
    progress_bar.progress(5, text="Paso 1 / 7 — Cargando modelos PKL y configuración...")
    sentiment_pipeline, topic_pipeline = load_ml_models()

    config_path = _find_config_file()
    if not config_path:
        st.error(
            "**No se encontró `Configuracion.xlsx`** en el repositorio. "
            "Asegúrate de que esté en la raíz del proyecto."
        )
        st.stop()

    try:
        region_map, internet_map, mention_map, final_topic_map = load_config(str(config_path))
    except Exception as e:
        st.error(f"**Error al cargar `Configuracion.xlsx`:** {e}")
        st.stop()

    # 2. Lectura del dossier
    progress_bar.progress(15, text="Paso 2 / 7 — Leyendo Dossier y normalizando estructura...")
    wb = load_workbook(dossier_file, data_only=True)
    df = read_and_normalize_dossier(wb.active, region_map, internet_map)

    if df['Fecha'].isna().any():
        st.warning("⚠️ Algunas fechas no se pudieron convertir. Revisa el archivo original.")

    # 3. Expansión por ; en Menciones
    progress_bar.progress(28, text="Paso 3 / 7 — Expandiendo filas por menciones (;)...")
    df = expand_menciones(df)

    # 4. Mapeo de menciones
    progress_bar.progress(35, text="Paso 4 / 7 — Aplicando mapeos de menciones y configuración...")
    df = apply_mention_map(df, mention_map)

    # 5. Detección de duplicados
    progress_bar.progress(45, text="Paso 4b / 7 — Detectando duplicados...")
    df = detect_duplicates(df)

    # 6 & 7. Clasificación PKL
    df = classify_with_pkl(df, sentiment_pipeline, topic_pipeline, final_topic_map, progress_bar)

    progress_bar.progress(95, text="Paso 7 / 7 — Generando archivo de salida...")
    excel_data  = to_excel(df)
    filename    = f"Dossier_Procesado_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    total       = len(df)
    dups_count  = int(df['is_duplicate'].sum())
    unique_count = total - dups_count

    progress_bar.progress(100, text="✓ Proceso completado")

    # Botón de descarga en el placeholder
    with download_placeholder:
        st.download_button(
            label="⬇ Descargar archivo procesado (.xlsx)",
            data=excel_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown(f"""
    <div class="success-banner">
        <div class="icon">⚡</div>
        <div class="text">
            <strong>Proceso finalizado correctamente</strong>
            <span>{total:,} filas procesadas · {unique_count:,} únicas · {dups_count:,} duplicadas</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<p class="results-header">Resumen del proceso</p>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="metrics-row">
        <div class="metric-card">
            <div class="metric-value">{total:,}</div>
            <div class="metric-label">Filas totales procesadas</div>
        </div>
        <div class="metric-card accent">
            <div class="metric-value">{unique_count:,}</div>
            <div class="metric-label">Noticias únicas analizadas</div>
        </div>
        <div class="metric-card muted">
            <div class="metric-value">{dups_count:,}</div>
            <div class="metric-label">Filas marcadas como duplicadas</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<p class="results-header">Previsualización de resultados</p>', unsafe_allow_html=True)
    cols_preview = [c for c in FINAL_ORDER if c in df.columns]
    df_display   = df[cols_preview].copy()

    if 'Fecha' in df_display.columns:
        df_display['Fecha'] = (
            pd.to_datetime(df_display['Fecha'], errors='coerce')
            .dt.strftime('%d/%m/%Y')
            .fillna('FECHA INVÁLIDA')
        )

    # Aplanar columnas dict para visualización
    for link_col in ['Link Nota', 'Link (Streaming - Imagen)']:
        if link_col in df_display.columns:
            df_display[link_col] = df_display[link_col].apply(
                lambda v: v.get('url', '') if isinstance(v, dict) else (v or '')
            )

    st.dataframe(df_display, use_container_width=True, hide_index=True)

# ──────────────────────────────────────────────────────────────────────────────
# INTERFAZ PRINCIPAL
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <div class="badge">Media Intelligence · PKL</div>
    <p>Limpieza, normalización y clasificación automática de dossiers · v1.0</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="card">
    <div class="card-title">Cómo usar esta herramienta</div>
    <div class="step">
        <div class="step-num">1</div>
        <div class="step-text">Asegúrate de que <strong>Configuracion.xlsx</strong>,
        <strong>pipeline_sentimiento.pkl</strong> y <strong>pipeline_tema.pkl</strong>
        estén en la raíz del repositorio.</div>
    </div>
    <div class="step">
        <div class="step-num">2</div>
        <div class="step-text">Sube el archivo <strong>Dossier (.xlsx)</strong> en el área de abajo.</div>
    </div>
    <div class="step">
        <div class="step-num">3</div>
        <div class="step-text">Haz clic en <strong>Iniciar proceso</strong>. El sistema normaliza,
        detecta duplicados y clasifica tono y tema automáticamente.</div>
    </div>
</div>
""", unsafe_allow_html=True)

with st.expander("📋  Ver estructura requerida para Configuracion.xlsx"):
    st.markdown("""
    | Hoja | Columna A | Columna B |
    |------|-----------|-----------|
    | `Regiones` | Medio | Región |
    | `Internet` | Medio Original | Medio Mapeado |
    | `Menciones` | Mención Original | Mención Mapeada |
    | `Mapa_Temas` | Temas Generales - Tema | Tema |
    """)

st.markdown("<br>", unsafe_allow_html=True)

# Carga del Dossier
st.markdown('<div class="card-title">Carga del Dossier</div>', unsafe_allow_html=True)
dossier_file = st.file_uploader(
    "Arrastra el archivo Dossier aquí o haz clic para seleccionarlo",
    type=["xlsx"],
    accept_multiple_files=False,
    label_visibility="collapsed",
)

if dossier_file:
    st.markdown(
        f'<div class="file-status ok">✓ Dossier cargado — <strong>{dossier_file.name}</strong></div>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        '<div class="file-status missing">⚠ Sube el archivo Dossier (.xlsx) para continuar</div>',
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# Fila de botones
col_start, col_download = st.columns(2)

with col_start:
    start_clicked = st.button(
        "▶  Iniciar proceso completo",
        disabled=dossier_file is None,
        type="primary",
    )

with col_download:
    download_placeholder = st.empty()
    if not start_clicked:
        with download_placeholder:
            st.button(
                "⬇ Descargar archivo procesado (.xlsx)",
                disabled=True,
                type="primary",
            )

if start_clicked:
    run_full_process(dossier_file, download_placeholder)
