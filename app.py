"""
Sistema de Acompanhamento de Estudos — versão evoluída.
Dashboard analítico avançado, menu moderno, KPIs inteligentes, análise por IA e tooltips completos.
"""

import hashlib
import hmac
import html as html_lib
import os
import re
import sqlite3
from sqlite3 import IntegrityError
from contextlib import contextmanager
from datetime import date, datetime, time, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    import psycopg
    from psycopg.errors import UniqueViolation
except ImportError:
    psycopg = None
    UniqueViolation = IntegrityError


# ─────────────────────────────────────────────
# CONFIGURAÇÃO
# ─────────────────────────────────────────────

def config_valor(chave, padrao=None):
    valor = os.getenv(chave)
    if valor not in (None, ""):
        return valor
    try:
        if chave in st.secrets and st.secrets[chave] not in (None, ""):
            return str(st.secrets[chave])
    except Exception:
        pass
    return padrao


APP_NAME = "Sistema de Acompanhamento de Estudos"
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / config_valor("ESTUDOS_DB", "trilha_tjs_ajaa.db")
APP_ENV = config_valor("APP_ENV", "development").lower()
DATABASE_URL = config_valor("DATABASE_URL")
DB_BACKEND = "postgresql" if DATABASE_URL else "sqlite"
IS_PRODUCTION = APP_ENV == "production"
DEBUG = config_valor("DEBUG", "0") == "1" and not IS_PRODUCTION
PLANILHA_REFERENCIA = Path(
    config_valor(
        "ESTUDOS_REFERENCIA",
        r"C:\Users\Rafael.000\Downloads\Acompanhamento da Trilha - TJs - AJAA.xlsx",
    )
)
ADMIN_EMAIL = config_valor("ESTUDOS_ADMIN_EMAIL", "admin@admin.com")
ADMIN_PASSWORD = config_valor("ESTUDOS_ADMIN_PASSWORD", "123")

STATUS_NAO_INICIADA = "NAO_INICIADA"
STATUS_EM_ANDAMENTO = "EM_ANDAMENTO"
STATUS_CONCLUIDA = "CONCLUIDA"
STATUS_VALIDOS = [STATUS_NAO_INICIADA, STATUS_EM_ANDAMENTO, STATUS_CONCLUIDA]
STATUS_ANALISE = [STATUS_EM_ANDAMENTO, STATUS_CONCLUIDA]
STATUS_LABELS = {
    STATUS_NAO_INICIADA: "Não iniciada",
    STATUS_EM_ANDAMENTO: "Em andamento",
    STATUS_CONCLUIDA: "Concluída",
}
STATUS_CORES = {
    STATUS_NAO_INICIADA: "#94a3b8",
    STATUS_EM_ANDAMENTO: "#f59e0b",
    STATUS_CONCLUIDA: "#22c55e",
}
TIPOS_ESTUDO = [
    "Leitura", "Exercícios", "Revisão", "Simulado",
    "Videoaula", "Resumo", "Flashcards", "Projeto", "Pesquisa", "Outro",
]

# Paleta de cores
COR_PRIMARIA = "#1e40af"
COR_SUCESSO  = "#16a34a"
COR_ALERTA   = "#d97706"
COR_PERIGO   = "#dc2626"

# Itens de menu com ícones (usados no menu lateral customizado)
MENU_ITENS_GESTOR = [
    ("📊", "Dashboard"), ("⚡", "Registro rápido"), ("📋", "Tarefas"),
    ("📖", "Aulas e assuntos"), ("🎓", "Disciplinas"), ("👥", "Alunos"),
    ("📥", "Importações"), ("⚙️", "Configurações"),
]
MENU_ITENS_ALUNO = [
    ("📊", "Dashboard"), ("⚡", "Registro rápido"), ("📋", "Tarefas"),
    ("📖", "Aulas e assuntos"), ("⚙️", "Configurações"),
]

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ═══════════════════════════════════════════════════════════
# CSS GLOBAL — menu moderno, dashboard analítico, tooltips
# ═══════════════════════════════════════════════════════════

CSS = """
<style>
:root {
  --c-bg:      #f1f5f9;
  --c-sidebar: #0f172a;
  --c-accent:  #3b82f6;
  --c-text:    #0f172a;
  --c-muted:   #64748b;
  --c-border:  #e2e8f0;
  --c-white:   #ffffff;
  --c-ok:      #16a34a;
  --c-warn:    #d97706;
  --c-danger:  #dc2626;
  color-scheme: light;
}

/* ══ LAYOUT ══ */
.block-container { padding: 1rem 1.5rem 2rem; max-width: 1560px; }
[data-testid="stAppViewContainer"] > .main { background: var(--c-bg); }

/* ══ SIDEBAR — textos sempre visíveis, menu moderno ══ */
[data-testid="stSidebar"] {
  background: var(--c-sidebar) !important;
  min-width: 240px !important;
  max-width: 260px !important;
}
/* Labels de widgets na sidebar sempre visíveis */
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stDateInput label,
[data-testid="stSidebar"] .stNumberInput label,
[data-testid="stSidebar"] .stTextInput label,
[data-testid="stSidebar"] .stToggle label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span {
  color: #cbd5e1 !important;
  opacity: 1 !important;
  visibility: visible !important;
}
/* Títulos de seção na sidebar */
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] .stMarkdown h3 {
  color: #94a3b8 !important;
  font-size: .68rem !important;
  font-weight: 800 !important;
  letter-spacing: .12em !important;
  text-transform: uppercase !important;
  margin: 1.2rem 0 .4rem !important;
  padding: 0 !important;
}
/* Radio buttons do menu — sempre visíveis */
[data-testid="stSidebar"] .stRadio > div { gap: 2px !important; }
[data-testid="stSidebar"] .stRadio label {
  display: flex !important;
  align-items: center !important;
  gap: 8px !important;
  padding: 9px 14px !important;
  border-radius: 8px !important;
  margin: 1px 0 !important;
  cursor: pointer !important;
  color: #cbd5e1 !important;
  font-size: .85rem !important;
  font-weight: 500 !important;
  opacity: 1 !important;
  visibility: visible !important;
  transition: background .15s, color .15s !important;
}
[data-testid="stSidebar"] .stRadio label:hover {
  background: rgba(255,255,255,.08) !important;
  color: #f1f5f9 !important;
}
[data-testid="stSidebar"] .stRadio label[data-testid="stWidgetLabel"] { display: none !important; }
/* Oculta o círculo padrão do radio e usa highlight de fundo */
[data-testid="stSidebar"] .stRadio input[type="radio"] { display: none !important; }
[data-testid="stSidebar"] .stRadio input[type="radio"]:checked + div + label,
[data-testid="stSidebar"] .stRadio input[type="radio"]:checked ~ label {
  background: rgba(59,130,246,.25) !important;
  color: #93c5fd !important;
  font-weight: 700 !important;
}
/* Botões da sidebar */
[data-testid="stSidebar"] .stButton > button {
  background: rgba(255,255,255,.07) !important;
  color: #e2e8f0 !important;
  border: 1px solid rgba(255,255,255,.12) !important;
  border-radius: 8px !important;
  font-size: .82rem !important;
  font-weight: 600 !important;
  width: 100% !important;
  text-align: left !important;
  padding: 9px 14px !important;
  margin-bottom: 3px !important;
  display: flex !important;
  align-items: center !important;
  gap: 8px !important;
  transition: background .15s !important;
}
[data-testid="stSidebar"] .stButton > button:hover {
  background: rgba(255,255,255,.14) !important;
  color: #f8fafc !important;
}
[data-testid="stSidebar"] .stButton > button p { color: inherit !important; font-size: .82rem !important; }
/* Selectbox/inputs da sidebar */
[data-testid="stSidebar"] .stSelectbox > div > div,
[data-testid="stSidebar"] .stMultiSelect > div > div {
  background: rgba(255,255,255,.07) !important;
  border-color: rgba(255,255,255,.15) !important;
  color: #e2e8f0 !important;
}
[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] .stNumberInput input {
  background: rgba(255,255,255,.07) !important;
  border-color: rgba(255,255,255,.15) !important;
  color: #e2e8f0 !important;
}
/* Cabeçalho do usuário na sidebar */
.sidebar-user {
  padding: 16px 14px 12px;
  border-bottom: 1px solid rgba(255,255,255,.1);
  margin-bottom: 8px;
}
.sidebar-user-name { color: #f1f5f9; font-size: .92rem; font-weight: 700; margin: 0; }
.sidebar-user-role { color: #64748b; font-size: .72rem; margin: 2px 0 0; }
.sidebar-nav-section {
  padding: 6px 14px 2px;
  color: #475569;
  font-size: .66rem;
  font-weight: 800;
  letter-spacing: .1em;
  text-transform: uppercase;
}

/* ══ MÉTRICAS NATIVAS ══ */
[data-testid="stMetric"] {
  background: var(--c-white);
  border: 1px solid var(--c-border);
  border-radius: 10px;
  padding: 14px 16px;
  box-shadow: 0 1px 3px rgba(15,23,42,.05);
}
[data-testid="stMetricValue"] { font-size: 1.4rem; font-weight: 800; color: var(--c-text); }
[data-testid="stMetricLabel"] { color: var(--c-muted); font-size: .78rem; font-weight: 600; }
[data-testid="stMetricDelta"] { font-size: .76rem !important; }

/* ══ DATAFRAMES ══ */
div[data-testid="stDataFrame"] { border: 1px solid var(--c-border); border-radius: 10px; overflow: hidden; }

/* ══ HERO ══ */
.hero {
  background: linear-gradient(135deg,#eff6ff 0%,#f8fafc 50%,#ecfdf5 100%);
  border: 1px solid #dbeafe; border-radius: 12px; padding: 18px 24px; margin-bottom: 16px;
}
.hero h1 { margin: 0 0 4px; font-size: 1.5rem; color: var(--c-text); font-weight: 800; }
.hero p  { margin: 0; color: var(--c-muted); font-size: .88rem; }

/* ══ CARDS DE ATIVIDADE ══ */
.quick-card {
  background: var(--c-white); border: 1px solid var(--c-border);
  border-left: 5px solid #94a3b8; border-radius: 10px;
  padding: 14px 18px; margin-bottom: 14px;
  box-shadow: 0 2px 8px rgba(15,23,42,.05);
}
.quick-card.ok   { border-left-color: var(--c-ok); }
.quick-card.warn { border-left-color: var(--c-warn); }
.quick-card.off  { border-left-color: #94a3b8; }
.quick-card h3   { margin: 4px 0 0; font-size: 1rem; }
.quick-grid { display: grid; grid-template-columns: repeat(4,minmax(0,1fr)); gap: 8px; margin-top: 12px; }
@media(max-width:900px){ .quick-grid{ grid-template-columns: repeat(2,1fr); } }
@media(max-width:560px){ .quick-grid{ grid-template-columns: 1fr; } }
.quick-item { background: #f8fafc; border: 1px solid var(--c-border); border-radius: 7px; padding: 8px 10px; }
.quick-label { color: var(--c-muted); font-size: .66rem; font-weight: 800; text-transform: uppercase; letter-spacing: .06em; }
.quick-value { color: var(--c-text); font-weight: 700; margin-top: 2px; font-size: .87rem; overflow-wrap: anywhere; }

/* ══ STATUS BADGES ══ */
.status-ok, .status-warn, .status-off {
  display: inline-flex; align-items: center; gap: 4px;
  padding: 3px 10px; border-radius: 999px; font-weight: 700; font-size: .77rem; margin-top: 4px;
}
.status-ok   { background: #dcfce7; color: #166534; }
.status-warn { background: #fef3c7; color: #92400e; }
.status-off  { background: #f1f5f9; color: #475569; }

/* ══ KPI CARD CUSTOMIZADO com tooltip ══ */
.kpi-card {
  background: var(--c-white); border: 1px solid var(--c-border); border-radius: 10px;
  padding: 14px 16px; position: relative; box-shadow: 0 1px 4px rgba(15,23,42,.05);
  height: 100%; min-height: 100px;
}
.kpi-label  { font-size: .67rem; font-weight: 800; color: var(--c-muted); text-transform: uppercase; letter-spacing: .07em; margin-bottom: 2px; }
.kpi-value  { font-size: 1.5rem; font-weight: 900; color: var(--c-text); margin: 4px 0 2px; line-height: 1.1; }
.kpi-sub    { font-size: .7rem; color: var(--c-muted); line-height: 1.4; }
.kpi-delta-pos { font-size: .7rem; color: var(--c-ok); font-weight: 700; margin-top: 3px; }
.kpi-delta-neg { font-size: .7rem; color: var(--c-danger); font-weight: 700; margin-top: 3px; }
/* Tooltip */
.kpi-tooltip {
  position: absolute; top: 8px; right: 8px;
  width: 17px; height: 17px; border-radius: 50%;
  background: #e2e8f0; color: #64748b;
  font-size: .64rem; font-weight: 900;
  display: inline-flex; align-items: center; justify-content: center;
  cursor: help; user-select: none;
}
.kpi-tooltip[title] { text-decoration: none; }
/* Tooltip via CSS puro — visível em hover */
.kpi-ttip-wrap { display: inline-block; position: absolute; top: 8px; right: 8px; }
.kpi-ttip-icon {
  width: 17px; height: 17px; border-radius: 50%; background: #e2e8f0;
  color: #64748b; font-size: .63rem; font-weight: 900;
  display: inline-flex; align-items: center; justify-content: center; cursor: help;
}
.kpi-ttip-box {
  display: none; position: absolute; z-index: 9999;
  bottom: 130%; right: 0;
  background: #1e293b; color: #f1f5f9;
  font-size: .72rem; line-height: 1.55; padding: 10px 13px;
  border-radius: 9px; width: 270px;
  box-shadow: 0 6px 20px rgba(0,0,0,.35); white-space: normal;
}
.kpi-ttip-box::after {
  content: ""; position: absolute; top: 100%; right: 4px;
  border: 6px solid transparent; border-top-color: #1e293b;
}
.kpi-ttip-wrap:hover .kpi-ttip-box { display: block; }

/* ══ INSIGHT CARDS ══ */
.insight-card {
  border-radius: 10px; padding: 12px 16px; margin-bottom: 8px;
  display: flex; gap: 12px; align-items: flex-start;
}
.insight-card.info    { background: #eff6ff; border-left: 4px solid #3b82f6; }
.insight-card.success { background: #f0fdf4; border-left: 4px solid #22c55e; }
.insight-card.warning { background: #fffbeb; border-left: 4px solid #f59e0b; }
.insight-card.danger  { background: #fef2f2; border-left: 4px solid #ef4444; }
.insight-icon  { font-size: 1.1rem; min-width: 20px; padding-top: 1px; }
.insight-body  { flex: 1; }
.insight-title { font-size: .83rem; font-weight: 800; color: var(--c-text); margin: 0 0 2px; }
.insight-text  { font-size: .77rem; color: var(--c-muted); margin: 0; line-height: 1.45; }

/* ══ REGRAS DE NEGÓCIO ══ */
.rule-warning { background:#fffbeb; border:1px solid #fde68a; border-left:5px solid #f59e0b; border-radius:8px; padding:12px 14px; margin:10px 0; color:#78350f; font-size:.88rem; }
.rule-error   { background:#fef2f2; border:1px solid #fecaca; border-left:5px solid #ef4444; border-radius:8px; padding:12px 14px; margin:10px 0; color:#7f1d1d; font-size:.88rem; }

/* ══ UTILITÁRIOS ══ */
.section-title { font-size:.93rem; font-weight:800; color:var(--c-text); margin:14px 0 6px; }
.muted { color:var(--c-muted); font-size:.84rem; }

/* ══ BOTÕES ══ */
.stButton > button[kind="primary"] { background: var(--c-accent); color:#fff; border-radius:8px; font-weight:700; }
.stButton > button[kind="primary"]:hover { background:#1d4ed8; }

/* ══ FORMULÁRIOS ══ */
[data-testid="stForm"] { border:1px solid var(--c-border) !important; border-radius:10px !important; padding:16px !important; }

/* ══ GRÁFICO RANK ══ */
.rank-row {
  display: flex; align-items: center; gap: 10px;
  padding: 9px 12px; border-radius: 8px; background: #f8fafc;
  border: 1px solid var(--c-border); margin-bottom: 6px;
}
.rank-pos { font-size: 1rem; font-weight: 900; color: var(--c-muted); min-width: 28px; text-align: center; }
.rank-name { flex: 1; font-size: .85rem; font-weight: 700; color: var(--c-text); }
.rank-bar-wrap { width: 100px; height: 8px; background: #e2e8f0; border-radius: 999px; overflow: hidden; }
.rank-bar { height: 100%; border-radius: 999px; background: var(--c-accent); }
.rank-val { font-size: .8rem; font-weight: 700; color: var(--c-text); min-width: 48px; text-align: right; }

/* ══ RESPONSIVIDADE ══ */
@media(max-width:768px){
  .block-container { padding: .75rem 1rem 1.5rem; }
  .kpi-value { font-size: 1.2rem !important; }
}
</style>
"""

st.markdown(CSS, unsafe_allow_html=True)



# ─────────────────────────────────────────────
# UTILITÁRIOS DE RENDERIZAÇÃO
# ─────────────────────────────────────────────

def render_html(conteudo: str):
    if hasattr(st, "html"):
        st.html(conteudo)
    else:
        st.markdown(conteudo, unsafe_allow_html=True)


def escape_html(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    return html_lib.escape(str(valor))


def status_badge(status: str) -> str:
    cls = {STATUS_CONCLUIDA:"status-ok", STATUS_EM_ANDAMENTO:"status-warn", STATUS_NAO_INICIADA:"status-off"}.get(status,"status-off")
    icone = {"CONCLUIDA":"✓","EM_ANDAMENTO":"●","NAO_INICIADA":"○"}.get(status,"○")
    return f'<span class="{cls}">{icone} {escape_html(STATUS_LABELS.get(status,status))}</span>'


def status_card_class(status: str) -> str:
    return {STATUS_CONCLUIDA:"ok",STATUS_EM_ANDAMENTO:"warn",STATUS_NAO_INICIADA:"off"}.get(status,"off")


def kpi_card(label: str, valor: str, subtexto: str = "", delta: str = "", delta_pos: bool = True, tooltip: str = "") -> str:
    """KPI card com tooltip sempre visível ao hover (CSS puro, sem JS)."""
    delta_html = ""
    if delta:
        cls = "kpi-delta-pos" if delta_pos else "kpi-delta-neg"
        seta = "▲" if delta_pos else "▼"
        delta_html = f'<div class="{cls}">{seta} {escape_html(delta)}</div>'
    tooltip_html = ""
    if tooltip:
        tooltip_html = (
            '<div class="kpi-ttip-wrap">'
            '<div class="kpi-ttip-icon">?</div>'
            f'<div class="kpi-ttip-box">{escape_html(tooltip)}</div>'
            '</div>'
        )
    return (
        '<div class="kpi-card">'
        + tooltip_html
        + f'<div class="kpi-label">{escape_html(label)}</div>'
        + f'<div class="kpi-value">{escape_html(str(valor))}</div>'
        + f'<div class="kpi-sub">{escape_html(subtexto)}</div>'
        + delta_html
        + '</div>'
    )


def insight_card(tipo: str, icone: str, titulo: str, texto: str) -> str:
    return (
        f'<div class="insight-card {tipo}">'
        f'<div class="insight-icon">{icone}</div>'
        '<div class="insight-body">'
        f'<div class="insight-title">{escape_html(titulo)}</div>'
        f'<p class="insight-text">{escape_html(texto)}</p>'
        '</div></div>'
    )


# ─────────────────────────────────────────────
# UTILITÁRIOS GERAIS
# ─────────────────────────────────────────────

def limpar_texto(valor):
    if pd.isna(valor):
        return None
    texto = str(valor).replace("\n", " ").strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto if texto else None


def quebrar_texto(texto, limite=26):
    texto = limpar_texto(texto) or ""
    partes, linha = [], []
    for palavra in texto.split():
        candidato = " ".join(linha + [palavra])
        if linha and len(candidato) > limite:
            partes.append(" ".join(linha))
            linha = [palavra]
        else:
            linha.append(palavra)
    if linha:
        partes.append(" ".join(linha))
    return "<br>".join(partes) if partes else texto


def normalizar_email(valor):
    email = limpar_texto(valor)
    return email.lower() if email else None


def email_local(nome):
    nome = str(nome or "aluno").lower().strip()
    nome = re.sub(r"[^a-z0-9\s.]", "", nome)
    nome = re.sub(r"\s+", ".", nome)
    nome = re.sub(r"\.+", ".", nome).strip(".")
    return f"{nome or 'aluno'}@local.com"


def chave_texto(texto):
    texto = limpar_texto(texto) or ""
    mapa = str.maketrans("áàâãéêíóôõúüçÁÀÂÃÉÊÍÓÔÕÚÜÇ", "aaaaeeiooouucAAAAEEIOOOUUC")
    texto = texto.translate(mapa).lower()
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def converter_horas(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, time):
        return valor.hour + valor.minute / 60 + valor.second / 3600
    if isinstance(valor, datetime):
        return valor.hour + valor.minute / 60 + valor.second / 3600
    if isinstance(valor, timedelta):
        return valor.total_seconds() / 3600
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        valor = valor.strip()
        if not valor:
            return 0.0
        if ":" in valor:
            partes = valor.split(":")
            try:
                h = int(partes[0])
                m = int(partes[1]) if len(partes) > 1 else 0
                s = int(partes[2]) if len(partes) > 2 else 0
                return h + m / 60 + s / 3600
            except ValueError:
                return 0.0
        try:
            return float(valor.replace(",", "."))
        except ValueError:
            return 0.0
    return 0.0


def converter_numero(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, (time, datetime, timedelta)):
        return converter_horas(valor)
    if isinstance(valor, str):
        valor = valor.strip().replace("%", "").replace(",", ".")
        if not valor:
            return 0.0
        try:
            return float(valor)
        except ValueError:
            return 0.0
    return 0.0


def converter_inteiro(valor):
    return int(round(converter_numero(valor)))


def converter_data(valor):
    if pd.isna(valor) or valor in ("", None):
        return None
    try:
        return str(pd.to_datetime(valor).date())
    except Exception:
        return None


def formatar_data_br(valor):
    data = converter_data(valor)
    if not data:
        return ""
    return pd.to_datetime(data).strftime("%d/%m/%Y")


def nome_aluno_data_arquivo(caminho):
    caminho = Path(caminho)
    nome = caminho.stem.split("-", 1)[0].strip()
    data_execucao = date.today()
    match = re.search(r"(\d{2})_(\d{2})_(\d{4})", caminho.stem)
    if match:
        dia, mes, ano = map(int, match.groups())
        data_execucao = date(ano, mes, dia)
    return nome, data_execucao


def distribuir_inteiro(total, pesos):
    total = int(round(total or 0))
    if not pesos:
        return []
    if total == 0:
        return [0 for _ in pesos]
    pesos = [float(p or 0) for p in pesos]
    if sum(pesos) <= 0:
        pesos = [1 for _ in pesos]
    bruto = [total * p / sum(pesos) for p in pesos]
    base = [int(v) for v in bruto]
    sobra = total - sum(base)
    ordem = sorted(range(len(bruto)), key=lambda i: bruto[i] - base[i], reverse=True)
    for i in ordem[:sobra]:
        base[i] += 1
    return base


def normalizar_tipo_estudo(valor):
    tipo = limpar_texto(valor) or "Outro"
    aliases = {
        "teoria": "Leitura",
        "questoes": "Exercícios",
        "questões": "Exercícios",
        "exercicios": "Exercícios",
        "exercícios": "Exercícios",
        "video aula": "Videoaula",
        "vídeo aula": "Videoaula",
    }
    tipo = aliases.get(chave_texto(tipo), tipo)
    return tipo if tipo in TIPOS_ESTUDO else "Outro"


def erro_usuario(mensagem, exc=None):
    if DEBUG and exc is not None:
        st.error(f"{mensagem}: {exc}")
    else:
        st.error(mensagem)


# ─────────────────────────────────────────────
# SEGURANÇA — SENHAS
# ─────────────────────────────────────────────

def hash_senha(senha):
    salt = os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", str(senha).encode(), salt, 120_000)
    return f"pbkdf2_sha256${salt.hex()}${digest.hex()}"


def verificar_senha(senha, senha_armazenada):
    if not senha_armazenada:
        return False
    if not str(senha_armazenada).startswith("pbkdf2_sha256$"):
        return hmac.compare_digest(str(senha), str(senha_armazenada))
    try:
        _, salt_hex, digest_hex = str(senha_armazenada).split("$", 2)
        digest = hashlib.pbkdf2_hmac("sha256", str(senha).encode(), bytes.fromhex(salt_hex), 120_000)
        return hmac.compare_digest(digest.hex(), digest_hex)
    except ValueError:
        return False


def senha_valida(nova):
    return (
        len(str(nova or "")) >= 8
        and bool(re.search(r"[A-Za-z]", str(nova)))
        and bool(re.search(r"\d", str(nova)))
    )


def atualizar_senha_usuario(usuario_id, nova_senha, forcar_troca=0):
    executar(
        "UPDATE alunos SET senha = ?, force_troca_senha = ? WHERE id = ?",
        (hash_senha(nova_senha), int(forcar_troca), int(usuario_id)),
    )


# ─────────────────────────────────────────────
# BANCO DE DADOS
# ─────────────────────────────────────────────

def adaptar_sql(sql):
    if DB_BACKEND != "postgresql":
        return sql
    return sql.replace("?", "%s")


class ConexaoDB:
    """Wrapper para compatibilidade entre SQLite e PostgreSQL."""
    def __init__(self, conn):
        self.conn = conn
        self._cursor = None

    def execute(self, sql, params=()):
        cur = self.conn.cursor()
        cur.execute(adaptar_sql(sql), params)
        self._cursor = cur
        return cur

    def cursor(self):
        return self

    def fetchone(self):
        return self._cursor.fetchone()

    def fetchall(self):
        return self._cursor.fetchall()

    def commit(self):
        self.conn.commit()

    def rollback(self):
        self.conn.rollback()

    def close(self):
        self.conn.close()


def abrir_conexao_raw():
    if DB_BACKEND == "postgresql":
        if psycopg is None:
            raise RuntimeError("Instale psycopg[binary] para usar PostgreSQL em produção.")
        return psycopg.connect(DATABASE_URL)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@contextmanager
def conectar():
    raw = abrir_conexao_raw()
    conn = ConexaoDB(raw) if DB_BACKEND == "postgresql" else raw
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def executar(sql, params=()):
    with conectar() as conn:
        conn.execute(sql, params)


def consultar(sql, params=()):
    conn = abrir_conexao_raw()
    try:
        return pd.read_sql_query(adaptar_sql(sql), conn, params=params)
    finally:
        conn.close()


def ultimo_id(conn):
    if DB_BACKEND == "postgresql":
        return int(conn.execute("SELECT LASTVAL()").fetchone()[0])
    return int(conn.execute("SELECT last_insert_rowid()").fetchone()[0])


def erro_integridade(exc):
    return isinstance(exc, (IntegrityError, UniqueViolation))


# ─────────────────────────────────────────────
# CRIAÇÃO DAS TABELAS
# ─────────────────────────────────────────────

def criar_tabelas():
    if DB_BACKEND == "postgresql":
        _criar_tabelas_postgresql()
        return
    _criar_tabelas_sqlite()


def _criar_tabelas_postgresql():
    with conectar() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS alunos (
                id SERIAL PRIMARY KEY,
                nome TEXT NOT NULL UNIQUE,
                email TEXT UNIQUE,
                senha TEXT NOT NULL,
                perfil TEXT NOT NULL DEFAULT 'Aluno' CHECK(perfil IN ('Gestor','Aluno')),
                ativo INTEGER DEFAULT 1,
                force_troca_senha INTEGER DEFAULT 1
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS disciplinas (
                id SERIAL PRIMARY KEY,
                nome TEXT NOT NULL UNIQUE,
                ativo INTEGER DEFAULT 1
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS aulas (
                id SERIAL PRIMARY KEY,
                disciplina_id INTEGER NOT NULL REFERENCES disciplinas(id),
                aula TEXT NOT NULL,
                assunto TEXT,
                estudada_padrao TEXT DEFAULT 'Não',
                revisao_24h_padrao TEXT DEFAULT 'Não',
                tipo_estudo TEXT DEFAULT 'Outro',
                ativo INTEGER DEFAULT 1
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS assuntos (
                id SERIAL PRIMARY KEY,
                aula_id INTEGER NOT NULL REFERENCES aulas(id),
                titulo TEXT NOT NULL,
                ativo INTEGER DEFAULT 1,
                UNIQUE(aula_id, titulo)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS tarefas (
                id SERIAL PRIMARY KEY,
                numero INTEGER NOT NULL UNIQUE,
                trilha INTEGER,
                disciplina_id INTEGER NOT NULL REFERENCES disciplinas(id),
                seq_disciplina INTEGER,
                aula TEXT,
                qtd_exercicios_previstos INTEGER DEFAULT 0,
                tipo TEXT,
                conteudo TEXT,
                ativo INTEGER DEFAULT 1,
                aula_id INTEGER REFERENCES aulas(id),
                assunto_id INTEGER REFERENCES assuntos(id)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS execucoes (
                id SERIAL PRIMARY KEY,
                aluno_id INTEGER NOT NULL REFERENCES alunos(id),
                tarefa_id INTEGER NOT NULL REFERENCES tarefas(id),
                data_execucao TEXT,
                ch_efetiva REAL DEFAULT 0,
                data_revisao_24h TEXT,
                ch_revisao REAL DEFAULT 0,
                qtd_acertos INTEGER DEFAULT 0,
                desempenho REAL DEFAULT 0,
                comentario TEXT,
                concluida INTEGER DEFAULT 0,
                atualizado_em TEXT DEFAULT CURRENT_TIMESTAMP,
                qtd_questoes_feitas INTEGER DEFAULT 0,
                status TEXT DEFAULT 'NAO_INICIADA',
                tipo_estudo TEXT DEFAULT 'Outro',
                UNIQUE(aluno_id, tarefa_id)
            )
        """)
        # Migrações seguras
        conn.execute("ALTER TABLE alunos ADD COLUMN IF NOT EXISTS force_troca_senha INTEGER DEFAULT 1")
        conn.execute("ALTER TABLE execucoes ADD COLUMN IF NOT EXISTS qtd_questoes_feitas INTEGER DEFAULT 0")
        conn.execute("ALTER TABLE execucoes ADD COLUMN IF NOT EXISTS status TEXT DEFAULT 'NAO_INICIADA'")
        conn.execute("ALTER TABLE execucoes ADD COLUMN IF NOT EXISTS tipo_estudo TEXT DEFAULT 'Outro'")
        conn.execute("ALTER TABLE aulas ADD COLUMN IF NOT EXISTS tipo_estudo TEXT DEFAULT 'Outro'")
        conn.execute("ALTER TABLE tarefas ADD COLUMN IF NOT EXISTS aula_id INTEGER")
        conn.execute("ALTER TABLE tarefas ADD COLUMN IF NOT EXISTS assunto_id INTEGER")
        _normalizar_status(conn)
        conn.execute("DELETE FROM alunos WHERE email = 'aluno@local.com' OR nome = 'Aluno Padrão'")
        _criar_indices(conn)


def _criar_tabelas_sqlite():
    with conectar() as conn:
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS alunos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL UNIQUE,
                email TEXT UNIQUE,
                senha TEXT NOT NULL,
                perfil TEXT NOT NULL DEFAULT 'Aluno' CHECK(perfil IN ('Gestor','Aluno')),
                force_troca_senha INTEGER DEFAULT 1,
                ativo INTEGER DEFAULT 1
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS disciplinas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL UNIQUE,
                ativo INTEGER DEFAULT 1
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS aulas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                disciplina_id INTEGER NOT NULL,
                aula TEXT NOT NULL,
                assunto TEXT,
                estudada_padrao TEXT DEFAULT 'Não',
                revisao_24h_padrao TEXT DEFAULT 'Não',
                tipo_estudo TEXT DEFAULT 'Outro',
                ativo INTEGER DEFAULT 1,
                FOREIGN KEY (disciplina_id) REFERENCES disciplinas(id)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS assuntos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                aula_id INTEGER NOT NULL,
                titulo TEXT NOT NULL,
                ativo INTEGER DEFAULT 1,
                UNIQUE(aula_id, titulo),
                FOREIGN KEY (aula_id) REFERENCES aulas(id)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS tarefas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                numero INTEGER NOT NULL UNIQUE,
                trilha INTEGER,
                disciplina_id INTEGER NOT NULL,
                seq_disciplina INTEGER,
                aula TEXT,
                qtd_exercicios_previstos INTEGER DEFAULT 0,
                tipo TEXT,
                conteudo TEXT,
                ativo INTEGER DEFAULT 1,
                aula_id INTEGER,
                assunto_id INTEGER,
                FOREIGN KEY (disciplina_id) REFERENCES disciplinas(id),
                FOREIGN KEY (aula_id) REFERENCES aulas(id),
                FOREIGN KEY (assunto_id) REFERENCES assuntos(id)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS execucoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                aluno_id INTEGER NOT NULL,
                tarefa_id INTEGER NOT NULL,
                data_execucao TEXT,
                ch_efetiva REAL DEFAULT 0,
                data_revisao_24h TEXT,
                ch_revisao REAL DEFAULT 0,
                qtd_acertos INTEGER DEFAULT 0,
                desempenho REAL DEFAULT 0,
                comentario TEXT,
                concluida INTEGER DEFAULT 0,
                atualizado_em TEXT DEFAULT CURRENT_TIMESTAMP,
                qtd_questoes_feitas INTEGER DEFAULT 0,
                status TEXT DEFAULT 'NAO_INICIADA',
                tipo_estudo TEXT DEFAULT 'Outro',
                UNIQUE(aluno_id, tarefa_id),
                FOREIGN KEY (aluno_id) REFERENCES alunos(id),
                FOREIGN KEY (tarefa_id) REFERENCES tarefas(id)
            )
        """)
        # Migrações: adiciona colunas ausentes com segurança
        _migrar_coluna_sqlite(cur, "execucoes", "qtd_questoes_feitas", "INTEGER DEFAULT 0")
        _migrar_coluna_sqlite(cur, "execucoes", "status", "TEXT DEFAULT 'NAO_INICIADA'")
        _migrar_coluna_sqlite(cur, "execucoes", "tipo_estudo", "TEXT DEFAULT 'Outro'")
        _migrar_coluna_sqlite(cur, "aulas", "tipo_estudo", "TEXT DEFAULT 'Outro'")
        _migrar_coluna_sqlite(cur, "tarefas", "aula_id", "INTEGER")
        _migrar_coluna_sqlite(cur, "tarefas", "assunto_id", "INTEGER")
        _migrar_coluna_sqlite(cur, "alunos", "force_troca_senha", "INTEGER DEFAULT 0")
        _normalizar_status(conn)
        cur.execute("DELETE FROM alunos WHERE email = 'aluno@local.com' OR nome = 'Aluno Padrão'")
        _criar_indices(conn)


def _migrar_coluna_sqlite(cur, tabela, coluna, definicao):
    colunas = {row[1] for row in cur.execute(f"PRAGMA table_info({tabela})").fetchall()}
    if coluna not in colunas:
        cur.execute(f"ALTER TABLE {tabela} ADD COLUMN {coluna} {definicao}")


def _normalizar_status(conn):
    conn.execute("""
        UPDATE execucoes
        SET status = CASE
            WHEN status IN ('NAO_INICIADA','EM_ANDAMENTO','CONCLUIDA') THEN status
            WHEN concluida = 1 THEN 'CONCLUIDA'
            ELSE 'NAO_INICIADA'
        END
    """)
    conn.execute("UPDATE execucoes SET concluida = CASE WHEN status = 'CONCLUIDA' THEN 1 ELSE 0 END")


def _criar_indices(conn):
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tarefas_disciplina ON tarefas(disciplina_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tarefas_trilha ON tarefas(trilha)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_exec_aluno ON execucoes(aluno_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_exec_tarefa ON execucoes(tarefa_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_exec_status ON execucoes(status)")


def inserir_admin():
    existe = consultar("SELECT id FROM alunos WHERE email = ?", (ADMIN_EMAIL,))
    if existe.empty:
        executar(
            "INSERT INTO alunos (nome, email, senha, perfil, ativo, force_troca_senha) VALUES (?, ?, ?, 'Gestor', 1, 1)",
            ("Administrador", ADMIN_EMAIL, hash_senha(ADMIN_PASSWORD)),
        )


# ─────────────────────────────────────────────
# CACHE
# ─────────────────────────────────────────────

def limpar_cache():
    carregar_visao_tarefas.clear()
    carregar_execucoes.clear()
    carregar_aulas.clear()
    carregar_assuntos.clear()
    carregar_tarefas_base.clear()


# ─────────────────────────────────────────────
# UPSERTS
# ─────────────────────────────────────────────

def upsert_disciplina(conn, nome):
    nome = limpar_texto(nome)
    if not nome:
        return None
    conn.execute(
        "INSERT INTO disciplinas (nome, ativo) VALUES (?, 1) ON CONFLICT(nome) DO UPDATE SET ativo = 1",
        (nome,),
    )
    return conn.execute("SELECT id FROM disciplinas WHERE nome = ?", (nome,)).fetchone()[0]


def upsert_aula(conn, disciplina_id, aula, estudada="Não", revisao="Não", tipo_estudo="Outro"):
    aula = str(limpar_texto(aula) or "Aula não informada")
    existente = conn.execute(
        "SELECT id FROM aulas WHERE disciplina_id = ? AND aula = ? AND ativo = 1 ORDER BY id LIMIT 1",
        (int(disciplina_id), aula),
    ).fetchone()
    if existente:
        conn.execute(
            "UPDATE aulas SET estudada_padrao = ?, revisao_24h_padrao = ?, tipo_estudo = ?, ativo = 1 WHERE id = ?",
            (estudada or "Não", revisao or "Não", normalizar_tipo_estudo(tipo_estudo), int(existente[0])),
        )
        return int(existente[0])
    conn.execute(
        "INSERT INTO aulas (disciplina_id, aula, estudada_padrao, revisao_24h_padrao, tipo_estudo, ativo) VALUES (?, ?, ?, ?, ?, 1)",
        (int(disciplina_id), aula, estudada or "Não", revisao or "Não", normalizar_tipo_estudo(tipo_estudo)),
    )
    return ultimo_id(conn)


def upsert_assunto(conn, aula_id, titulo):
    titulo = limpar_texto(titulo) or "Conteúdo não informado"
    conn.execute(
        "INSERT INTO assuntos (aula_id, titulo, ativo) VALUES (?, ?, 1) ON CONFLICT(aula_id, titulo) DO UPDATE SET ativo = 1",
        (int(aula_id), titulo),
    )
    return conn.execute("SELECT id FROM assuntos WHERE aula_id = ? AND titulo = ?", (int(aula_id), titulo)).fetchone()[0]


def upsert_execucao(
    conn,
    aluno_id,
    tarefa_id,
    data_execucao=None,
    ch_efetiva=0,
    data_revisao=None,
    ch_revisao=0,
    acertos=0,
    comentario=None,
    questoes_feitas=0,
    status=STATUS_NAO_INICIADA,
    tipo_estudo=None,
):
    if status not in STATUS_VALIDOS:
        status = STATUS_NAO_INICIADA
    tipo_estudo = normalizar_tipo_estudo(tipo_estudo)
    if tipo_estudo == "Outro":
        tarefa_tipo = conn.execute(
            "SELECT COALESCE(tipo, 'Outro') FROM tarefas WHERE id = ?", (int(tarefa_id),)
        ).fetchone()
        tipo_estudo = normalizar_tipo_estudo(tarefa_tipo[0] if tarefa_tipo else None)
    questoes = int(round(questoes_feitas or 0))
    acertos = int(acertos or 0)
    desempenho = round((acertos / questoes) * 100, 2) if questoes > 0 else 0
    concluida = 1 if status == STATUS_CONCLUIDA else 0
    conn.execute(
        """
        INSERT INTO execucoes
            (aluno_id, tarefa_id, data_execucao, ch_efetiva, data_revisao_24h, ch_revisao,
             qtd_questoes_feitas, qtd_acertos, desempenho, status, comentario, concluida, tipo_estudo)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(aluno_id, tarefa_id) DO UPDATE SET
            data_execucao      = excluded.data_execucao,
            ch_efetiva         = excluded.ch_efetiva,
            data_revisao_24h   = excluded.data_revisao_24h,
            ch_revisao         = excluded.ch_revisao,
            qtd_questoes_feitas= excluded.qtd_questoes_feitas,
            qtd_acertos        = excluded.qtd_acertos,
            desempenho         = excluded.desempenho,
            status             = excluded.status,
            comentario         = excluded.comentario,
            concluida          = excluded.concluida,
            tipo_estudo        = excluded.tipo_estudo,
            atualizado_em      = CURRENT_TIMESTAMP
        """,
        (
            int(aluno_id), int(tarefa_id), data_execucao,
            float(ch_efetiva or 0), data_revisao, float(ch_revisao or 0),
            questoes, acertos, desempenho, status, comentario, concluida, tipo_estudo,
        ),
    )


# ─────────────────────────────────────────────
# IMPORTADORES
# ─────────────────────────────────────────────

def importar_planilha_referencia(caminho=PLANILHA_REFERENCIA, substituir=False):
    """
    Importa a planilha de referência (disciplinas, aulas, assuntos, tarefas).

    NUNCA apaga execuções existentes.
    Apenas insere ou atualiza estrutura (disciplinas, aulas, assuntos, tarefas).
    Execuções já existentes (status, datas, horas, acertos) são preservadas.
    O parâmetro `substituir` foi mantido por compatibilidade mas não apaga execuções.
    """
    caminho = Path(caminho)
    if not caminho.exists():
        st.error(f"Planilha não encontrada: {caminho}")
        return False
    try:
        xl = pd.ExcelFile(caminho)
    except Exception as exc:
        erro_usuario("Erro ao abrir a planilha.", exc)
        return False
    if "Tarefas" not in xl.sheet_names:
        st.error("A planilha precisa ter a aba Tarefas.")
        return False

    abas_ignoradas = {"Tutorial", "Estatísticas", "Tarefas"}
    try:
        with conectar() as conn:
            # ── Estrutura: disciplinas, aulas, assuntos ──
            # Nunca apaga. Apenas insere novos ou reactiva inativados.
            for aba in xl.sheet_names:
                if aba in abas_ignoradas:
                    continue
                try:
                    raw = pd.read_excel(caminho, sheet_name=aba, header=None, nrows=1)
                    nome_disciplina = limpar_texto(raw.iloc[0, 0]) or aba
                except Exception:
                    nome_disciplina = aba
                disciplina_id = upsert_disciplina(conn, nome_disciplina)

                try:
                    aulas_df = pd.read_excel(caminho, sheet_name=aba, header=1).dropna(how="all")
                except Exception:
                    continue
                if "Aula" not in aulas_df.columns or "Assunto" not in aulas_df.columns:
                    continue
                for _, row in aulas_df.iterrows():
                    aula    = limpar_texto(row.get("Aula"))
                    assunto = limpar_texto(row.get("Assunto"))
                    if not aula or not assunto:
                        continue
                    aula_id = upsert_aula(
                        conn, disciplina_id, aula,
                        limpar_texto(row.get("Estudada")) or "Não",
                        limpar_texto(row.get("Revisão 24h")) or "Não",
                    )
                    upsert_assunto(conn, aula_id, assunto)

            # ── Tarefas: upsert seguro — não toca em execuções ──
            tarefas = pd.read_excel(caminho, sheet_name="Tarefas", header=2).dropna(how="all")
            for _, row in tarefas.iterrows():
                numero     = converter_inteiro(row.get("Tarefa"))
                disciplina = limpar_texto(row.get("Disciplina"))
                if not numero or not disciplina:
                    continue
                disciplina_id = upsert_disciplina(conn, disciplina)
                aula_valor    = str(limpar_texto(row.get("Aula")) or "Aula não informada")
                aula_id       = upsert_aula(conn, disciplina_id, aula_valor)
                assunto_id    = upsert_assunto(conn, aula_id, limpar_texto(row.get("Conteúdo")))
                tipo          = normalizar_tipo_estudo(row.get("Tipo"))
                conn.execute(
                    """
                    INSERT INTO tarefas
                        (numero, trilha, disciplina_id, seq_disciplina, aula,
                         qtd_exercicios_previstos, tipo, conteudo, ativo, aula_id, assunto_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
                    ON CONFLICT(numero) DO UPDATE SET
                        trilha                  = excluded.trilha,
                        disciplina_id           = excluded.disciplina_id,
                        seq_disciplina          = excluded.seq_disciplina,
                        aula                    = excluded.aula,
                        qtd_exercicios_previstos= excluded.qtd_exercicios_previstos,
                        tipo                    = excluded.tipo,
                        conteudo                = excluded.conteudo,
                        ativo                   = 1,
                        aula_id                 = excluded.aula_id,
                        assunto_id              = excluded.assunto_id
                    """,
                    (
                        numero,
                        converter_inteiro(row.get("Trilha")),
                        disciplina_id,
                        converter_inteiro(row.get("Seq Disciplina")),
                        aula_valor,
                        converter_inteiro(row.get("Qtd. Exercícios")),
                        tipo,
                        limpar_texto(row.get("Conteúdo")),
                        aula_id,
                        assunto_id,
                    ),
                )
    except Exception as exc:
        erro_usuario("Importação cancelada.", exc)
        return False
    limpar_cache()
    return True


def importar_execucoes_ciclo(caminho_excel):
    """
    Importador legado: lê aba CICLO_REG com um aluno por arquivo.
    Mantido por compatibilidade com arquivos individuais.
    """
    caminho_excel = Path(caminho_excel)
    if not caminho_excel.exists():
        st.error(f"Arquivo não encontrado: {caminho_excel}")
        return False
    aluno_nome, data_execucao = nome_aluno_data_arquivo(caminho_excel)
    try:
        df = pd.read_excel(caminho_excel, sheet_name="CICLO_REG", header=2)
    except Exception as exc:
        erro_usuario(f"Erro ao ler {caminho_excel.name}.", exc)
        return False

    df.columns = [str(c).strip().upper() for c in df.columns]
    if "BLOCO" not in df.columns or "DISCIPLINA" not in df.columns:
        st.error(f"A aba CICLO_REG de {caminho_excel.name} não possui BLOCO/DISCIPLINA.")
        return False
    df["BLOCO"] = df["BLOCO"].ffill()
    df = df[df["DISCIPLINA"].notna()].copy()
    aliases = {
        "matematica e raciocinio logico": "matematica e raciocinio logico",
        "administracao geral e publica": "administracao geral e publica e gestao de pessoas",
        "administracao financeira e orcamentaria": "administracao financeira e orcamentaria",
    }

    try:
        with conectar() as conn:
            conn.execute(
                """
                INSERT INTO alunos (nome, email, senha, perfil, ativo, force_troca_senha)
                VALUES (?, ?, ?, 'Aluno', 1, 1)
                ON CONFLICT(nome) DO UPDATE SET email = excluded.email, ativo = 1
                """,
                (aluno_nome, email_local(aluno_nome), hash_senha("123")),
            )
            aluno_id = conn.execute("SELECT id FROM alunos WHERE nome = ?", (aluno_nome,)).fetchone()[0]
            conn.execute("DELETE FROM execucoes WHERE aluno_id = ?", (aluno_id,))
            tarefas = conn.execute(
                """
                SELECT t.id, COALESCE(t.trilha, 0), d.nome, COALESCE(t.qtd_exercicios_previstos, 0)
                FROM tarefas t
                JOIN disciplinas d ON d.id = t.disciplina_id
                WHERE t.ativo = 1 AND d.ativo = 1
                ORDER BY t.numero
                """
            ).fetchall()
            por_chave = {}
            for tarefa_id, trilha, disciplina, previstos in tarefas:
                chave = aliases.get(chave_texto(disciplina), chave_texto(disciplina))
                por_chave.setdefault((int(trilha or 0), chave), []).append((int(tarefa_id), int(previstos or 0)))

            for _, row in df.iterrows():
                bloco = limpar_texto(row.get("BLOCO"))
                disciplina = limpar_texto(row.get("DISCIPLINA"))
                if not bloco or not disciplina:
                    continue
                match = re.search(r"(\d+)", bloco)
                if not match:
                    continue
                trilha = int(match.group(1))
                chave_disc = aliases.get(chave_texto(disciplina), chave_texto(disciplina))
                tarefas_match = por_chave.get((trilha, chave_disc), [])
                if not tarefas_match:
                    continue
                ch_total = converter_horas(row.get("CH (EFETIVA)"))
                questoes_total = converter_inteiro(row.get("TOT QUEST FEITAS"))
                acertos_total = converter_inteiro(row.get("TOT ACERTOS"))
                comentario = limpar_texto(row.get("AULA ATUAL"))
                if not (ch_total > 0 or questoes_total > 0 or acertos_total > 0 or comentario):
                    continue
                pesos = [previstos for _, previstos in tarefas_match]
                questoes_dist = distribuir_inteiro(questoes_total, pesos)
                acertos_dist = distribuir_inteiro(acertos_total, questoes_dist)
                ch_por_tarefa = ch_total / len(tarefas_match) if tarefas_match else 0
                for idx, (tarefa_id, _) in enumerate(tarefas_match):
                    upsert_execucao(
                        conn, aluno_id, tarefa_id, str(data_execucao),
                        ch_por_tarefa, None, 0, acertos_dist[idx],
                        comentario, questoes_dist[idx], STATUS_CONCLUIDA,
                    )
    except Exception as exc:
        erro_usuario(f"Importação cancelada para {caminho_excel.name}.", exc)
        return False
    limpar_cache()
    return True


def _status_ciclo_para_interno(status_planilha: str) -> str:
    """Converte status da planilha Ciclo Consolidado para o status interno."""
    mapa = {
        "concluido":     STATUS_CONCLUIDA,
        "concluído":     STATUS_CONCLUIDA,
        "em andamento":  STATUS_EM_ANDAMENTO,
        "nao iniciado":  STATUS_NAO_INICIADA,
        "não iniciado":  STATUS_NAO_INICIADA,
    }
    return mapa.get(chave_texto(str(status_planilha or "")), STATUS_NAO_INICIADA)


def _data_segura(valor) -> str | None:
    """Converte datetime/string do Excel para 'YYYY-MM-DD' ou None."""
    if valor is None:
        return None
    if isinstance(valor, str) and valor.strip() in ("", "—", "-"):
        return None
    try:
        ts = pd.to_datetime(valor)
        if pd.isnull(ts):
            return None
        return str(ts.date())
    except Exception:
        return None


def importar_ciclo_consolidado(caminho_excel, modo: str = "substituir") -> dict:
    """
    Importa o arquivo 'Ciclo Consolidado' (formato com dois alunos por linha).

    Estrutura esperada — aba CICLO_CONSOLIDADO, linha de cabeçalho na linha 3:
        col 0  BLOCO
        col 1  DISCIPLINA
        col 2  OBJETIVO DO BLOCO
        col 3  STATUS          (Concluído / Em Andamento / Não Iniciado)
        col 4  DATA ESTUDO (PROGRAMADA)
        col 5  (mesclada — ignorada)
        col 6  CH EFETIVA (LILIAN)
        col 7  AULA ATUAL (LILIAN)
        col 8  QUESTÕES (L)
        col 9  ACERTOS (L)
        col 10 CH EFETIVA (JESSICA)
        col 11 AULA ATUAL (JESSICA)
        col 12 QUESTÕES (J)
        col 13 ACERTOS (J)
        col 14 CH TOTAL
        col 15 TOTAL QUESTÕES
        col 16 TOTAL ACERTOS
        col 17 DESEMPENHO (%)

    Retorna dict com {ok, registros, erros, avisos}.
    """
    caminho_excel = Path(caminho_excel)
    resultado = {"ok": False, "registros": 0, "erros": [], "avisos": []}

    if not caminho_excel.exists():
        resultado["erros"].append(f"Arquivo não encontrado: {caminho_excel}")
        return resultado

    # ── Detecta os nomes dos alunos na linha 2 (índice 1) ──
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(caminho_excel), read_only=True, data_only=True)
    except Exception as exc:
        resultado["erros"].append(f"Erro ao abrir arquivo: {exc}")
        return resultado

    if "CICLO_CONSOLIDADO" not in wb.sheetnames:
        resultado["erros"].append("Aba 'CICLO_CONSOLIDADO' não encontrada. Verifique o arquivo.")
        return resultado

    ws = wb["CICLO_CONSOLIDADO"]

    # Linha 2 (índice 1) → identificação dos alunos
    # Formato: ['IDENTIFICAÇÃO', None, ..., 'LILIAN', None, ..., 'JESSICA', None, ..., 'CONSOLIDADO', ...]
    todas_linhas = list(ws.iter_rows(values_only=True))
    if len(todas_linhas) < 4:
        resultado["erros"].append("Arquivo com poucas linhas — formato inesperado.")
        return resultado

    linha_identificacao = list(todas_linhas[1])  # linha 2
    # Detecta nomes de alunos nas colunas 6 e 10 (posições fixas do formato)
    nome_aluno_a = limpar_texto(linha_identificacao[6]) if len(linha_identificacao) > 6 else None
    nome_aluno_b = limpar_texto(linha_identificacao[10]) if len(linha_identificacao) > 10 else None

    if not nome_aluno_a:
        nome_aluno_a = "Aluno A"
        resultado["avisos"].append("Nome do primeiro aluno não detectado — usando 'Aluno A'.")
    if not nome_aluno_b:
        nome_aluno_b = "Aluno B"
        resultado["avisos"].append("Nome do segundo aluno não detectado — usando 'Aluno B'.")

    # Linhas de dados: a partir da linha 4 (índice 3), pula cabeçalho e linhas de bloco sem disciplina
    linhas_dados = []
    bloco_atual = None
    for linha in todas_linhas[3:]:
        bloco_val = limpar_texto(linha[0]) if linha[0] else None
        disc_val  = limpar_texto(linha[1]) if len(linha) > 1 else None

        # Atualiza bloco atual (ffill)
        if bloco_val and bloco_val != "TOTAIS GERAIS":
            bloco_atual = bloco_val

        # Só processa linhas com disciplina e bloco válido
        if not disc_val or not bloco_atual:
            continue
        if bloco_atual == "TOTAIS GERAIS":
            continue

        linhas_dados.append((bloco_atual, linha))

    if not linhas_dados:
        resultado["erros"].append("Nenhuma linha de dado encontrada no arquivo.")
        return resultado

    # ── Garante alunos no banco ──
    try:
        with conectar() as conn:
            for nome in [nome_aluno_a, nome_aluno_b]:
                # Verifica se já existe pelo nome antes de tentar inserir,
                # evitando conflito de unique no email em PostgreSQL.
                ja_existe = conn.execute(
                    "SELECT id FROM alunos WHERE nome = ?", (nome,)
                ).fetchone()
                if ja_existe:
                    # Apenas reativa se estiver inativo
                    conn.execute(
                        "UPDATE alunos SET ativo = 1 WHERE nome = ?", (nome,)
                    )
                else:
                    # Gera email e verifica se também já está em uso
                    email = email_local(nome)
                    email_em_uso = conn.execute(
                        "SELECT id FROM alunos WHERE email = ?", (email,)
                    ).fetchone()
                    if email_em_uso:
                        # Usa email alternativo com sufixo numérico
                        import time as _time
                        email = f"{email_local(nome).split('@')[0]}_{int(_time.time()) % 10000}@local.com"
                    conn.execute(
                        """
                        INSERT INTO alunos (nome, email, senha, perfil, ativo, force_troca_senha)
                        VALUES (?, ?, ?, 'Aluno', 1, 1)
                        """,
                        (nome, email, hash_senha("123")),
                    )

            id_a = conn.execute("SELECT id FROM alunos WHERE nome = ?", (nome_aluno_a,)).fetchone()[0]
            id_b = conn.execute("SELECT id FROM alunos WHERE nome = ?", (nome_aluno_b,)).fetchone()[0]

            if modo == "substituir":
                conn.execute("DELETE FROM execucoes WHERE aluno_id IN (?, ?)", (id_a, id_b))

            # Carrega tarefas do banco indexadas por (trilha, chave_disciplina)
            tarefas_db = conn.execute(
                """
                SELECT t.id, COALESCE(t.trilha, 0), d.nome,
                       COALESCE(t.conteudo, ''), COALESCE(t.qtd_exercicios_previstos, 0)
                FROM tarefas t
                JOIN disciplinas d ON d.id = t.disciplina_id
                WHERE t.ativo = 1 AND d.ativo = 1
                ORDER BY t.numero
                """
            ).fetchall()

            # Aliases de normalização para disciplinas com variações de nome
            aliases_disc = {
                "matematica e raciocinio logico":               "matematica e raciocinio logico",
                "matematica e raciocínio lógico":               "matematica e raciocinio logico",
                "resolucao de questoes":                        "resolucao de questoes",
                "resolução de questões":                        "resolucao de questoes",
                "resolucao de questoes erradas":                "resolucao de questoes erradas",
                "resolução de questões erradas":                "resolucao de questoes erradas",
                "administracao geral e publica":                "administracao geral e publica",
                "administração geral e pública":                "administracao geral e publica",
                "administracao financeira e orcamentaria":      "administracao financeira e orcamentaria",
                "administração financeira e orçamentária":      "administracao financeira e orcamentaria",
                "contabilidade publica":                        "contabilidade publica",
                "contabilidade pública":                        "contabilidade publica",
            }

            por_chave: dict = {}
            for tid, trilha, disc_nome, conteudo, previstos in tarefas_db:
                ck = aliases_disc.get(chave_texto(disc_nome), chave_texto(disc_nome))
                por_chave.setdefault((int(trilha or 0), ck), []).append(
                    (int(tid), int(previstos or 0), limpar_texto(conteudo) or "")
                )

            registros_gravados = 0

            for bloco, linha in linhas_dados:
                # Extrai número do bloco → trilha
                m = re.search(r"(\d+)", str(bloco))
                if not m:
                    resultado["avisos"].append(f"Bloco sem número: '{bloco}' — ignorado.")
                    continue
                trilha_num = int(m.group(1))

                disciplina_raw = limpar_texto(linha[1])
                ck_disc = aliases_disc.get(chave_texto(disciplina_raw), chave_texto(disciplina_raw))

                status_planilha = limpar_texto(linha[3]) if len(linha) > 3 else None
                status_interno  = _status_ciclo_para_interno(status_planilha or "")
                data_programada = _data_segura(linha[4] if len(linha) > 4 else None)

                # Dados Aluno A
                ch_a      = converter_horas(linha[6]  if len(linha) > 6  else None)
                aula_a    = limpar_texto(linha[7]  if len(linha) > 7  else None)
                quest_a   = converter_inteiro(linha[8]  if len(linha) > 8  else None)
                acert_a   = converter_inteiro(linha[9]  if len(linha) > 9  else None)

                # Dados Aluno B
                ch_b      = converter_horas(linha[10] if len(linha) > 10 else None)
                aula_b    = limpar_texto(linha[11] if len(linha) > 11 else None)
                quest_b   = converter_inteiro(linha[12] if len(linha) > 12 else None)
                acert_b   = converter_inteiro(linha[13] if len(linha) > 13 else None)

                # Encontra tarefas correspondentes no banco
                tarefas_match = por_chave.get((trilha_num, ck_disc), [])
                if not tarefas_match:
                    resultado["avisos"].append(
                        f"Sem tarefa no banco para Bloco {trilha_num} / {disciplina_raw} — linha ignorada."
                    )
                    continue

                # Distribui questões/acertos proporcionalmente entre as tarefas do bloco
                pesos = [p for _, p, _ in tarefas_match]

                for aluno_id, ch_aluno, quest_aluno, acert_aluno, aula_aluno in [
                    (id_a, ch_a, quest_a, acert_a, aula_a),
                    (id_b, ch_b, quest_b, acert_b, aula_b),
                ]:
                    tem_dado = ch_aluno > 0 or quest_aluno > 0 or acert_aluno > 0 or bool(aula_aluno)
                    if not tem_dado and status_interno == STATUS_NAO_INICIADA:
                        # Garante que o vínculo existe como NAO_INICIADA sem sobrescrever dados
                        for tid, _, _ in tarefas_match:
                            upsert_execucao(
                                conn, aluno_id, tid,
                                data_programada, 0, None, 0, 0,
                                None, 0, STATUS_NAO_INICIADA,
                            )
                        continue

                    qtd_tarefas = len(tarefas_match)
                    ch_por = ch_aluno / qtd_tarefas if qtd_tarefas and ch_aluno else 0
                    qdist  = distribuir_inteiro(quest_aluno, pesos)
                    adist  = distribuir_inteiro(acert_aluno, qdist)

                    for idx, (tid, _, _) in enumerate(tarefas_match):
                        upsert_execucao(
                            conn, aluno_id, tid,
                            data_programada,
                            ch_por, None, 0,
                            adist[idx],
                            aula_aluno,
                            qdist[idx],
                            status_interno,
                        )
                        registros_gravados += 1

    except Exception as exc:
        resultado["erros"].append(f"Erro durante importação: {exc}")
        if DEBUG:
            import traceback
            resultado["erros"].append(traceback.format_exc())
        return resultado

    limpar_cache()
    resultado["ok"] = True
    resultado["registros"] = registros_gravados
    return resultado


# ─────────────────────────────────────────────
# CONSULTAS COM CACHE
# ─────────────────────────────────────────────

@st.cache_data(ttl=20)
def carregar_visao_tarefas(aluno_id=None):
    filtro = ""
    params = []
    if aluno_id:
        filtro = "AND e.aluno_id = ?"
        params.append(int(aluno_id))
    return consultar(
        f"""
        SELECT
            t.id AS tarefa_id,
            t.numero AS tarefa,
            t.trilha,
            d.nome AS disciplina,
            t.disciplina_id,
            t.aula_id,
            t.assunto_id,
            COALESCE(ass.titulo, t.conteudo) AS assunto,
            t.seq_disciplina,
            COALESCE(a.aula, t.aula) AS aula,
            t.qtd_exercicios_previstos,
            COALESCE(e.tipo_estudo, t.tipo, 'Outro') AS tipo,
            t.conteudo,
            e.id AS execucao_id,
            e.aluno_id,
            al.nome AS aluno,
            e.data_execucao,
            COALESCE(e.ch_efetiva, 0) AS ch_efetiva,
            e.data_revisao_24h,
            COALESCE(e.ch_revisao, 0) AS ch_revisao,
            COALESCE(e.qtd_questoes_feitas, 0) AS qtd_questoes_feitas,
            COALESCE(e.qtd_acertos, 0) AS qtd_acertos,
            COALESCE(e.desempenho, 0) AS desempenho,
            COALESCE(e.status, 'NAO_INICIADA') AS status,
            e.comentario,
            e.atualizado_em,
            CASE WHEN COALESCE(e.status, 'NAO_INICIADA') = 'CONCLUIDA' THEN 1 ELSE 0 END AS concluida
        FROM tarefas t
        JOIN disciplinas d ON d.id = t.disciplina_id AND d.ativo = 1
        LEFT JOIN aulas a ON a.id = t.aula_id AND a.ativo = 1
        LEFT JOIN assuntos ass ON ass.id = t.assunto_id AND ass.ativo = 1
        LEFT JOIN execucoes e ON e.tarefa_id = t.id {filtro}
        LEFT JOIN alunos al ON al.id = e.aluno_id
        WHERE t.ativo = 1
        ORDER BY t.numero
        """,
        tuple(params),
    )


@st.cache_data(ttl=20)
def carregar_execucoes():
    return consultar(
        """
        SELECT
            e.id AS execucao_id,
            e.aluno_id,
            al.nome AS aluno,
            e.tarefa_id,
            t.numero AS tarefa,
            t.trilha,
            d.nome AS disciplina,
            t.disciplina_id,
            COALESCE(a.aula, t.aula) AS aula,
            t.aula_id,
            COALESCE(ass.titulo, t.conteudo) AS assunto,
            t.assunto_id,
            COALESCE(e.tipo_estudo, t.tipo, 'Outro') AS tipo,
            t.conteudo,
            t.qtd_exercicios_previstos,
            e.data_execucao,
            e.ch_efetiva,
            e.data_revisao_24h,
            e.ch_revisao,
            e.qtd_questoes_feitas,
            e.qtd_acertos,
            e.desempenho,
            COALESCE(e.status, 'NAO_INICIADA') AS status,
            e.comentario,
            e.atualizado_em
        FROM execucoes e
        JOIN alunos al ON al.id = e.aluno_id AND al.ativo = 1 AND al.perfil = 'Aluno'
        JOIN tarefas t ON t.id = e.tarefa_id AND t.ativo = 1
        JOIN disciplinas d ON d.id = t.disciplina_id AND d.ativo = 1
        LEFT JOIN aulas a ON a.id = t.aula_id AND a.ativo = 1
        LEFT JOIN assuntos ass ON ass.id = t.assunto_id AND ass.ativo = 1
        ORDER BY e.atualizado_em DESC, al.nome, t.numero
        """
    )


@st.cache_data(ttl=60)
def carregar_tarefas_base():
    return consultar(
        """
        SELECT
            t.id AS tarefa_id,
            t.numero AS tarefa,
            t.trilha,
            t.disciplina_id,
            d.nome AS disciplina,
            t.aula_id,
            COALESCE(a.aula, t.aula) AS aula,
            t.assunto_id,
            COALESCE(ass.titulo, t.conteudo) AS assunto,
            t.seq_disciplina,
            t.qtd_exercicios_previstos,
            COALESCE(t.tipo, 'Outro') AS tipo,
            t.conteudo
        FROM tarefas t
        JOIN disciplinas d ON d.id = t.disciplina_id AND d.ativo = 1
        LEFT JOIN aulas a ON a.id = t.aula_id AND a.ativo = 1
        LEFT JOIN assuntos ass ON ass.id = t.assunto_id AND ass.ativo = 1
        WHERE t.ativo = 1
        ORDER BY t.numero
        """
    )


@st.cache_data(ttl=60)
def carregar_aulas():
    return consultar(
        """
        SELECT a.id, a.disciplina_id, d.nome AS disciplina, a.aula,
               COALESCE(a.tipo_estudo, 'Outro') AS tipo_estudo,
               a.estudada_padrao, a.revisao_24h_padrao
        FROM aulas a
        JOIN disciplinas d ON d.id = a.disciplina_id
        WHERE a.ativo = 1 AND d.ativo = 1
        ORDER BY d.nome, a.aula
        """
    )


@st.cache_data(ttl=60)
def carregar_assuntos():
    return consultar(
        """
        SELECT ass.id, ass.aula_id, d.nome AS disciplina, a.aula, ass.titulo AS assunto
        FROM assuntos ass
        JOIN aulas a ON a.id = ass.aula_id
        JOIN disciplinas d ON d.id = a.disciplina_id
        WHERE ass.ativo = 1 AND a.ativo = 1 AND d.ativo = 1
        ORDER BY d.nome, a.aula, ass.titulo
        """
    )


def alunos_ativos(incluir_gestor=False):
    filtro = "" if incluir_gestor else "AND perfil = 'Aluno'"
    return consultar(f"SELECT id, nome, email, perfil FROM alunos WHERE ativo = 1 {filtro} ORDER BY nome")


def aluno_logado():
    return st.session_state["usuario"]


def disciplinas_ativas():
    return consultar("SELECT id, nome FROM disciplinas WHERE ativo = 1 ORDER BY nome")


# ─────────────────────────────────────────────
# FILTROS E MÉTRICAS
# ─────────────────────────────────────────────

def periodo_datas(periodo):
    hoje = date.today()
    if periodo == "Hoje":
        return hoje, hoje
    if periodo == "Semana":
        return hoje - timedelta(days=hoje.weekday()), hoje
    if periodo == "Mês":
        return hoje.replace(day=1), hoje
    if periodo == "Ano":
        return hoje.replace(month=1, day=1), hoje
    return None, None


def painel_filtros(df, prefixo="dash"):
    st.sidebar.markdown("### Filtros")
    if st.sidebar.button("Limpar filtros", use_container_width=True, key=f"{prefixo}_limpar"):
        for chave in list(st.session_state.keys()):
            if chave.startswith(f"{prefixo}_"):
                del st.session_state[chave]
        st.rerun()

    favoritos = st.session_state.setdefault("filtros_favoritos", {})
    favoritos_opcoes = ["Nenhum"] + list(favoritos.keys())
    favorito = st.sidebar.selectbox("Filtro favorito", favoritos_opcoes, key=f"{prefixo}_favorito")
    if favorito != "Nenhum" and st.sidebar.button("Aplicar favorito", use_container_width=True, key=f"{prefixo}_aplicar"):
        for chave, valor in favoritos[favorito].items():
            st.session_state[f"{prefixo}_{chave}"] = valor
        st.rerun()

    periodo = st.sidebar.selectbox(
        "Período", ["Todos", "Hoje", "Semana", "Mês", "Ano", "Personalizado"], key=f"{prefixo}_periodo"
    )
    inicio_padrao, fim_padrao = periodo_datas(periodo)
    if periodo == "Personalizado":
        inicio = st.sidebar.date_input("Data inicial", value=inicio_padrao or date.today(), key=f"{prefixo}_inicio")
        fim = st.sidebar.date_input("Data final", value=fim_padrao or date.today(), key=f"{prefixo}_fim")
    else:
        inicio, fim = inicio_padrao, fim_padrao

    usuario = aluno_logado()
    alunos = ["Todos"] + sorted(df["aluno"].dropna().unique().tolist())
    if usuario["perfil"] == "Aluno":
        aluno = usuario["nome"]
        st.sidebar.text_input("Aluno", value=aluno, disabled=True)
    else:
        aluno = st.sidebar.selectbox("Aluno", alunos, key=f"{prefixo}_aluno")

    disciplinas = sorted(df["disciplina"].dropna().unique().tolist())
    assuntos    = sorted(df["assunto"].dropna().unique().tolist())
    aulas       = sorted(df["aula"].dropna().unique().tolist())
    tipos       = sorted(df["tipo"].dropna().unique().tolist())

    status_escolhidos = st.sidebar.multiselect(
        "Status", STATUS_VALIDOS, default=STATUS_VALIDOS,
        format_func=lambda v: STATUS_LABELS.get(v, v), key=f"{prefixo}_status",
    )
    disciplina    = st.sidebar.multiselect("Disciplinas", disciplinas, key=f"{prefixo}_disciplina")
    aula          = st.sidebar.multiselect("Aulas", aulas, key=f"{prefixo}_aula")
    assunto       = st.sidebar.multiselect("Assuntos", assuntos, key=f"{prefixo}_assunto")
    tipo          = st.sidebar.multiselect("Tipos de estudo", tipos, key=f"{prefixo}_tipo")
    minimo_horas  = st.sidebar.number_input("Tempo mínimo estudado", min_value=0.0, value=0.0, step=0.25, key=f"{prefixo}_min_horas")
    recentes      = st.sidebar.toggle("Atividades recentes", value=False, key=f"{prefixo}_recentes")

    nome_filtro = st.sidebar.text_input("Nome para salvar filtro", key=f"{prefixo}_nome_filtro")
    if st.sidebar.button("Salvar filtro favorito", use_container_width=True, key=f"{prefixo}_salvar_filtro"):
        if nome_filtro:
            favoritos[nome_filtro] = {
                "periodo": periodo, "aluno": aluno, "status": status_escolhidos,
                "disciplina": disciplina, "aula": aula, "assunto": assunto,
                "tipo": tipo, "min_horas": minimo_horas, "recentes": recentes,
            }
            _toast_sucesso("Filtro salvo como favorito.")

    filtrado = df.copy()
    filtrado["data_ref"] = pd.to_datetime(filtrado["data_execucao"], errors="coerce")
    if inicio:
        filtrado = filtrado[filtrado["data_ref"].dt.date >= inicio]
    if fim:
        filtrado = filtrado[filtrado["data_ref"].dt.date <= fim]
    if aluno != "Todos":
        filtrado = filtrado[filtrado["aluno"] == aluno]
    if status_escolhidos:
        filtrado = filtrado[filtrado["status"].isin(status_escolhidos)]
    if disciplina:
        filtrado = filtrado[filtrado["disciplina"].isin(disciplina)]
    if aula:
        filtrado = filtrado[filtrado["aula"].isin(aula)]
    if assunto:
        filtrado = filtrado[filtrado["assunto"].isin(assunto)]
    if tipo:
        filtrado = filtrado[filtrado["tipo"].isin(tipo)]
    if minimo_horas > 0:
        filtrado = filtrado[filtrado["ch_efetiva"] >= minimo_horas]
    if recentes:
        limite = pd.Timestamp.now() - pd.Timedelta(days=15)
        filtrado = filtrado[filtrado["data_ref"] >= limite]
    return filtrado, aluno


def base_metricas(df):
    return df[df["status"].isin(STATUS_ANALISE)].copy()


def calcular_kpis_avancados(df: pd.DataFrame) -> dict:
    """Calcula todos os KPIs analíticos de forma centralizada."""
    hoje = date.today()
    semana_ini = hoje - timedelta(days=hoje.weekday())
    semana_pas = semana_ini - timedelta(days=7)

    df = df.copy()
    df["data_ref"] = pd.to_datetime(df.get("data_execucao", pd.Series(dtype=str)), errors="coerce")

    analisavel = df[df["status"].isin(STATUS_ANALISE)].copy()
    concluidas = analisavel[analisavel["status"] == STATUS_CONCLUIDA]
    andamento  = analisavel[analisavel["status"] == STATUS_EM_ANDAMENTO]
    nao_inic   = df[df["status"] == STATUS_NAO_INICIADA]

    total_tarefas    = len(df)
    qtd_concluidas   = len(concluidas)
    qtd_andamento    = len(andamento)
    qtd_nao_iniciada = len(nao_inic)
    pct_conclusao    = qtd_concluidas / total_tarefas * 100 if total_tarefas else 0

    horas_total = float(analisavel["ch_efetiva"].sum())
    questoes    = int(analisavel["qtd_questoes_feitas"].sum())
    acertos     = int(analisavel["qtd_acertos"].sum())
    desempenho  = acertos / questoes * 100 if questoes else 0

    dias_ativos_series = analisavel["data_ref"].dropna().dt.date
    dias_unicos  = sorted(dias_ativos_series.unique()) if len(dias_ativos_series) else []
    qtd_dias     = len(dias_unicos)
    media_diaria = horas_total / qtd_dias if qtd_dias else 0

    exec_semana  = analisavel[analisavel["data_ref"].dt.date >= semana_ini]
    horas_semana = float(exec_semana["ch_efetiva"].sum())
    conc_semana  = int((exec_semana["status"] == STATUS_CONCLUIDA).sum())

    exec_sem_pas  = analisavel[(analisavel["data_ref"].dt.date >= semana_pas) & (analisavel["data_ref"].dt.date < semana_ini)]
    horas_sem_pas = float(exec_sem_pas["ch_efetiva"].sum())
    conc_sem_pas  = int((exec_sem_pas["status"] == STATUS_CONCLUIDA).sum())

    sequencia = 0
    if dias_unicos:
        d = hoje
        while d in dias_unicos:
            sequencia += 1; d -= timedelta(days=1)

    dias_sem_estudar = (hoje - max(dias_unicos)).days if dias_unicos else 0
    produtividade    = qtd_concluidas / horas_total if horas_total else 0
    ritmo_semanal    = conc_semana if conc_semana > 0 else (qtd_concluidas / max(1, qtd_dias) * 7)
    tarefas_rest     = qtd_andamento + qtd_nao_iniciada
    semanas_rest     = tarefas_rest / ritmo_semanal if ritmo_semanal > 0 else None
    previsao         = (hoje + timedelta(weeks=semanas_rest)) if semanas_rest is not None else None
    media_h_tarefa   = horas_total / qtd_concluidas if qtd_concluidas else 0
    horas_restantes  = tarefas_rest * media_h_tarefa

    return {
        "analisavel": analisavel, "concluidas_df": concluidas, "andamento_df": andamento,
        "total_tarefas": total_tarefas, "qtd_concluidas": qtd_concluidas,
        "qtd_andamento": qtd_andamento, "qtd_nao_iniciada": qtd_nao_iniciada,
        "pct_conclusao": pct_conclusao, "horas_total": horas_total,
        "horas_semana": horas_semana, "horas_sem_pas": horas_sem_pas,
        "delta_horas_semana": horas_semana - horas_sem_pas,
        "conc_semana": conc_semana, "conc_sem_pas": conc_sem_pas,
        "delta_conc_semana": conc_semana - conc_sem_pas,
        "questoes": questoes, "acertos": acertos, "desempenho": desempenho,
        "qtd_dias_ativos": qtd_dias, "media_diaria": media_diaria,
        "sequencia": sequencia, "dias_sem_estudar": dias_sem_estudar,
        "produtividade": produtividade, "ritmo_semanal": ritmo_semanal,
        "tarefas_restantes": tarefas_rest, "previsao_conclusao": previsao,
        "horas_restantes": horas_restantes, "dias_unicos": dias_unicos,
    }


def calcular_metricas(df):
    """Compatível com código legado."""
    k = calcular_kpis_avancados(df)
    return {
        "analisavel": k["analisavel"], "concluidas_df": k["concluidas_df"],
        "andamento_df": k["andamento_df"], "total": k["total_tarefas"],
        "concluidas": k["qtd_concluidas"], "andamento": k["qtd_andamento"],
        "nao_iniciadas": k["qtd_nao_iniciada"], "progresso": k["pct_conclusao"],
        "horas": k["horas_total"], "questoes": k["questoes"], "acertos": k["acertos"],
        "desempenho": k["desempenho"], "media_dia": k["media_diaria"],
        "produtividade": k["produtividade"],
    }


def gerar_insights(df: pd.DataFrame, kpis: dict, nome_aluno: str = "") -> list:
    """Gera insights analíticos contextuais baseados nos dados do aluno."""
    insights = []
    ana = kpis["analisavel"]
    nome = nome_aluno or "O aluno"
    if ana.empty:
        return [{"tipo":"info","icone":"📭","titulo":"Sem dados suficientes","texto":"Registre atividades para obter análise personalizada."}]

    # Tendência de horas
    dh = kpis["delta_horas_semana"]
    hs = kpis["horas_semana"]; hsp = kpis["horas_sem_pas"]
    if hsp > 0:
        pct_d = dh / hsp * 100
        if pct_d < -20:
            insights.append({"tipo":"warning","icone":"📉","titulo":"Queda de produtividade",
                "texto":f"{nome} estudou {abs(pct_d):.0f}% menos esta semana ({hs:.1f}h) vs semana anterior ({hsp:.1f}h)."})
        elif pct_d > 20:
            insights.append({"tipo":"success","icone":"📈","titulo":"Semana mais produtiva",
                "texto":f"{nome} aumentou {pct_d:.0f}% as horas esta semana ({hs:.1f}h vs {hsp:.1f}h). Excelente ritmo!"})

    # Sequência
    seq = kpis["sequencia"]
    if seq >= 7:
        insights.append({"tipo":"success","icone":"🔥","titulo":f"Sequência de {seq} dias!",
            "texto":f"{nome} está estudando há {seq} dias consecutivos. Consistência é fundamental para aprovação."})
    elif seq == 0 and kpis["dias_sem_estudar"] >= 3:
        insights.append({"tipo":"warning","icone":"⏰","titulo":"Dias sem estudar",
            "texto":f"{nome} não registra atividade há {kpis['dias_sem_estudar']} dias. Retome a rotina para não perder o ritmo."})

    # Disciplinas
    if "disciplina" in ana.columns and len(ana["disciplina"].unique()) > 1:
        disc_agg = ana.groupby("disciplina", as_index=False).agg(
            questoes=("qtd_questoes_feitas","sum"), acertos=("qtd_acertos","sum"),
            horas=("ch_efetiva","sum"), tarefas=("tarefa_id","count"),
            concluidas=("status", lambda s:(s==STATUS_CONCLUIDA).sum()))
        disc_agg["desempenho"] = disc_agg.apply(lambda r: r["acertos"]/r["questoes"]*100 if r["questoes"] else 0, axis=1)
        disc_agg["pct_conc"]   = disc_agg.apply(lambda r: r["concluidas"]/r["tarefas"]*100 if r["tarefas"] else 0, axis=1)

        criticas = disc_agg[(disc_agg["questoes"]>0) & (disc_agg["desempenho"]<60)].sort_values("desempenho")
        for _, r in criticas.head(2).iterrows():
            insights.append({"tipo":"danger","icone":"🚨","titulo":f"Atenção: {r['disciplina']}",
                "texto":f"Desempenho de {r['desempenho']:.1f}% em {r['disciplina']} abaixo de 60%. Dedique sessões de revisão e questões comentadas."})

        total_h = disc_agg["horas"].sum()
        if total_h > 0:
            neg = disc_agg[(disc_agg["horas"]/total_h < 0.05) & (disc_agg["tarefas"]>0)]
            for _, r in neg.head(2).iterrows():
                insights.append({"tipo":"warning","icone":"📌","titulo":f"Disciplina negligenciada: {r['disciplina']}",
                    "texto":f"Apenas {r['horas']:.1f}h em {r['disciplina']} (<5% do tempo). Reequilibre a distribuição."})

        avancadas = disc_agg[disc_agg["pct_conc"]>=80].sort_values("pct_conc", ascending=False)
        if not avancadas.empty:
            nomes = ", ".join(avancadas.head(3)["disciplina"].tolist())
            insights.append({"tipo":"success","icone":"🏆","titulo":"Disciplinas avançadas",
                "texto":f"Ótimo progresso em: {nomes}. Continue com revisões."})

    # Risco de atraso
    prev = kpis.get("previsao_conclusao")
    if prev and kpis["tarefas_restantes"] > 0:
        dias_p = (prev - date.today()).days
        if dias_p > 365:
            insights.append({"tipo":"danger","icone":"⚠️","titulo":"Risco de atraso alto",
                "texto":f"No ritmo atual, conclusão estimada em {prev.strftime('%d/%m/%Y')} ({dias_p} dias). Considere aumentar a carga semanal."})
        elif dias_p > 180:
            insights.append({"tipo":"warning","icone":"📅","titulo":"Previsão de conclusão",
                "texto":f"Estimativa: {prev.strftime('%d/%m/%Y')} ({dias_p} dias). Mantenha o ritmo."})

    # Produtividade
    prod = kpis["produtividade"]
    if prod >= 1.5:
        insights.append({"tipo":"success","icone":"⚡","titulo":"Alta produtividade",
            "texto":f"{prod:.2f} tarefas/hora — excelente eficiência."})
    elif 0 < prod < 0.3:
        insights.append({"tipo":"info","icone":"🕐","titulo":"Sessões longas, pouco avanço",
            "texto":f"Produtividade de {prod:.2f} tarefas/hora. Experimente sessões mais curtas com foco (técnica Pomodoro)."})

    # Consistência
    md = kpis["media_diaria"]
    if md >= 4:
        insights.append({"tipo":"success","icone":"📚","titulo":"Consistência acima da média",
            "texto":f"Média de {md:.1f}h/dia ativo — ritmo sólido para aprovação."})
    elif 0 < md < 1:
        insights.append({"tipo":"warning","icone":"⏱️","titulo":"Carga diária baixa",
            "texto":f"Média de {md:.1f}h/dia ativo. Para concursos competitivos, recomenda-se ≥4h/dia."})

    if not insights:
        insights.append({"tipo":"info","icone":"✅","titulo":"Estudos em dia",
            "texto":f"{nome} não apresenta alertas críticos. Continue monitorando o progresso."})
    return insights


def fig_layout(fig, height=320):
    """Aplica estilo visual padrão a qualquer figura Plotly."""
    fig.update_layout(
        legend_title_text="Legenda",
        margin=dict(l=10, r=10, t=46, b=10),
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#0f172a", family="Inter, sans-serif"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.35, xanchor="center", x=0.5),
    )
    fig.update_xaxes(showgrid=False, linecolor="#e2e8f0", tickfont=dict(size=10))
    fig.update_yaxes(showgrid=True, gridcolor="#f1f5f9", linecolor="rgba(0,0,0,0)", tickfont=dict(size=10))
    return fig


def grafico_vazio(msg="Sem dados suficientes para este gráfico."):
    """Retorna figura vazia com mensagem amigável."""
    fig = go.Figure()
    fig.add_annotation(text=msg, x=0.5, y=0.5, xref="paper", yref="paper",
        showarrow=False, font=dict(size=13, color="#94a3b8"))
    fig.update_layout(height=260, paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)", xaxis=dict(visible=False), yaxis=dict(visible=False))
    return fig


def preparar_tabela(df):
    tabela = df.copy()
    if "data_execucao" in tabela:
        tabela["data_execucao"] = tabela["data_execucao"].apply(formatar_data_br)
    if "status" in tabela:
        tabela["status_label"] = tabela["status"].map(STATUS_LABELS)
    return tabela


# ─────────────────────────────────────────────
# TELA: LOGIN
# ─────────────────────────────────────────────

def tela_login():
    render_html(f"""
        <div class="hero">
          <h1>{APP_NAME}</h1>
          <p>Acompanhamento de estudos, produtividade, tarefas educacionais e desempenho por aluno.</p>
        </div>
    """)
    with st.expander("Usuário inicial"):
        st.write(f"Gestor: **{ADMIN_EMAIL}** / senha **{ADMIN_PASSWORD}**")
    with st.form("login"):
        email = st.text_input("E-mail").strip().lower()
        senha = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("Entrar", use_container_width=True)
    if entrar:
        usuario = consultar("SELECT * FROM alunos WHERE email = ? AND ativo = 1", (email,))
        if usuario.empty or not verificar_senha(senha, usuario.iloc[0]["senha"]):
            st.error("Usuário ou senha inválidos.")
            return
        st.session_state["usuario"] = usuario.iloc[0].to_dict()
        st.rerun()


def tela_troca_obrigatoria():
    usuario = aluno_logado()
    render_html(f"""
        <div class="hero">
          <h1>{APP_NAME}</h1>
          <p>Por segurança, altere sua senha inicial antes de acessar o sistema.</p>
        </div>
    """)
    with st.form("troca_obrigatoria"):
        nova = st.text_input("Nova senha", type="password")
        confirmar = st.text_input("Confirmar nova senha", type="password")
        salvar = st.form_submit_button("Alterar senha e entrar", use_container_width=True)
    if salvar:
        if nova != confirmar:
            st.error("A confirmação não confere.")
        elif not senha_valida(nova):
            st.error("Use uma senha com pelo menos 8 caracteres, contendo letras e números.")
        else:
            atualizar_senha_usuario(usuario["id"], nova, 0)
            usuario_atualizado = consultar("SELECT * FROM alunos WHERE id = ?", (int(usuario["id"]),))
            st.session_state["usuario"] = usuario_atualizado.iloc[0].to_dict()
            _toast_sucesso("Senha alterada com sucesso. Faça login novamente.")
            st.rerun()


# ─────────────────────────────────────────────
# TELA: DASHBOARD — funções auxiliares
# ─────────────────────────────────────────────

def _render_kpis_produtividade(kpis: dict):
    hs  = kpis["horas_semana"]
    hsp = kpis["horas_sem_pas"]
    dh  = kpis["delta_horas_semana"]
    seq = kpis["sequencia"]
    dsem = kpis["dias_sem_estudar"]
    cols = st.columns(5)
    cards = [
        kpi_card("Horas esta semana", f"{hs:.1f}h", f"Semana anterior: {hsp:.1f}h",
            f"{abs(dh):.1f}h vs sem. passada", delta_pos=(dh >= 0),
            tooltip="Total de horas desta semana (seg–hoje). Fórmula: soma de ch_efetiva com data ≥ início da semana. ▲ = mais horas que a semana anterior."),
        kpi_card("Média diária", f"{kpis['media_diaria']:.1f}h", "por dia ativo (com registro)",
            tooltip="Horas médias por dia com atividade registrada. Fórmula: total de horas ÷ dias distintos. ≥ 4h/dia = consistente para concursos."),
        kpi_card("Sequência atual", f"{seq} dia{'s' if seq!=1 else ''}", "dias consecutivos com registro",
            tooltip="Dias seguidos (até hoje) com pelo menos uma atividade. Calculado para trás a partir de hoje. Sequências longas = hábito consolidado."),
        kpi_card("Dias sem estudar", f"{dsem}", "desde o último registro",
            delta=f"{dsem}d de pausa" if dsem > 0 else "", delta_pos=False,
            tooltip="Dias desde o último registro. 0 = estudou hoje. Acima de 3 dias = alerta de pausa."),
        kpi_card("Produtividade", f"{kpis['produtividade']:.2f}", "tarefas concluídas / hora",
            tooltip="Eficiência: tarefas concluídas por hora. Fórmula: concluídas ÷ horas. > 1,0 = boa taxa. < 0,3 = sessões longas com pouco resultado."),
    ]
    for col, card in zip(cols, cards):
        with col: render_html(card)


def _render_kpis_avanco(kpis: dict):
    pct   = kpis["pct_conclusao"]
    prev  = kpis.get("previsao_conclusao")
    prev_str = prev.strftime("%d/%m/%Y") if prev else "—"
    cols = st.columns(5)
    cards = [
        kpi_card("Progresso geral", f"{pct:.1f}%",
            f"{kpis['qtd_concluidas']} de {kpis['total_tarefas']} tarefas concluídas",
            tooltip="% de tarefas concluídas. Fórmula: (concluídas ÷ total) × 100. Considera apenas status CONCLUIDA."),
        kpi_card("Em andamento", f"{kpis['qtd_andamento']}", f"+ {kpis['qtd_nao_iniciada']} não iniciadas",
            tooltip="Tarefas com status EM_ANDAMENTO. Muitas simultâneas podem indicar falta de foco."),
        kpi_card("Tarefas restantes", f"{kpis['tarefas_restantes']}", "andamento + não iniciadas",
            tooltip="Total de tarefas não concluídas. Base para estimativa de conclusão do plano."),
        kpi_card("Horas restantes (est.)", f"{kpis['horas_restantes']:.0f}h", "base: média por tarefa concluída",
            tooltip="Estimativa de horas para concluir o restante. Fórmula: (horas ÷ concluídas) × restantes. Baseado no ritmo atual."),
        kpi_card("Previsão de conclusão", prev_str, "baseado no ritmo semanal atual",
            tooltip="Data estimada de término. Fórmula: hoje + (restantes ÷ ritmo semanal) semanas. Ritmo = média de conclusões/semana."),
    ]
    for col, card in zip(cols, cards):
        with col: render_html(card)


def _render_kpis_desempenho(kpis: dict):
    des = kpis["desempenho"]; q = kpis["questoes"]; ac = kpis["acertos"]
    cs  = kpis["conc_semana"]; csp = kpis["conc_sem_pas"]; dc = kpis["delta_conc_semana"]
    cols = st.columns(4)
    cards = [
        kpi_card("Desempenho geral", f"{des:.1f}%", f"{ac} acertos em {q} questões",
            tooltip="Taxa de acerto. Fórmula: (acertos ÷ questões) × 100. ≥ 70% = satisfatório para concursos."),
        kpi_card("Questões feitas", f"{q:,}".replace(",","."), "total acumulado",
            tooltip="Soma de todas as questões feitas em atividades iniciadas/concluídas. Mais questões = maior treinamento."),
        kpi_card("Concluídas esta semana", f"{cs}", f"semana anterior: {csp}",
            delta=f"{abs(dc)} tarefa{'s' if abs(dc)!=1 else ''} vs sem. ant.", delta_pos=(dc >= 0),
            tooltip="Tarefas concluídas na semana atual (segunda a hoje). ▲ = aceleração em relação à semana passada."),
        kpi_card("Dias ativos", f"{kpis['qtd_dias_ativos']}", "dias com pelo menos 1 registro",
            tooltip="Dias distintos com algum registro. Maior número = hábito mais sólido e frequência maior."),
    ]
    for col, card in zip(cols, cards):
        with col: render_html(card)


def _aba_visao_geral(df_filtrado, analisavel):
    status_df = df_filtrado.groupby("status", as_index=False)["tarefa_id"].count().rename(columns={"tarefa_id":"qtd"})
    status_df["label"] = status_df["status"].map(STATUS_LABELS)
    col_a, col_b = st.columns(2)
    with col_a:
        render_html(_tooltip_grafico(
            "Barras verticais mostrando quantas tarefas estão em cada status. "
            "Não iniciada = fila pendente · Em andamento = em execução · Concluída = finalizada. "
            "Quanto mais barras verdes, melhor o avanço geral."
        ))
        fig = go.Figure(go.Bar(
            x=status_df["label"], y=status_df["qtd"],
            marker_color=[STATUS_CORES.get(s,"#94a3b8") for s in status_df["status"]],
            text=status_df["qtd"], textposition="outside"))
        fig.update_layout(title="Distribuição por status", showlegend=False)
        st.plotly_chart(fig_layout(fig, 300), use_container_width=True)
    with col_b:
        if not analisavel.empty and "aluno" in analisavel.columns:
            pa = analisavel.groupby("aluno", as_index=False).agg(
                horas=("ch_efetiva","sum"), concluidas=("status", lambda s:(s==STATUS_CONCLUIDA).sum())
            ).sort_values("horas", ascending=True)
            render_html(_tooltip_grafico(
                "Barras horizontais agrupadas comparando horas estudadas e tarefas concluídas por aluno. "
                "Ideal para identificar quem está mais ativo e quem está concluindo mais atividades."
            ))
            fig = go.Figure()
            fig.add_trace(go.Bar(name="Horas", y=pa["aluno"], x=pa["horas"], orientation="h", marker_color="#3b82f6"))
            fig.add_trace(go.Bar(name="Concluídas", y=pa["aluno"], x=pa["concluidas"], orientation="h", marker_color="#22c55e"))
            fig.update_layout(title="Horas × Conclusões por aluno", barmode="group")
            st.plotly_chart(fig_layout(fig, 300), use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio("Sem atividades para comparação."), use_container_width=True)


def _aba_disciplinas(analisavel):
    if analisavel.empty:
        st.info("Sem dados analisáveis."); return
    disc = analisavel.groupby("disciplina", as_index=False).agg(
        tarefas=("tarefa_id","count"), horas=("ch_efetiva","sum"),
        questoes=("qtd_questoes_feitas","sum"), acertos=("qtd_acertos","sum"),
        concluidas=("status", lambda s:(s==STATUS_CONCLUIDA).sum()),
    )
    disc["progresso"]  = disc.apply(lambda r: r["concluidas"]/r["tarefas"]*100 if r["tarefas"] else 0, axis=1)
    disc["desempenho"] = disc.apply(lambda r: r["acertos"]/r["questoes"]*100 if r["questoes"] else 0, axis=1)
    col_a, col_b = st.columns(2)
    with col_a:
        render_html(_tooltip_grafico(
            "Barra horizontal mostrando o % de tarefas concluídas por disciplina. "
            "Verde ≥ 80% · Amarelo 40–79% · Vermelho < 40%. "
            "Fórmula: (tarefas concluídas ÷ total de tarefas) × 100."
        ))
        d = disc.sort_values("progresso")
        fig = go.Figure(go.Bar(
            x=d["progresso"], y=d["disciplina"], orientation="h",
            text=d["progresso"].map(lambda v: f"{v:.0f}%"), textposition="outside",
            marker_color=d["progresso"].map(lambda v: "#22c55e" if v>=80 else ("#f59e0b" if v>=40 else "#ef4444")),
        ))
        fig.update_layout(title="Progresso por disciplina (%)")
        st.plotly_chart(fig_layout(fig, max(280, len(disc)*36)), use_container_width=True)
    with col_b:
        render_html(_tooltip_grafico(
            "Barras = horas estudadas (eixo esquerdo). "
            "Linha = taxa de acerto em questões, eixo direito (0–100%). "
            "Disciplinas com muitas horas e baixo desempenho indicam estudo ineficiente."
        ))
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Horas", x=disc["disciplina"], y=disc["horas"], marker_color="#3b82f6"))
        fig.add_trace(go.Scatter(name="Desempenho (%)", x=disc["disciplina"], y=disc["desempenho"],
            mode="lines+markers", marker_color="#f59e0b", yaxis="y2", line=dict(width=2)))
        fig.update_layout(title="Horas × Desempenho", yaxis2=dict(overlaying="y", side="right", range=[0,100]))
        st.plotly_chart(fig_layout(fig, 300), use_container_width=True)
    if len(disc) >= 3:
        render_html(_tooltip_grafico(
            "Radar comparando a taxa de acerto (%) de cada disciplina. "
            "Disciplinas mais distantes do centro têm melhor desempenho. "
            "Áreas retraídas indicam onde concentrar esforços de revisão."
        ))
        fig_r = go.Figure(go.Scatterpolar(
            r=disc["desempenho"].tolist()+[disc["desempenho"].tolist()[0]],
            theta=disc["disciplina"].tolist()+[disc["disciplina"].tolist()[0]],
            fill="toself", fillcolor="rgba(59,130,246,0.15)",
            line=dict(color="#3b82f6", width=2),
        ))
        fig_r.update_layout(title="Radar de desempenho", polar=dict(radialaxis=dict(visible=True, range=[0,100])), showlegend=False)
        st.plotly_chart(fig_layout(fig_r, 360), use_container_width=True)
    st.dataframe(disc.rename(columns={"progresso":"Progresso (%)","desempenho":"Desempenho (%)","horas":"Horas","questoes":"Questões","acertos":"Acertos","tarefas":"Tarefas","concluidas":"Concluídas"}), use_container_width=True, hide_index=True)


def _aba_evolucao(analisavel):
    if analisavel.empty or "data_ref" not in analisavel.columns:
        st.info("Sem dados de evolução."); return
    evo = analisavel.dropna(subset=["data_ref"]).copy()
    if evo.empty:
        st.info("Sem datas de execução."); return
    evo["dia"] = evo["data_ref"].dt.date
    diario = evo.groupby("dia", as_index=False).agg(horas=("ch_efetiva","sum"), tarefas=("tarefa_id","count"))
    diario["dia_ts"] = pd.to_datetime(diario["dia"])
    diario["media7"] = diario["horas"].rolling(7, min_periods=1).mean()
    render_html(_tooltip_grafico(
        "Barras = horas estudadas por dia. Linha pontilhada = média móvel dos últimos 7 dias. "
        "A média móvel suaviza oscilações e revela a tendência real do ritmo de estudo. "
        "Quando a média sobe: ritmo crescente. Quando desce: possível desaceleração."
    ))
    fig = go.Figure()
    fig.add_trace(go.Bar(x=diario["dia_ts"], y=diario["horas"], name="Horas/dia", marker_color="#3b82f6", opacity=0.7))
    fig.add_trace(go.Scatter(x=diario["dia_ts"], y=diario["media7"], name="Média 7 dias", mode="lines", line=dict(color="#f59e0b", width=2, dash="dot")))
    fig.update_layout(title="Evolução diária de horas + média móvel 7 dias")
    st.plotly_chart(fig_layout(fig, 300), use_container_width=True)
    col_a, col_b = st.columns(2)
    with col_a:
        sem = evo.set_index("data_ref").resample("W").agg(horas=("ch_efetiva","sum"), concluidas=("status", lambda s:(s==STATUS_CONCLUIDA).sum())).reset_index()
        sem["label"] = sem["data_ref"].dt.strftime("Sem %d/%m")
        if not sem.empty:
            render_html(_tooltip_grafico(
                "Barras = horas por semana. Linha = tarefas concluídas na mesma semana (eixo direito). "
                "Permite ver se semanas com mais horas resultam em mais conclusões — ou se há problema de eficiência."
            ))
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(x=sem["label"], y=sem["horas"], name="Horas", marker_color="#3b82f6"))
            fig2.add_trace(go.Scatter(x=sem["label"], y=sem["concluidas"], name="Concluídas", mode="lines+markers", yaxis="y2", marker_color="#22c55e", line=dict(width=2)))
            fig2.update_layout(title="Horas × Conclusões semanais", yaxis2=dict(overlaying="y", side="right"))
            st.plotly_chart(fig_layout(fig2, 300), use_container_width=True)
    with col_b:
        ordem_dias = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
        labels_dias = {"Monday":"Seg","Tuesday":"Ter","Wednesday":"Qua","Thursday":"Qui","Friday":"Sex","Saturday":"Sáb","Sunday":"Dom"}
        evo["dia_semana"] = evo["data_ref"].dt.day_name()
        ds = evo.groupby("dia_semana", as_index=False).agg(horas=("ch_efetiva","sum"))
        ds = ds[ds["dia_semana"].isin(ordem_dias)]
        ds["dia_semana"] = pd.Categorical(ds["dia_semana"], categories=ordem_dias, ordered=True)
        ds = ds.sort_values("dia_semana")
        ds["label"] = ds["dia_semana"].map(labels_dias)
        if not ds.empty:
            render_html(_tooltip_grafico(
                "Horas totais acumuladas em cada dia da semana ao longo de todo o período. "
                "A barra verde destaca o dia com mais horas — provavelmente o dia de maior rendimento. "
                "Útil para planejar os dias de estudo mais intenso."
            ))
            fig3 = go.Figure(go.Bar(
                x=ds["label"], y=ds["horas"],
                marker_color=["#22c55e" if h==ds["horas"].max() else "#3b82f6" for h in ds["horas"]],
                text=ds["horas"].map(lambda v: f"{v:.1f}h"), textposition="outside"))
            fig3.update_layout(title="Distribuição por dia da semana")
            st.plotly_chart(fig_layout(fig3, 300), use_container_width=True)


def _aba_gestao_tempo(analisavel):
    if analisavel.empty:
        st.info("Sem dados para análise de tempo."); return
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        tipos = analisavel.groupby("tipo", as_index=False).agg(horas=("ch_efetiva","sum"))
        tipos = tipos[tipos["horas"]>0].sort_values("horas", ascending=False)
        render_html(_tooltip_grafico(
            "Pizza mostrando como as horas de estudo estão distribuídas entre os diferentes tipos de atividade "
            "(Leitura, Exercícios, Revisão, Videoaula etc.). "
            "Ideal para identificar se há desequilíbrio — ex: muita teoria e pouca prática."
        ))
        fig = px.pie(tipos, names="tipo", values="horas", color_discrete_sequence=px.colors.qualitative.Set3)
        fig.update_traces(textposition="inside", textinfo="percent+label")
        fig.update_layout(title="Distribuição por tipo de estudo", showlegend=False)
        st.plotly_chart(fig_layout(fig, 300), use_container_width=True)
    with col_b:
        bd = analisavel.groupby("disciplina", as_index=False).agg(horas=("ch_efetiva","sum")).sort_values("horas", ascending=True)
        tot = bd["horas"].sum() or 1
        bd["pct"] = bd["horas"]/tot*100
        render_html(_tooltip_grafico(
            "Barras horizontais com horas acumuladas por disciplina. "
            "O % indica a proporção de cada disciplina no tempo total de estudo. "
            "Disciplinas com % muito baixo podem estar sendo negligenciadas."
        ))
        fig2 = go.Figure(go.Bar(x=bd["horas"], y=bd["disciplina"], orientation="h",
            text=bd["pct"].map(lambda v: f"{v:.0f}%"), textposition="outside", marker_color="#3b82f6"))
        fig2.update_layout(title="Horas por disciplina")
        st.plotly_chart(fig_layout(fig2, max(280, len(bd)*36)), use_container_width=True)
    with col_c:
        ef = analisavel.groupby("tipo", as_index=False).agg(questoes=("qtd_questoes_feitas","sum"), horas=("ch_efetiva","sum"))
        ef = ef[ef["horas"]>0]
        ef["eficiencia"] = ef["questoes"]/ef["horas"]
        ef = ef.sort_values("eficiencia", ascending=False)
        if not ef.empty:
            render_html(_tooltip_grafico(
                "Quantas questões são feitas por hora em cada tipo de atividade. "
                "Fórmula: questões feitas ÷ horas de estudo. "
                "A barra verde destaca o tipo mais eficiente em volume de prática."
            ))
            fig3 = go.Figure(go.Bar(x=ef["tipo"], y=ef["eficiencia"],
                text=ef["eficiencia"].map(lambda v: f"{v:.1f}"), textposition="outside",
                marker_color=["#22c55e" if v==ef["eficiencia"].max() else "#3b82f6" for v in ef["eficiencia"]]))
            fig3.update_layout(title="Questões/hora por tipo")
            st.plotly_chart(fig_layout(fig3, 300), use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio("Sem dados de questões."), use_container_width=True)


def _aba_ranking(analisavel):
    if analisavel.empty:
        st.info("Sem dados para ranking."); return
    ag = analisavel.groupby("aluno", as_index=False).agg(
        horas=("ch_efetiva","sum"), concluidas=("status", lambda s:(s==STATUS_CONCLUIDA).sum()),
        questoes=("qtd_questoes_feitas","sum"), acertos=("qtd_acertos","sum"),
        dias=("data_ref", lambda s: s.dropna().dt.date.nunique()),
    )
    ag["desempenho"]   = ag.apply(lambda r: r["acertos"]/r["questoes"]*100 if r["questoes"] else 0, axis=1)
    ag["produtividade"] = ag.apply(lambda r: r["concluidas"]/r["horas"] if r["horas"] else 0, axis=1)

    def _rank_rows(df_s, col, fmt, cor):
        mx = df_s[col].max() or 1; html = ""
        medals = {0:"🥇",1:"🥈",2:"🥉"}
        for i,(_, r) in enumerate(df_s.iterrows()):
            pos = medals.get(i, str(i+1))
            pct = r[col]/mx*100
            val = fmt.format(r[col])
            html += (f'<div class="rank-row"><div class="rank-pos">{pos}</div>'
                f'<div class="rank-name">{escape_html(r["aluno"])}</div>'
                f'<div class="rank-bar-wrap"><div class="rank-bar" style="width:{pct:.0f}%;background:{cor}"></div></div>'
                f'<div class="rank-val">{escape_html(val)}</div></div>')
        return html

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        render_html('<div class="section-title">🏆 Produtividade (t/h)</div>')
        render_html(_rank_rows(ag.sort_values("produtividade",ascending=False), "produtividade", "{:.2f}", "#3b82f6"))
    with col_b:
        render_html('<div class="section-title">🔥 Consistência (dias)</div>')
        render_html(_rank_rows(ag.sort_values("dias",ascending=False), "dias", "{:.0f}", "#22c55e"))
    with col_c:
        render_html('<div class="section-title">📊 Desempenho (%)</div>')
        render_html(_rank_rows(ag.sort_values("desempenho",ascending=False), "desempenho", "{:.1f}%", "#f59e0b"))
    st.markdown("---")
    st.dataframe(ag.sort_values("produtividade",ascending=False).rename(columns={
        "horas":"Horas","concluidas":"Concluídas","questoes":"Questões","acertos":"Acertos",
        "dias":"Dias ativos","desempenho":"Desempenho (%)","produtividade":"Produtividade (t/h)"}),
        use_container_width=True, hide_index=True)


def _aba_analise_ia(df_filtrado, visao, kpis):
    analisavel = kpis["analisavel"]
    if analisavel.empty:
        st.info("Sem atividades para análise."); return
    alunos_lista = sorted(analisavel["aluno"].dropna().unique().tolist()) if visao == "Todos" else [visao]
    for nome_aluno in alunos_lista:
        grupo = df_filtrado[df_filtrado["aluno"]==nome_aluno].copy() if nome_aluno != "Todos" else df_filtrado.copy()
        grupo["data_ref"] = pd.to_datetime(grupo.get("data_execucao", pd.Series(dtype=str)), errors="coerce")
        kpis_al = calcular_kpis_avancados(grupo)
        insights = gerar_insights(grupo, kpis_al, nome_aluno)
        with st.expander(f"🧠 {nome_aluno}", expanded=(len(alunos_lista)==1)):
            c1,c2,c3,c4 = st.columns(4)
            c1.metric("Horas esta semana", f"{kpis_al['horas_semana']:.1f}h", delta=f"{kpis_al['delta_horas_semana']:+.1f}h vs sem. ant.")
            c2.metric("Progresso geral", f"{kpis_al['pct_conclusao']:.1f}%")
            c3.metric("Desempenho", f"{kpis_al['desempenho']:.1f}%")
            c4.metric("Sequência", f"{kpis_al['sequencia']} dias",
                delta=f"{kpis_al['dias_sem_estudar']}d sem estudar" if kpis_al["dias_sem_estudar"]>0 else "Estudou hoje")
            render_html('<div class="section-title">Análise automática</div>')
            for ins in insights:
                render_html(insight_card(ins["tipo"],ins["icone"],ins["titulo"],ins["texto"]))
            ana_al = kpis_al["analisavel"]
            if not ana_al.empty and "data_ref" in ana_al.columns:
                evo = ana_al.dropna(subset=["data_ref"]).copy()
                if not evo.empty:
                    sem = evo.set_index("data_ref").resample("W").agg(horas=("ch_efetiva","sum"), concluidas=("status", lambda s:(s==STATUS_CONCLUIDA).sum())).reset_index()
                    sem["label"] = sem["data_ref"].dt.strftime("Sem %d/%m")
                    if len(sem) > 1:
                        fig = go.Figure()
                        fig.add_trace(go.Bar(x=sem["label"], y=sem["horas"], name="Horas", marker_color="#3b82f6"))
                        fig.add_trace(go.Scatter(x=sem["label"], y=sem["concluidas"], name="Concluídas", mode="lines+markers", yaxis="y2", marker_color="#22c55e", line=dict(width=2)))
                        fig.update_layout(title=f"Evolução semanal — {nome_aluno}", yaxis2=dict(overlaying="y", side="right"))
                        st.plotly_chart(fig_layout(fig, 240), use_container_width=True)
            if not ana_al.empty and "disciplina" in ana_al.columns:
                da = ana_al.groupby("disciplina", as_index=False).agg(
                    questoes=("qtd_questoes_feitas","sum"), acertos=("qtd_acertos","sum"),
                    tarefas=("tarefa_id","count"), concluidas=("status", lambda s:(s==STATUS_CONCLUIDA).sum()),
                )
                da["desempenho"] = da.apply(lambda r: r["acertos"]/r["questoes"]*100 if r["questoes"] else 0, axis=1)
                da["progresso"]  = da.apply(lambda r: r["concluidas"]/r["tarefas"]*100 if r["tarefas"] else 0, axis=1)
                frageis  = da[(da["questoes"]>0) & (da["desempenho"]<70)].sort_values("desempenho")
                criticas = da[da["progresso"]<30].sort_values("progresso")
                if not frageis.empty or not criticas.empty:
                    render_html('<div class="section-title">Disciplinas que precisam de atenção</div>')
                    fc1, fc2 = st.columns(2)
                    with fc1:
                        if not frageis.empty:
                            st.caption("🔴 Baixo desempenho em questões (<70%)")
                            st.dataframe(frageis[["disciplina","desempenho","questoes","acertos"]].rename(columns={"disciplina":"Disciplina","desempenho":"Desempenho (%)","questoes":"Questões","acertos":"Acertos"}), hide_index=True, use_container_width=True)
                    with fc2:
                        if not criticas.empty:
                            st.caption("⚠️ Baixo progresso (<30% concluído)")
                            st.dataframe(criticas[["disciplina","progresso","tarefas","concluidas"]].rename(columns={"disciplina":"Disciplina","progresso":"Progresso (%)","tarefas":"Tarefas","concluidas":"Concluídas"}), hide_index=True, use_container_width=True)


# ─────────────────────────────────────────────
# TELA: DASHBOARD
# ─────────────────────────────────────────────

def _tooltip_grafico(texto: str) -> str:
    """Retorna HTML de um ícone ? com tooltip flutuante para gráficos."""
    return (
        f'<div class="kpi-ttip-wrap" style="display:inline-block;margin-left:6px;vertical-align:middle">'
        f'<div class="kpi-ttip-icon">?</div>'
        f'<div class="kpi-ttip-box">{escape_html(texto)}</div>'
        f'</div>'
    )


def _titulo_secao(label: str, tooltip: str = "", periodo: str = "") -> None:
    badge = ""
    if periodo:
        cor   = "#eff6ff" if "7" in periodo else "#f0fdf4"
        borda = "#bfdbfe" if "7" in periodo else "#bbf7d0"
        txt   = "#1d4ed8" if "7" in periodo else "#166534"
        badge = (
            f'<span style="background:{cor};color:{txt};border:1px solid {borda};'
            f'border-radius:999px;padding:2px 10px;font-size:.67rem;font-weight:800;'
            f'margin-left:8px;vertical-align:middle">{periodo}</span>'
        )
    tip = _tooltip_grafico(tooltip) if tooltip else ""
    render_html(
        f'<div style="font-size:.88rem;font-weight:800;color:#0f172a;margin:18px 0 6px;'
        f'display:flex;align-items:center;gap:4px">'
        f'{escape_html(label)}{badge}{tip}</div>'
    )


def _resumo_7dias(df_total: pd.DataFrame) -> None:
    """Bloco de KPIs dos últimos 7 dias com labels e tooltips."""
    hoje   = date.today()
    ini7   = hoje - timedelta(days=6)
    df7    = df_total[df_total["data_ref"].dt.date >= ini7].copy()
    ana7   = df7[df7["status"].isin(STATUS_ANALISE)]

    h7     = float(ana7["ch_efetiva"].sum())
    q7     = int(ana7["qtd_questoes_feitas"].sum())
    ac7    = int(ana7["qtd_acertos"].sum())
    des7   = ac7 / q7 * 100 if q7 else 0
    conc7  = int((ana7["status"] == STATUS_CONCLUIDA).sum())
    dias7  = ana7["data_ref"].dropna().dt.date.nunique()

    _titulo_secao(
        "Últimos 7 dias",
        "Indicadores calculados apenas com execuções dos últimos 7 dias (hoje inclusive). "
        "Útil para monitorar o ritmo recente independente do histórico total.",
        "Últimos 7 dias",
    )
    cols = st.columns(5)
    cards = [
        kpi_card("Horas estudadas", f"{h7:.1f}h", f"{dias7} dia(s) com registro",
            tooltip="Total de horas de estudo registradas nos últimos 7 dias. Inclui atividades Em andamento e Concluídas."),
        kpi_card("Questões feitas", f"{q7}", "nos últimos 7 dias",
            tooltip="Soma de todas as questões resolvidas em atividades iniciadas ou concluídas nos últimos 7 dias."),
        kpi_card("Acertos", f"{ac7}", f"de {q7} questões",
            tooltip="Total de questões acertadas nos últimos 7 dias. Fórmula: soma de qtd_acertos das execuções do período."),
        kpi_card("Desempenho", f"{des7:.1f}%", "taxa de acerto (7d)",
            tooltip="Taxa de acerto nos últimos 7 dias. Fórmula: (acertos ÷ questões) × 100. Acima de 70% é satisfatório para concursos."),
        kpi_card("Tarefas concluídas", f"{conc7}", "nos últimos 7 dias",
            tooltip="Quantidade de tarefas marcadas como Concluída com data de execução nos últimos 7 dias."),
    ]
    for col, card in zip(cols, cards):
        with col:
            render_html(card)


def dashboard():
    render_html(f"""
        <div class="hero">
          <h1>📊 {APP_NAME}</h1>
          <p>Dashboard analítico: produtividade, avanço, desempenho, evolução e análise inteligente.</p>
        </div>
    """)
    df = carregar_execucoes()
    if df.empty:
        st.info("Cadastre alunos, tarefas e registros para iniciar o acompanhamento.")
        return

    df_filtrado, visao = painel_filtros(df, "dash")
    if df_filtrado.empty:
        st.info("Nenhum registro encontrado com os filtros selecionados.")
        return

    df_filtrado = df_filtrado.copy()
    df_filtrado["data_ref"] = pd.to_datetime(
        df_filtrado.get("data_execucao", pd.Series(dtype=str)), errors="coerce"
    )
    kpis       = calcular_kpis_avancados(df_filtrado)
    analisavel = kpis["analisavel"]

    st.caption(
        "📌 Passe o mouse sobre ? para ver fórmula e interpretação. "
        "Indicadores em azul = últimos 7 dias · Verde = histórico completo."
    )

    # ── Bloco 7 dias ──
    _resumo_7dias(df_filtrado)

    st.markdown("---")

    # ── Bloco histórico ──
    _titulo_secao(
        "Histórico completo",
        "Indicadores calculados sobre todo o período disponível nos filtros selecionados.",
        "Histórico completo",
    )
    _render_kpis_produtividade(kpis)

    _titulo_secao("Avanço no plano",
        "Métricas de progresso em relação ao total de tarefas do plano de estudos.", "Histórico completo")
    _render_kpis_avanco(kpis)

    _titulo_secao("Desempenho acadêmico",
        "Indicadores de performance nas questões e tarefas concluídas.", "Histórico completo")
    _render_kpis_desempenho(kpis)

    st.markdown("---")

    abas = st.tabs([
        "📊 Visão geral",
        "📚 Disciplinas",
        "📅 Evolução",
        "⏱️ Gestão do tempo",
        "🏆 Rankings",
        "🧠 Análise IA",
        "📋 Atividades",
    ])

    with abas[0]:
        _titulo_secao("Distribuição por status",
            "Quantas tarefas estão em cada status (Não iniciada, Em andamento, Concluída). "
            "Permite visualizar a fila de trabalho e o progresso geral.")
        _aba_visao_geral(df_filtrado, analisavel)

    with abas[1]:
        _titulo_secao("Análise por disciplina",
            "Progresso e desempenho separados por disciplina. "
            "Identifica quais áreas estão avançando bem e quais precisam de mais atenção.")
        _aba_disciplinas(analisavel)

    with abas[2]:
        _titulo_secao("Evolução temporal",
            "Gráficos de horas e conclusões ao longo do tempo. "
            "Mostra tendência de crescimento, constância e variações no ritmo de estudo.")
        _aba_evolucao(analisavel)

    with abas[3]:
        _titulo_secao("Gestão do tempo",
            "Como as horas de estudo estão distribuídas entre tipos de atividade e disciplinas. "
            "Ajuda a identificar desequilíbrios e priorizar melhor o tempo disponível.")
        _aba_gestao_tempo(analisavel)

    with abas[4]:
        _titulo_secao("Rankings comparativos",
            "Classificação dos alunos por produtividade, consistência e desempenho. "
            "Cada ranking usa uma métrica diferente para uma visão multidimensional.")
        _aba_ranking(analisavel)

    with abas[5]:
        _titulo_secao("Análise inteligente por aluno",
            "Insights automáticos gerados com base no histórico individual: padrões, riscos, disciplinas frágeis e recomendações.")
        _aba_analise_ia(df_filtrado, visao, kpis)

    with abas[6]:
        _titulo_secao("Todas as atividades",
            "Tabela detalhada com todos os registros de execução do período filtrado. "
            "Útil para auditoria e acompanhamento granular.")
        tabela  = preparar_tabela(df_filtrado)
        colunas = ["aluno","tarefa","disciplina","aula","assunto","tipo","status_label",
                   "data_execucao","ch_efetiva","qtd_questoes_feitas","qtd_acertos","desempenho","comentario"]
        st.dataframe(
            tabela[[c for c in colunas if c in tabela.columns]],
            use_container_width=True, hide_index=True, row_height=72,
        )


# ─────────────────────────────────────────────
# TELA: REGISTRO RÁPIDO
# ─────────────────────────────────────────────

def ultima_atividade(aluno_id=None):
    """Retorna o registro mais recente iniciado ou concluído do aluno."""
    filtro = ""
    params = []
    if aluno_id:
        filtro = "AND e.aluno_id = ?"
        params.append(int(aluno_id))
    df = consultar(
        f"""
        SELECT
            e.id AS execucao_id, al.nome AS aluno, t.numero AS tarefa, d.nome AS disciplina,
            COALESCE(ass.titulo, t.conteudo) AS assunto, COALESCE(t.tipo, 'Outro') AS tipo,
            e.status, e.data_execucao, e.ch_efetiva, e.comentario, e.atualizado_em,
            e.tarefa_id, e.aluno_id
        FROM execucoes e
        JOIN alunos al ON al.id = e.aluno_id AND al.ativo = 1
        JOIN tarefas t ON t.id = e.tarefa_id AND t.ativo = 1
        JOIN disciplinas d ON d.id = t.disciplina_id AND d.ativo = 1
        LEFT JOIN assuntos ass ON ass.id = t.assunto_id
        WHERE e.status IN ('EM_ANDAMENTO','CONCLUIDA') {filtro}
        ORDER BY e.atualizado_em DESC
        LIMIT 1
        """,
        tuple(params),
    )
    return None if df.empty else df.iloc[0].to_dict()


def exibir_card_ultima(atividade):
    if not atividade:
        st.info("Ainda não há atividade recente iniciada ou concluída.")
        return
    classe = status_card_class(atividade["status"])
    render_html(f"""
        <div class="quick-card {classe}">
          <div class="quick-label">Última atividade registrada</div>
          <h3>Tarefa {escape_html(atividade['tarefa'])} · {escape_html(atividade['disciplina'])}</h3>
          <div>{status_badge(atividade['status'])}</div>
          <div class="quick-grid">
            <div class="quick-item"><div class="quick-label">Aluno</div>
              <div class="quick-value">{escape_html(atividade['aluno'])}</div></div>
            <div class="quick-item"><div class="quick-label">Assunto</div>
              <div class="quick-value">{escape_html(atividade['assunto'] or '-')}</div></div>
            <div class="quick-item"><div class="quick-label">Tipo</div>
              <div class="quick-value">{escape_html(atividade['tipo'])}</div></div>
            <div class="quick-item"><div class="quick-label">Tempo</div>
              <div class="quick-value">{float(atividade['ch_efetiva'] or 0):.2f}h</div></div>
            <div class="quick-item"><div class="quick-label">Data</div>
              <div class="quick-value">{escape_html(formatar_data_br(atividade['data_execucao']))}</div></div>
            <div class="quick-item"><div class="quick-label">Atualização</div>
              <div class="quick-value">{escape_html(str(atividade['atualizado_em']))}</div></div>
            <div class="quick-item" style="grid-column: span 2;"><div class="quick-label">Observações</div>
              <div class="quick-value">{escape_html(atividade['comentario'] or '-')}</div></div>
          </div>
        </div>
    """)


def _label_tarefa(row) -> str:
    return f"Tarefa {int(row['tarefa'])} — {row['disciplina']} | {str(row['assunto'] or row['conteudo'] or '')[:60]}"


def _verificar_regra_andamento(tarefas_df: pd.DataFrame, tarefa_selecionada, novo_status: str, disciplina_id: int) -> tuple[bool, str]:
    """
    Regra: uma tarefa NAO_INICIADA só pode ser iniciada se não existir outra
    tarefa em andamento para a mesma disciplina.
    Retorna (pode_salvar, mensagem_de_erro).
    """
    status_atual = str(tarefa_selecionada.get("status", STATUS_NAO_INICIADA))
    if status_atual != STATUS_NAO_INICIADA:
        return True, ""
    if novo_status == STATUS_NAO_INICIADA:
        return True, ""

    # Há alguma outra tarefa em andamento na mesma disciplina?
    outras_em_andamento = tarefas_df[
        (tarefas_df["status"] == STATUS_EM_ANDAMENTO)
        & (tarefas_df["disciplina_id"] == disciplina_id)
        & (tarefas_df["tarefa_id"] != tarefa_selecionada["tarefa_id"])
    ]
    if outras_em_andamento.empty:
        return True, ""

    tarefa_bloqueante = outras_em_andamento.iloc[0]
    msg = (
        f"❌ Não é possível iniciar esta tarefa. "
        f"A Tarefa {int(tarefa_bloqueante['tarefa'])} ({tarefa_bloqueante['disciplina']}) "
        f"está em andamento. Conclua ou atualize essa tarefa antes de iniciar outra na mesma disciplina."
    )
    return False, msg


def _painel_tarefa(tarefa) -> None:
    """
    Exibe o conteúdo completo da tarefa selecionada antes do formulário:
    disciplina, aula, assunto, descrição e exercícios previstos.
    Texto nunca é cortado.
    """
    disciplina = escape_html(str(tarefa.get("disciplina") or ""))
    aula       = escape_html(str(tarefa.get("aula") or ""))
    assunto    = escape_html(str(tarefa.get("assunto") or ""))
    conteudo   = str(tarefa.get("conteudo") or "").strip()
    previstos  = int(tarefa.get("qtd_exercicios_previstos") or 0)
    tipo       = escape_html(str(tarefa.get("tipo") or ""))
    num_tarefa = int(tarefa.get("tarefa", 0))

    # Cabeçalho da tarefa
    render_html(f"""
    <div style="
        background:#f8fafc;
        border:1px solid #e2e8f0;
        border-left:4px solid #3b82f6;
        border-radius:10px;
        padding:16px 20px;
        margin-bottom:14px;
    ">
      <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:10px">
        <div>
          <div style="font-size:.68rem;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:.07em">
            Tarefa {num_tarefa}
          </div>
          <div style="font-size:1.05rem;font-weight:800;color:#0f172a;margin-top:2px">{disciplina}</div>
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <span style="background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe;
            border-radius:999px;padding:3px 10px;font-size:.73rem;font-weight:700">{tipo}</span>
          {f'<span style="background:#f0fdf4;color:#166534;border:1px solid #bbf7d0;border-radius:999px;padding:3px 10px;font-size:.73rem;font-weight:700">📝 {previstos} exercícios previstos</span>' if previstos else ''}
        </div>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:{'12px' if conteudo else '0'}">
        <div>
          <div style="font-size:.66rem;font-weight:800;color:#94a3b8;text-transform:uppercase;margin-bottom:3px">Aula</div>
          <div style="font-size:.87rem;color:#0f172a;font-weight:500;line-height:1.45">{aula or '—'}</div>
        </div>
        <div>
          <div style="font-size:.66rem;font-weight:800;color:#94a3b8;text-transform:uppercase;margin-bottom:3px">Assunto</div>
          <div style="font-size:.87rem;color:#0f172a;font-weight:500;line-height:1.45">{assunto or '—'}</div>
        </div>
      </div>
      {f'''<div style="border-top:1px solid #e2e8f0;padding-top:10px;margin-top:4px">
        <div style="font-size:.66rem;font-weight:800;color:#94a3b8;text-transform:uppercase;margin-bottom:6px">
          Descrição / Objetivos do bloco
        </div>
        <div style="font-size:.84rem;color:#334155;line-height:1.65;white-space:pre-wrap">{escape_html(conteudo)}</div>
      </div>''' if conteudo else ''}
    </div>
    """)


def _toast_sucesso(msg: str) -> None:
    """
    Exibe uma mensagem de sucesso destacada com auto-fechamento visual.
    Usa st.toast quando disponível (Streamlit ≥ 1.28), senão st.success.
    """
    try:
        st.toast(msg, icon="✅")
    except AttributeError:
        st.success(msg)


def _renderizar_secao_registro(
    tarefas_df: pd.DataFrame,
    aluno_id: int,
    status_alvo: str,
    titulo: str,
    key_prefix: str,
) -> None:
    """
    Seção de registro por status.
    Ao selecionar uma tarefa:
      • exibe o painel completo com nome, aula, assunto e descrição (sem cortes);
      • formula pré-preenchida com valores do banco;
      • valida e grava;
      • exibe toast de sucesso ou erros em destaque.
    """
    grupo = tarefas_df[tarefas_df["status"] == status_alvo].copy()
    if grupo.empty:
        render_html(
            f'<div style="text-align:center;padding:40px 20px;color:var(--c-muted)">'
            f'<div style="font-size:2rem;margin-bottom:8px">✅</div>'
            f'<div style="font-size:.9rem">Nenhuma tarefa '
            f'<strong>{STATUS_LABELS[status_alvo].lower()}</strong> no momento.</div>'
            f'</div>'
        )
        return

    # ── Seletor de tarefa ──
    tarefa_id = st.selectbox(
        "Selecione a tarefa",
        grupo["tarefa_id"].tolist(),
        format_func=lambda v: _label_tarefa(grupo[grupo["tarefa_id"] == v].iloc[0]),
        key=f"{key_prefix}_sel",
    )
    tarefa = grupo[grupo["tarefa_id"] == tarefa_id].iloc[0]

    # ── Painel completo da tarefa selecionada ──
    _painel_tarefa(tarefa)

    # ── Aviso para concluídas ──
    eh_concluida = (status_alvo == STATUS_CONCLUIDA)
    if eh_concluida:
        render_html(
            '<div class="rule-warning" style="margin-bottom:14px">'
            '⚠️ <strong>Atenção:</strong> Esta tarefa já está <strong>Concluída</strong>. '
            'Qualquer alteração modifica dados históricos. '
            'Você precisará confirmar antes de salvar.'
            '</div>'
        )

    # ── Preparar valores do banco ──
    status_atual = str(tarefa.get("status", STATUS_NAO_INICIADA))
    tipo_atual   = str(tarefa.get("tipo") or "Outro")
    tipo_idx     = TIPOS_ESTUDO.index(tipo_atual) if tipo_atual in TIPOS_ESTUDO else TIPOS_ESTUDO.index("Outro")

    data_atual = date.today()
    raw_data   = tarefa.get("data_execucao")
    if raw_data is not None:
        try:
            parsed = pd.to_datetime(raw_data)
            if not pd.isnull(parsed):
                data_atual = parsed.date()
        except Exception:
            pass

    def _sf(v, p=0.0):
        try:
            f = float(v); return p if f != f else f
        except (TypeError, ValueError):
            return p

    def _si(v, p=0):
        try:
            return int(round(float(v)))
        except (TypeError, ValueError):
            return p

    ch_atual        = _sf(tarefa.get("ch_efetiva"), 0.0)
    questoes_atual  = _si(tarefa.get("qtd_questoes_feitas"), 0)
    acertos_atual   = _si(tarefa.get("qtd_acertos"), 0)
    comentario_atual = str(tarefa.get("comentario") or "")

    # ── Formulário pré-preenchido ──
    render_html('<div style="font-size:.8rem;font-weight:700;color:#475569;margin-bottom:6px">✏️ Editar registro</div>')

    with st.form(f"form_{key_prefix}_{tarefa_id}"):
        col1, col2, col3, col4 = st.columns(4)
        novo_status = col1.selectbox(
            "Status",
            STATUS_VALIDOS,
            index=STATUS_VALIDOS.index(status_atual),
            format_func=lambda v: STATUS_LABELS[v],
        )
        tipo_estudo = col2.selectbox(
            "Tipo de estudo",
            TIPOS_ESTUDO,
            index=tipo_idx,
        )
        data_estudo = col3.date_input(
            "Data do estudo",
            value=data_atual,
        )
        ch = col4.number_input(
            "Tempo gasto (h)",
            min_value=0.0,
            value=ch_atual,
            step=0.25,
            format="%.2f",
        )

        col5, col6 = st.columns(2)
        questoes = col5.number_input(
            "Questões feitas",
            min_value=0,
            value=questoes_atual,
            step=1,
        )
        acertos = col6.number_input(
            "Acertos",
            min_value=0,
            value=acertos_atual,
            step=1,
        )

        comentario = st.text_area(
            "Observações",
            value=comentario_atual,
            placeholder="Anotações sobre esta sessão de estudo (opcional)…",
            height=100,
        )

        if eh_concluida:
            confirmado = st.checkbox(
                "✅ Confirmo que desejo alterar este registro já concluído (dados históricos serão modificados)",
                value=False,
            )
        else:
            confirmado = True

        col_b1, col_b2 = st.columns([3, 1])
        salvar   = col_b1.form_submit_button("💾 Salvar", use_container_width=True, type="primary")
        cancelar = col_b2.form_submit_button("✖ Cancelar", use_container_width=True)

    # ── Pós-form ──
    if cancelar:
        st.info("Nenhuma alteração foi salva.")
        return

    if not salvar:
        return

    # ── Validações ──
    erros = []
    if acertos > questoes:
        erros.append("O número de acertos não pode ser maior que o número de questões feitas.")
    if eh_concluida and not confirmado:
        erros.append("Para alterar uma tarefa já concluída, marque a confirmação acima.")
    pode, msg_bloqueio = _verificar_regra_andamento(
        tarefas_df, tarefa, novo_status, int(tarefa["disciplina_id"])
    )
    if not pode:
        erros.append(msg_bloqueio)

    if erros:
        for e in erros:
            render_html(f'<div class="rule-error">❌ {escape_html(e)}</div>')
        return

    # ── Gravação ──
    try:
        with conectar() as conn:
            upsert_execucao(
                conn, aluno_id, int(tarefa_id), str(data_estudo),
                ch, None, 0, acertos, limpar_texto(comentario),
                questoes, novo_status, tipo_estudo,
            )
        limpar_cache()

        status_ant = STATUS_LABELS.get(status_atual, status_atual)
        status_nov = STATUS_LABELS.get(novo_status, novo_status)
        mudou      = status_atual != novo_status
        partes     = [f"Tarefa {int(tarefa['tarefa'])} — {tarefa['disciplina']} atualizada com sucesso."]
        if mudou:
            partes.append(f"Status: {status_ant} → {status_nov}.")
        if ch > 0:
            partes.append(f"{ch:.2f}h registradas.")
        if questoes > 0:
            partes.append(f"{acertos}/{questoes} acertos.")

        _toast_sucesso(" ".join(partes))
        st.rerun()

    except Exception as exc:
        erro_usuario("❌ Não foi possível salvar. Verifique os dados e tente novamente.", exc)


def tela_registro_rapido():
    render_html(
        '<div class="hero">'
        '<h1>⚡ Registro Rápido</h1>'
        '<p>Selecione uma tarefa para ver seu conteúdo completo e registrar o progresso. '
        'O formulário é pré-preenchido automaticamente com os dados do último registro.</p>'
        '</div>'
    )

    usuario = aluno_logado()
    alunos  = alunos_ativos()

    # ── Seleção do aluno ──
    if usuario["perfil"] == "Aluno":
        aluno_id = int(usuario["id"])
        render_html(
            f'<div style="display:inline-flex;align-items:center;gap:8px;'
            f'background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;'
            f'padding:8px 14px;margin-bottom:14px">'
            f'<span style="font-size:1.1rem">👤</span>'
            f'<div><div style="font-size:.66rem;font-weight:800;color:#94a3b8;text-transform:uppercase">Aluno</div>'
            f'<div style="font-size:.9rem;font-weight:700;color:#0f172a">{escape_html(usuario["nome"])}</div></div>'
            f'</div>'
        )
    else:
        if alunos.empty:
            st.info("Nenhum aluno cadastrado. Vá em **Alunos** para cadastrar.")
            return
        aluno_id = st.selectbox(
            "Aluno",
            alunos["id"].tolist(),
            format_func=lambda v: alunos.loc[alunos["id"] == v, "nome"].iloc[0],
        )

    # ── Última atividade ──
    exibir_card_ultima(ultima_atividade(aluno_id))

    # ── Carrega tarefas ──
    tarefas = carregar_visao_tarefas(aluno_id)
    if tarefas.empty:
        st.info("Nenhuma tarefa vinculada a este aluno. Vá em **Tarefas** para vincular.")
        return

    # ── Contadores ──
    qtd_nao  = len(tarefas[tarefas["status"] == STATUS_NAO_INICIADA])
    qtd_and  = len(tarefas[tarefas["status"] == STATUS_EM_ANDAMENTO])
    qtd_conc = len(tarefas[tarefas["status"] == STATUS_CONCLUIDA])

    # ── Abas por status ──
    aba_nao, aba_and, aba_conc = st.tabs([
        f"🔘 Não iniciadas ({qtd_nao})",
        f"🟡 Em andamento ({qtd_and})",
        f"🟢 Concluídas ({qtd_conc})",
    ])

    with aba_nao:
        render_html(
            '<div class="insight-card info" style="margin-bottom:14px">'
            '<div class="insight-icon">ℹ️</div>'
            '<div class="insight-body">'
            '<div class="insight-title">Regra: tarefas não iniciadas</div>'
            '<p class="insight-text">Uma tarefa não iniciada só pode ser iniciada se '
            '<strong>não houver outra tarefa em andamento na mesma disciplina</strong>. '
            'Conclua ou atualize a tarefa em andamento primeiro.</p>'
            '</div></div>'
        )
        _renderizar_secao_registro(tarefas, aluno_id, STATUS_NAO_INICIADA, "Não iniciadas", "rr_nao")

    with aba_and:
        _renderizar_secao_registro(tarefas, aluno_id, STATUS_EM_ANDAMENTO, "Em andamento", "rr_and")

    with aba_conc:
        render_html(
            '<div class="insight-card warning" style="margin-bottom:14px">'
            '<div class="insight-icon">⚠️</div>'
            '<div class="insight-body">'
            '<div class="insight-title">Atenção: tarefas concluídas</div>'
            '<p class="insight-text">Alterações em tarefas concluídas modificam o histórico de estudos. '
            'O sistema exigirá confirmação explícita antes de salvar.</p>'
            '</div></div>'
        )
        _renderizar_secao_registro(tarefas, aluno_id, STATUS_CONCLUIDA, "Concluídas", "rr_conc")

    # ── Histórico recente ──
    recentes = carregar_execucoes()
    recentes = recentes[recentes["aluno_id"] == aluno_id].head(10)
    if not recentes.empty:
        st.markdown("---")
        render_html('<div class="section-title">📋 Histórico recente</div>')
        tabela = preparar_tabela(recentes)
        colunas = ["tarefa", "disciplina", "assunto", "tipo", "status_label",
                   "data_execucao", "ch_efetiva", "qtd_questoes_feitas", "qtd_acertos",
                   "desempenho", "comentario"]
        st.dataframe(
            tabela[[c for c in colunas if c in tabela.columns]].rename(columns={
                "tarefa": "Tarefa", "disciplina": "Disciplina", "assunto": "Assunto",
                "tipo": "Tipo", "status_label": "Status", "data_execucao": "Data",
                "ch_efetiva": "Horas", "qtd_questoes_feitas": "Questões",
                "qtd_acertos": "Acertos", "desempenho": "Desempenho (%)",
                "comentario": "Observações",
            }),
            use_container_width=True,
            hide_index=True,
            row_height=64,
        )


# ─────────────────────────────────────────────
# TELA: TAREFAS (CRUD)
# ─────────────────────────────────────────────

def tela_tarefas():
    st.header("Gestão de tarefas educacionais")
    usuario    = aluno_logado()
    alunos     = alunos_ativos()
    disciplinas = disciplinas_ativas()
    aulas      = carregar_aulas()
    assuntos   = carregar_assuntos()
    tarefas    = carregar_tarefas_base()

    abas = st.tabs(["Execução e status", "Criar tarefa", "Editar tarefas", "Vincular alunos"])

    # ── Aba 1: Execução e status ──
    with abas[0]:
        if usuario["perfil"] == "Aluno":
            aluno_id = int(usuario["id"])
            st.text_input("Aluno", value=usuario["nome"], disabled=True)
        else:
            if alunos.empty:
                st.info("Cadastre um aluno.")
                return
            aluno_id = st.selectbox(
                "Aluno", alunos["id"].tolist(),
                format_func=lambda v: alunos.loc[alunos["id"] == v, "nome"].iloc[0],
            )
        df = carregar_visao_tarefas(aluno_id)
        col1, col2, col3 = st.columns(3)
        disciplina_f = col1.multiselect("Disciplina", sorted(df["disciplina"].dropna().unique().tolist()))
        tipo_f       = col2.multiselect("Tipo de estudo", sorted(df["tipo"].dropna().unique().tolist()))
        status_f     = col3.multiselect("Status", STATUS_VALIDOS, default=STATUS_VALIDOS, format_func=lambda v: STATUS_LABELS[v])
        filtrado = df.copy()
        if disciplina_f: filtrado = filtrado[filtrado["disciplina"].isin(disciplina_f)]
        if tipo_f:       filtrado = filtrado[filtrado["tipo"].isin(tipo_f)]
        if status_f:     filtrado = filtrado[filtrado["status"].isin(status_f)]

        colunas_editor = ["tarefa_id", "tarefa", "disciplina", "aula", "assunto", "tipo", "status",
                          "data_execucao", "ch_efetiva", "qtd_questoes_feitas", "qtd_acertos", "desempenho", "comentario"]
        editor = filtrado[[c for c in colunas_editor if c in filtrado.columns]].copy()
        editor["data_execucao"] = editor["data_execucao"].fillna("")
        editado = st.data_editor(
            editor,
            use_container_width=True,
            hide_index=True,
            row_height=72,
            disabled=["tarefa_id", "tarefa", "disciplina", "aula", "assunto", "desempenho"],
            column_config={
                "tarefa_id": None,
                "tarefa": "Tarefa",
                "tipo": st.column_config.SelectboxColumn("Tipo de estudo", options=TIPOS_ESTUDO),
                "status": st.column_config.SelectboxColumn("Status", options=STATUS_VALIDOS, required=True),
                "data_execucao": st.column_config.TextColumn("Data do estudo (aaaa-mm-dd)"),
                "ch_efetiva": st.column_config.NumberColumn("Tempo gasto", min_value=0.0, step=0.25),
                "qtd_questoes_feitas": st.column_config.NumberColumn("Questões feitas", min_value=0, step=1),
                "qtd_acertos": st.column_config.NumberColumn("Acertos", min_value=0, step=1),
                "comentario": st.column_config.TextColumn("Observações", width="large"),
            },
            key="editor_execucoes",
        )
        if st.button("Salvar alterações de execução", type="primary"):
            try:
                with conectar() as conn:
                    for _, row in editado.iterrows():
                        if converter_inteiro(row["qtd_acertos"]) > converter_inteiro(row["qtd_questoes_feitas"]):
                            raise ValueError(f"Tarefa {row['tarefa']}: acertos maiores que questões.")
                        upsert_execucao(
                            conn, aluno_id, int(row["tarefa_id"]),
                            converter_data(row["data_execucao"]),
                            converter_horas(row["ch_efetiva"]), None, 0,
                            converter_inteiro(row["qtd_acertos"]),
                            limpar_texto(row["comentario"]),
                            converter_inteiro(row["qtd_questoes_feitas"]),
                            row["status"], row["tipo"],
                        )
                limpar_cache()
                _toast_sucesso("Registros de execução salvos com sucesso.")
                st.rerun()
            except Exception as exc:
                erro_usuario("Não foi possível salvar.", exc)

    # ── Aba 2: Criar tarefa ──
    with abas[1]:
        if usuario["perfil"] != "Gestor":
            st.info("Somente gestores podem criar tarefas.")
        elif disciplinas.empty:
            st.info("Cadastre disciplinas/aulas primeiro.")
        else:
            with st.form("nova_tarefa", clear_on_submit=True):
                col1, col2, col3 = st.columns(3)
                numero       = col1.number_input("Número da tarefa", min_value=1, step=1)
                trilha       = col2.number_input("Trilha", min_value=0, step=1)
                disciplina_id = col3.selectbox("Disciplina", disciplinas["id"].tolist(), format_func=lambda v: disciplinas.loc[disciplinas["id"] == v, "nome"].iloc[0])
                aulas_disc   = aulas[aulas["disciplina_id"] == disciplina_id]
                aula_id      = st.selectbox("Aula", aulas_disc["id"].tolist(), format_func=lambda v: aulas_disc.loc[aulas_disc["id"] == v, "aula"].iloc[0]) if not aulas_disc.empty else None
                assuntos_aula = assuntos[assuntos["aula_id"] == aula_id] if aula_id else assuntos.iloc[0:0]
                assunto_id   = st.selectbox("Assunto", assuntos_aula["id"].tolist(), format_func=lambda v: assuntos_aula.loc[assuntos_aula["id"] == v, "assunto"].iloc[0]) if not assuntos_aula.empty else None
                col4, col5   = st.columns(2)
                tipo         = col4.selectbox("Tipo de estudo", TIPOS_ESTUDO)
                previstos    = col5.number_input("Exercícios previstos", min_value=0, step=1)
                conteudo     = st.text_area("Conteúdo")
                alunos_vinc  = st.multiselect("Vincular alunos", alunos["id"].tolist(), format_func=lambda v: alunos.loc[alunos["id"] == v, "nome"].iloc[0])
                salvar       = st.form_submit_button("Criar tarefa", use_container_width=True)
            if salvar:
                try:
                    aula_nome = aulas.loc[aulas["id"] == aula_id, "aula"].iloc[0] if aula_id else "Aula não informada"
                    with conectar() as conn:
                        conn.execute(
                            "INSERT INTO tarefas (numero, trilha, disciplina_id, aula_id, assunto_id, aula, qtd_exercicios_previstos, tipo, conteudo, ativo) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1)",
                            (int(numero), int(trilha), int(disciplina_id), aula_id, assunto_id, aula_nome, int(previstos), tipo, limpar_texto(conteudo)),
                        )
                        tarefa_id = ultimo_id(conn)
                        for av in alunos_vinc:
                            upsert_execucao(conn, int(av), tarefa_id, None, 0, None, 0, 0, None, 0, STATUS_NAO_INICIADA)
                    limpar_cache()
                    _toast_sucesso("Tarefa criada com sucesso.")
                    st.rerun()
                except (IntegrityError, UniqueViolation):
                    st.error("Já existe uma tarefa com esse número.")

    # ── Aba 3: Editar tarefas ──
    with abas[2]:
        if usuario["perfil"] != "Gestor":
            st.info("Somente gestores podem editar tarefas.")
        elif tarefas.empty:
            st.info("Nenhuma tarefa cadastrada.")
        else:
            colunas_edit = ["tarefa_id", "tarefa", "trilha", "disciplina", "aula", "assunto", "tipo", "qtd_exercicios_previstos", "conteudo"]
            edit    = tarefas[[c for c in colunas_edit if c in tarefas.columns]].copy()
            editado = st.data_editor(
                edit, use_container_width=True, hide_index=True, row_height=72,
                disabled=["tarefa_id", "disciplina", "aula", "assunto"],
                column_config={
                    "tarefa_id": None,
                    "tipo": st.column_config.SelectboxColumn("Tipo", options=TIPOS_ESTUDO),
                    "conteudo": st.column_config.TextColumn("Conteúdo", width="large"),
                    "qtd_exercicios_previstos": st.column_config.NumberColumn("Exercícios previstos", min_value=0, step=1),
                },
                key="editor_tarefas_crud",
            )
            col1, col2 = st.columns(2)
            if col1.button("Salvar tarefas", type="primary"):
                with conectar() as conn:
                    for _, row in editado.iterrows():
                        conn.execute(
                            "UPDATE tarefas SET numero = ?, trilha = ?, tipo = ?, qtd_exercicios_previstos = ?, conteudo = ? WHERE id = ?",
                            (int(row["tarefa"]), converter_inteiro(row["trilha"]), row["tipo"], converter_inteiro(row["qtd_exercicios_previstos"]), limpar_texto(row["conteudo"]), int(row["tarefa_id"])),
                        )
                limpar_cache()
                _toast_sucesso("Alterações nas tarefas salvas com sucesso.")
                st.rerun()
            excluir = col2.selectbox("Excluir tarefa", tarefas["tarefa_id"].tolist(), format_func=lambda v: f"Tarefa {int(tarefas.loc[tarefas['tarefa_id'] == v, 'tarefa'].iloc[0])}")
            if col2.button("Excluir tarefa selecionada"):
                executar("UPDATE tarefas SET ativo = 0 WHERE id = ?", (int(excluir),))
                limpar_cache()
                _toast_sucesso("Tarefa excluída com sucesso.")
                st.rerun()

    # ── Aba 4: Vincular alunos ──
    with abas[3]:
        if usuario["perfil"] != "Gestor":
            st.info("Somente gestores podem vincular alunos.")
        elif tarefas.empty or alunos.empty:
            st.info("Cadastre alunos e tarefas.")
        else:
            tarefa_id = st.selectbox(
                "Tarefa", tarefas["tarefa_id"].tolist(),
                format_func=lambda v: f"Tarefa {int(tarefas.loc[tarefas['tarefa_id'] == v, 'tarefa'].iloc[0])} - {tarefas.loc[tarefas['tarefa_id'] == v, 'disciplina'].iloc[0]}",
            )
            vincular = st.multiselect("Alunos", alunos["id"].tolist(), format_func=lambda v: alunos.loc[alunos["id"] == v, "nome"].iloc[0])
            if st.button("Vincular alunos à tarefa", type="primary"):
                with conectar() as conn:
                    for av in vincular:
                        upsert_execucao(conn, int(av), int(tarefa_id), None, 0, None, 0, 0, None, 0, STATUS_NAO_INICIADA)
                limpar_cache()
                _toast_sucesso("Alunos vinculados à tarefa com sucesso.")
                st.rerun()


# ─────────────────────────────────────────────
# TELA: AULAS E ASSUNTOS
# ─────────────────────────────────────────────

def tela_aulas():
    st.header("Aulas e assuntos")
    usuario     = aluno_logado()
    disciplinas = disciplinas_ativas()
    aulas       = carregar_aulas()
    assuntos    = carregar_assuntos()
    abas        = st.tabs(["Aulas", "Assuntos"])

    with abas[0]:
        if usuario["perfil"] == "Gestor" and not disciplinas.empty:
            with st.form("nova_aula", clear_on_submit=True):
                col1, col2, col3 = st.columns(3)
                disciplina_id = col1.selectbox("Disciplina", disciplinas["id"].tolist(), format_func=lambda v: disciplinas.loc[disciplinas["id"] == v, "nome"].iloc[0])
                aula          = col2.text_input("Nome da aula")
                tipo_estudo   = col3.selectbox("Tipo de estudo", TIPOS_ESTUDO)
                salvar        = st.form_submit_button("Criar aula")
            if salvar:
                if not limpar_texto(aula):
                    st.error("Informe o nome da aula.")
                else:
                    with conectar() as conn:
                        upsert_aula(conn, disciplina_id, aula, tipo_estudo=tipo_estudo)
                    limpar_cache()
                    _toast_sucesso("Aula criada com sucesso.")
                    st.rerun()

        if aulas.empty:
            st.info("Nenhuma aula cadastrada.")
        else:
            editado = st.data_editor(
                aulas, use_container_width=True, hide_index=True,
                disabled=["id", "disciplina_id", "disciplina"],
                column_config={"tipo_estudo": st.column_config.SelectboxColumn("Tipo de estudo", options=TIPOS_ESTUDO)},
                key="editor_aulas",
            )
            if usuario["perfil"] == "Gestor":
                col1, col2 = st.columns(2)
                if col1.button("Salvar alterações de aulas", type="primary"):
                    with conectar() as conn:
                        for _, row in editado.iterrows():
                            conn.execute(
                                "UPDATE aulas SET aula = ?, tipo_estudo = ?, estudada_padrao = ?, revisao_24h_padrao = ? WHERE id = ?",
                                (limpar_texto(row["aula"]), normalizar_tipo_estudo(row["tipo_estudo"]), limpar_texto(row["estudada_padrao"]), limpar_texto(row["revisao_24h_padrao"]), int(row["id"])),
                            )
                    limpar_cache()
                    _toast_sucesso("Alterações nas aulas salvas com sucesso.")
                    st.rerun()
                excluir = col2.selectbox("Excluir aula", aulas["id"].tolist(), format_func=lambda v: aulas.loc[aulas["id"] == v, "aula"].iloc[0])
                if col2.button("Excluir aula selecionada"):
                    executar("UPDATE aulas SET ativo = 0 WHERE id = ?", (int(excluir),))
                    limpar_cache()
                    _toast_sucesso("Aula excluída com sucesso.")
                    st.rerun()

    with abas[1]:
        if usuario["perfil"] == "Gestor" and not aulas.empty:
            with st.form("novo_assunto", clear_on_submit=True):
                aula_id = st.selectbox(
                    "Aula", aulas["id"].tolist(),
                    format_func=lambda v: f"{aulas.loc[aulas['id'] == v, 'disciplina'].iloc[0]} — {aulas.loc[aulas['id'] == v, 'aula'].iloc[0]}",
                )
                titulo = st.text_input("Assunto")
                salvar = st.form_submit_button("Criar assunto")
            if salvar:
                if not limpar_texto(titulo):
                    st.error("Informe o assunto.")
                else:
                    with conectar() as conn:
                        upsert_assunto(conn, aula_id, titulo)
                    limpar_cache()
                    _toast_sucesso("Assunto criado com sucesso.")
                    st.rerun()

        if assuntos.empty:
            st.info("Nenhum assunto cadastrado.")
        else:
            editado = st.data_editor(
                assuntos, use_container_width=True, hide_index=True,
                disabled=["id", "aula_id", "disciplina", "aula"],
                key="editor_assuntos",
            )
            if usuario["perfil"] == "Gestor":
                col1, col2 = st.columns(2)
                if col1.button("Salvar alterações de assuntos", type="primary"):
                    with conectar() as conn:
                        for _, row in editado.iterrows():
                            conn.execute("UPDATE assuntos SET titulo = ? WHERE id = ?", (limpar_texto(row["assunto"]), int(row["id"])))
                    limpar_cache()
                    _toast_sucesso("Alterações nos assuntos salvas com sucesso.")
                    st.rerun()
                excluir = col2.selectbox("Excluir assunto", assuntos["id"].tolist(), format_func=lambda v: assuntos.loc[assuntos["id"] == v, "assunto"].iloc[0])
                if col2.button("Excluir assunto selecionado"):
                    executar("UPDATE assuntos SET ativo = 0 WHERE id = ?", (int(excluir),))
                    limpar_cache()
                    _toast_sucesso("Assunto excluído com sucesso.")
                    st.rerun()


# ─────────────────────────────────────────────
# TELA: DISCIPLINAS
# ─────────────────────────────────────────────

def tela_disciplinas():
    st.header("Disciplinas")
    usuario = aluno_logado()
    df      = disciplinas_ativas()

    if usuario["perfil"] == "Gestor":
        with st.form("nova_disciplina", clear_on_submit=True):
            nome   = st.text_input("Nome da disciplina")
            salvar = st.form_submit_button("Criar disciplina")
        if salvar:
            try:
                executar("INSERT INTO disciplinas (nome, ativo) VALUES (?, 1)", (limpar_texto(nome),))
                limpar_cache()
                _toast_sucesso("Disciplina criada com sucesso.")
                st.rerun()
            except (IntegrityError, UniqueViolation):
                st.error("Já existe uma disciplina com esse nome.")

        if not df.empty:
            editado = st.data_editor(df, use_container_width=True, hide_index=True, disabled=["id"], key="editor_disciplinas")
            col1, col2 = st.columns(2)
            if col1.button("Salvar disciplinas", type="primary"):
                with conectar() as conn:
                    for _, row in editado.iterrows():
                        conn.execute("UPDATE disciplinas SET nome = ? WHERE id = ?", (limpar_texto(row["nome"]), int(row["id"])))
                limpar_cache()
                _toast_sucesso("Alterações nas disciplinas salvas com sucesso.")
                st.rerun()
            excluir = col2.selectbox("Excluir disciplina", df["id"].tolist(), format_func=lambda v: df.loc[df["id"] == v, "nome"].iloc[0])
            if col2.button("Excluir disciplina selecionada"):
                executar("UPDATE disciplinas SET ativo = 0 WHERE id = ?", (int(excluir),))
                limpar_cache()
                _toast_sucesso("Disciplina excluída com sucesso.")
                st.rerun()

    st.dataframe(df, use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────
# TELA: ALUNOS
# ─────────────────────────────────────────────

def _excluir_aluno_cascade(aluno_id: int) -> str:
    """
    Remove permanentemente um aluno e todos os dados relacionados:
    execuções, vínculos e o próprio registro na tabela alunos.
    Retorna o nome do aluno excluído.
    """
    with conectar() as conn:
        row = conn.execute("SELECT nome FROM alunos WHERE id = ?", (int(aluno_id),)).fetchone()
        nome = row[0] if row else str(aluno_id)
        conn.execute("DELETE FROM execucoes WHERE aluno_id = ?", (int(aluno_id),))
        conn.execute("DELETE FROM alunos WHERE id = ? AND perfil = 'Aluno'", (int(aluno_id),))
    limpar_cache()
    return nome


def tela_alunos():
    st.header("Alunos")
    usuario = aluno_logado()
    if usuario["perfil"] != "Gestor":
        st.info("Somente gestores podem administrar alunos.")
        return

    abas = st.tabs(["👥 Alunos ativos", "➕ Novo aluno", "🧹 Limpeza de usuários"])

    # ── Aba 1: Alunos ativos ──
    with abas[0]:
        alunos = alunos_ativos(incluir_gestor=False)
        if alunos.empty:
            st.info("Nenhum aluno cadastrado.")
        else:
            editado = st.data_editor(
                alunos, use_container_width=True, hide_index=True,
                disabled=["id", "perfil"], key="editor_alunos",
            )
            col1, col2 = st.columns([2, 1])
            if col1.button("💾 Salvar alterações", type="primary"):
                try:
                    with conectar() as conn:
                        for _, row in editado.iterrows():
                            conn.execute(
                                "UPDATE alunos SET nome = ?, email = ? WHERE id = ? AND perfil = 'Aluno'",
                                (limpar_texto(row["nome"]), normalizar_email(row["email"]), int(row["id"])),
                            )
                    limpar_cache()
                    _toast_sucesso("Dados dos alunos atualizados com sucesso.")
                    st.rerun()
                except (IntegrityError, UniqueViolation):
                    st.error("Nome ou e-mail duplicado.")

            st.markdown("---")
            render_html('<div class="section-title">🗑️ Excluir aluno</div>')
            render_html(
                '<div class="rule-warning">'
                '⚠️ A exclusão remove <strong>permanentemente</strong> o aluno e todo o seu histórico '
                '(execuções, status, horas, acertos). Esta ação <strong>não pode ser desfeita</strong>.'
                '</div>'
            )
            excluir_id = st.selectbox(
                "Selecione o aluno a excluir",
                alunos["id"].tolist(),
                format_func=lambda v: f"{alunos.loc[alunos['id']==v,'nome'].iloc[0]} ({alunos.loc[alunos['id']==v,'email'].iloc[0]})",
                key="sel_excluir_aluno",
            )
            nome_excluir = alunos.loc[alunos["id"] == excluir_id, "nome"].iloc[0]
            confirmar = st.checkbox(
                f"Confirmo que desejo excluir permanentemente o aluno **{nome_excluir}** e todos os seus dados.",
                key="chk_excluir_aluno",
            )
            if st.button("🗑️ Excluir aluno", type="primary", disabled=not confirmar):
                nome_removido = _excluir_aluno_cascade(excluir_id)
                _toast_sucesso(f"Aluno '{nome_removido}' e todos os seus dados foram removidos com sucesso.")
                st.rerun()

    # ── Aba 2: Novo aluno ──
    with abas[1]:
        with st.form("novo_aluno", clear_on_submit=True):
            col1, col2 = st.columns(2)
            nome  = col1.text_input("Nome")
            email = col2.text_input("E-mail (opcional)")
            salvar = st.form_submit_button("➕ Adicionar aluno", use_container_width=True)
        if salvar:
            nome  = limpar_texto(nome)
            email = normalizar_email(email) or email_local(nome)
            if not nome:
                st.error("Informe o nome do aluno.")
            else:
                try:
                    executar(
                        "INSERT INTO alunos (nome, email, senha, perfil, ativo, force_troca_senha) VALUES (?, ?, ?, 'Aluno', 1, 1)",
                        (nome, email, hash_senha("123")),
                    )
                    limpar_cache()
                    _toast_sucesso(f"Aluno '{nome}' cadastrado com sucesso. Senha inicial: 123")
                    st.rerun()
                except (IntegrityError, UniqueViolation):
                    st.error("Já existe um aluno com esse nome ou e-mail.")

    # ── Aba 3: Limpeza de usuários ──
    with abas[2]:
        render_html(
            '<div class="hero" style="margin-bottom:16px">'
            '<h1 style="font-size:1.2rem">🧹 Limpeza de usuários</h1>'
            '<p>Identifique e remova usuários duplicados, sem execuções ou indesejados. '
            'Use os filtros abaixo para selecionar quem remover, confirme e execute a limpeza.</p>'
            '</div>'
        )

        # Carrega TODOS os alunos (inclusive inativos) com contagem de execuções
        todos = consultar("""
            SELECT
                a.id,
                a.nome,
                a.email,
                a.perfil,
                a.ativo,
                COUNT(e.id) AS qtd_execucoes,
                MAX(e.atualizado_em) AS ultima_atividade
            FROM alunos a
            LEFT JOIN execucoes e ON e.aluno_id = a.id
            WHERE a.perfil = 'Aluno'
            GROUP BY a.id, a.nome, a.email, a.perfil, a.ativo
            ORDER BY a.ativo DESC, qtd_execucoes ASC, a.nome
        """)

        if todos.empty:
            st.info("Nenhum aluno cadastrado.")
        else:
            # Métricas rápidas
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total de alunos", len(todos))
            c2.metric("Ativos", int(todos["ativo"].sum()))
            c3.metric("Inativos", int((todos["ativo"] == 0).sum()))
            c4.metric("Sem nenhuma execução", int((todos["qtd_execucoes"] == 0).sum()))

            st.markdown("---")

            # Filtros de seleção para limpeza
            col_f1, col_f2 = st.columns(2)
            mostrar_sem_exec = col_f1.checkbox("Mostrar apenas alunos sem execuções", value=False)
            mostrar_inativos = col_f2.checkbox("Mostrar apenas alunos inativos", value=False)

            df_view = todos.copy()
            if mostrar_sem_exec:
                df_view = df_view[df_view["qtd_execucoes"] == 0]
            if mostrar_inativos:
                df_view = df_view[df_view["ativo"] == 0]

            if df_view.empty:
                st.info("Nenhum aluno encontrado com os filtros aplicados.")
            else:
                st.dataframe(
                    df_view.rename(columns={
                        "id": "ID", "nome": "Nome", "email": "E-mail",
                        "perfil": "Perfil", "ativo": "Ativo",
                        "qtd_execucoes": "Execuções", "ultima_atividade": "Última atividade",
                    }),
                    use_container_width=True, hide_index=True,
                )

                st.markdown("---")
                render_html('<div class="section-title">🗑️ Remover selecionados</div>')
                render_html(
                    '<div class="rule-warning">'
                    '⚠️ A remoção exclui permanentemente o aluno e todos os seus dados de execução. '
                    'Esta ação <strong>não pode ser desfeita</strong>. Revise a lista acima antes de confirmar.'
                    '</div>'
                )

                ids_disponiveis = df_view["id"].tolist()
                nomes_map = dict(zip(df_view["id"], df_view["nome"]))

                ids_remover = st.multiselect(
                    "Selecione os alunos a remover",
                    ids_disponiveis,
                    format_func=lambda v: f"{nomes_map.get(v,'?')} — {int(df_view.loc[df_view['id']==v,'qtd_execucoes'].iloc[0])} execuções",
                    key="ms_limpar_alunos",
                )

                if ids_remover:
                    nomes_selecionados = [nomes_map.get(i, str(i)) for i in ids_remover]
                    confirmar_limpeza = st.checkbox(
                        f"Confirmo a remoção de {len(ids_remover)} aluno(s): {', '.join(nomes_selecionados)}",
                        key="chk_limpar_alunos",
                    )
                    if st.button("🧹 Executar limpeza", type="primary", disabled=not confirmar_limpeza):
                        removidos = []
                        erros = []
                        for aid in ids_remover:
                            try:
                                nome_rem = _excluir_aluno_cascade(aid)
                                removidos.append(nome_rem)
                            except Exception as exc:
                                erros.append(f"{nomes_map.get(aid,'?')}: {exc}")
                        if removidos:
                            _toast_sucesso(f"{len(removidos)} aluno(s) removido(s) com sucesso: {', '.join(removidos)}.")
                        if erros:
                            for e in erros:
                                st.error(f"Erro ao remover: {e}")
                        st.rerun()

                st.markdown("---")
                render_html('<div class="section-title">⚡ Limpeza rápida</div>')
                col_r1, col_r2 = st.columns(2)
                with col_r1:
                    st.caption("Remove todos os alunos **sem nenhuma execução** registrada.")
                    sem_exec_ids = todos[todos["qtd_execucoes"] == 0]["id"].tolist()
                    if sem_exec_ids:
                        conf_sem_exec = st.checkbox(
                            f"Confirmo remoção de {len(sem_exec_ids)} aluno(s) sem execuções",
                            key="chk_sem_exec",
                        )
                        if st.button("🗑️ Remover sem execuções", disabled=not conf_sem_exec, key="btn_sem_exec"):
                            for aid in sem_exec_ids:
                                _excluir_aluno_cascade(aid)
                            _toast_sucesso(f"{len(sem_exec_ids)} aluno(s) sem execuções removidos com sucesso.")
                            st.rerun()
                    else:
                        st.success("Nenhum aluno sem execuções.")

                with col_r2:
                    st.caption("Remove todos os alunos com status **inativo** (`ativo = 0`).")
                    inativos_ids = todos[todos["ativo"] == 0]["id"].tolist()
                    if inativos_ids:
                        conf_inativos = st.checkbox(
                            f"Confirmo remoção de {len(inativos_ids)} aluno(s) inativo(s)",
                            key="chk_inativos",
                        )
                        if st.button("🗑️ Remover inativos", disabled=not conf_inativos, key="btn_inativos"):
                            for aid in inativos_ids:
                                _excluir_aluno_cascade(aid)
                            _toast_sucesso(f"{len(inativos_ids)} aluno(s) inativo(s) removidos com sucesso.")
                            st.rerun()
                    else:
                        st.success("Nenhum aluno inativo.")


# ─────────────────────────────────────────────
# TELA: IMPORTAÇÕES
# ─────────────────────────────────────────────

def tela_importacao():
    st.header("Importações")

    abas = st.tabs([
        "📋 Planilha de referência",
        "📊 Ciclo Consolidado (multi-aluno)",
        "👤 Planilhas individuais (legado)",
    ])

    # ── Aba 1: Planilha de referência ──
    with abas[0]:
        render_html("""
            <div class="insight-card info" style="margin-bottom:12px">
              <div class="insight-icon">🛡️</div>
              <div class="insight-body">
                <div class="insight-title">Importação segura — histórico preservado</div>
                <p class="insight-text">
                  A importação da planilha de referência <strong>nunca apaga execuções existentes</strong>.
                  Apenas insere novas disciplinas, aulas, assuntos e tarefas que ainda não existam,
                  ou atualiza campos de estrutura (trilha, tipo, exercícios previstos).
                  Status, datas, horas e acertos já registrados são sempre preservados.
                </p>
              </div>
            </div>
        """)
        st.caption(f"Planilha de referência padrão: `{PLANILHA_REFERENCIA}`")
        arquivo = st.file_uploader("Planilha referência (.xlsx)", type=["xlsx"], key="ref_upload")
        if arquivo is not None:
            destino = BASE_DIR / "planilha_referencia_importada.xlsx"
            destino.write_bytes(arquivo.getbuffer())
            if st.button("Importar referência enviada", type="primary", key="ref_envio"):
                if importar_planilha_referencia(destino):
                    _toast_sucesso("Planilha de referência importada com sucesso. Histórico preservado.")
                    st.rerun()
        elif st.button("Reimportar referência padrão", key="ref_padrao"):
            if importar_planilha_referencia(PLANILHA_REFERENCIA):
                _toast_sucesso("Planilha de referência padrão importada com sucesso. Histórico preservado.")
                st.rerun()

        st.markdown("---")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Estudantes", int(consultar("SELECT COUNT(*) qtd FROM alunos WHERE perfil='Aluno' AND ativo=1").iloc[0].qtd))
        c2.metric("Disciplinas", int(consultar("SELECT COUNT(*) qtd FROM disciplinas WHERE ativo=1").iloc[0].qtd))
        c3.metric("Aulas",      int(consultar("SELECT COUNT(*) qtd FROM aulas WHERE ativo=1").iloc[0].qtd))
        c4.metric("Tarefas",    int(consultar("SELECT COUNT(*) qtd FROM tarefas WHERE ativo=1").iloc[0].qtd))
        c5.metric("Execuções",  int(consultar("SELECT COUNT(*) qtd FROM execucoes").iloc[0].qtd))

    # ── Aba 2: Ciclo Consolidado ──
    with abas[1]:
        render_html("""
            <div class="insight-card info" style="margin-bottom:14px">
              <div class="insight-icon">ℹ️</div>
              <div class="insight-body">
                <div class="insight-title">Formato esperado: Ciclo Consolidado</div>
                <p class="insight-text">
                  Arquivo com aba <strong>CICLO_CONSOLIDADO</strong>, linha de cabeçalho na linha 3.<br>
                  Colunas: <em>BLOCO · DISCIPLINA · OBJETIVO · STATUS · DATA PROGRAMADA ·
                  CH/AULA/QUESTÕES/ACERTOS (Aluno A) · CH/AULA/QUESTÕES/ACERTOS (Aluno B) ·
                  CH TOTAL · TOTAL QUESTÕES · TOTAL ACERTOS · DESEMPENHO</em>.<br>
                  Os nomes dos alunos são lidos automaticamente da linha 2 do arquivo.
                </p>
              </div>
            </div>
        """)

        arquivo_cc = st.file_uploader(
            "Arquivo Ciclo Consolidado (.xlsx)",
            type=["xlsx"],
            key="cc_upload",
        )

        col_modo, col_btn = st.columns([2, 1])
        modo = col_modo.radio(
            "Modo de importação",
            ["substituir", "acumular"],
            format_func=lambda v: "🔄 Substituir execuções existentes" if v == "substituir" else "➕ Acumular (não apaga execuções anteriores)",
            horizontal=True,
            key="cc_modo",
        )

        if arquivo_cc is not None:
            destino_cc = BASE_DIR / arquivo_cc.name
            destino_cc.write_bytes(arquivo_cc.getbuffer())

            # Pré-visualização: mostra nomes detectados
            try:
                import openpyxl as _opx
                _wb = _opx.load_workbook(str(destino_cc), read_only=True, data_only=True)
                if "CICLO_CONSOLIDADO" in _wb.sheetnames:
                    _linhas = list(_wb["CICLO_CONSOLIDADO"].iter_rows(values_only=True, max_row=3))
                    _nome_a = limpar_texto(_linhas[1][6]) if len(_linhas) > 1 and len(_linhas[1]) > 6 else "?"
                    _nome_b = limpar_texto(_linhas[1][10]) if len(_linhas) > 1 and len(_linhas[1]) > 10 else "?"
                    st.info(f"Alunos detectados: **{_nome_a}** e **{_nome_b}**")
            except Exception:
                pass

            if col_btn.button("▶ Importar", type="primary", use_container_width=True, key="cc_importar"):
                with st.spinner("Importando…"):
                    resultado = importar_ciclo_consolidado(destino_cc, modo=modo)

                if resultado["erros"]:
                    for e in resultado["erros"]:
                        st.error(e)
                if resultado["avisos"]:
                    with st.expander(f"⚠️ {len(resultado['avisos'])} aviso(s)"):
                        for av in resultado["avisos"]:
                            st.warning(av)
                if resultado["ok"]:
                    _toast_sucesso(
                        f"✅ Importação concluída! "
                        f"**{resultado['registros']}** execuções gravadas."
                    )
                    st.rerun()

    # ── Aba 3: Planilhas individuais (legado) ──
    with abas[2]:
        st.caption("Formato legado: um arquivo por aluno com aba CICLO_REG.")
        arquivos = st.file_uploader(
            "Planilhas individuais (.xlsx)",
            type=["xlsx"],
            accept_multiple_files=True,
            key="ind_upload",
        )
        if arquivos and st.button("Importar planilhas individuais", type="primary", key="ind_importar"):
            ok = 0
            for arq in arquivos:
                destino = BASE_DIR / arq.name
                destino.write_bytes(arq.getbuffer())
                ok += int(importar_execucoes_ciclo(destino))
            _toast_sucesso(f"Planilhas processadas com sucesso: {ok}.")
            st.rerun()


# ─────────────────────────────────────────────
# TELA: CONFIGURAÇÕES
# ─────────────────────────────────────────────

def tela_configuracoes():
    st.header("Configurações")
    usuario = aluno_logado()
    aba_senha, aba_usuarios, aba_backup = st.tabs(["Minha senha", "Usuários", "Backup"])

    with aba_senha:
        with st.form("senha"):
            atual     = st.text_input("Senha atual", type="password")
            nova      = st.text_input("Nova senha", type="password")
            confirmar = st.text_input("Confirmar nova senha", type="password")
            salvar    = st.form_submit_button("Alterar senha")
        if salvar:
            usuario_db = consultar("SELECT senha FROM alunos WHERE id = ?", (usuario["id"],))
            if usuario_db.empty or not verificar_senha(atual, usuario_db.iloc[0]["senha"]):
                st.error("Senha atual inválida.")
            elif not senha_valida(nova):
                st.error("Use uma senha com pelo menos 8 caracteres, contendo letras e números.")
            elif nova != confirmar:
                st.error("A confirmação não confere.")
            else:
                atualizar_senha_usuario(usuario["id"], nova, 0)
                _toast_sucesso("Senha alterada com sucesso.")

    with aba_usuarios:
        if usuario["perfil"] != "Gestor":
            st.info("Somente gestores podem administrar senhas e bloqueios.")
        else:
            usuarios = consultar("SELECT id, nome, email, perfil, ativo, force_troca_senha FROM alunos ORDER BY perfil, nome")
            st.dataframe(usuarios, use_container_width=True, hide_index=True)
            if not usuarios.empty:
                usuario_id = st.selectbox(
                    "Usuário", usuarios["id"].tolist(),
                    format_func=lambda v: f"{usuarios.loc[usuarios['id'] == v, 'nome'].iloc[0]} ({usuarios.loc[usuarios['id'] == v, 'perfil'].iloc[0]})",
                )
                col1, col2, col3 = st.columns(3)
                nova_senha = col1.text_input("Nova senha temporária", type="password")
                forcar     = col2.checkbox("Forçar troca no próximo login", value=True)
                if col3.button("Redefinir senha", type="primary"):
                    if not senha_valida(nova_senha):
                        st.error("A senha temporária deve ter pelo menos 8 caracteres, contendo letras e números.")
                    else:
                        atualizar_senha_usuario(usuario_id, nova_senha, 1 if forcar else 0)
                        _toast_sucesso("Senha redefinida com sucesso. O aluno deverá trocar no próximo acesso.")
                        st.rerun()
                ativo_atual = int(usuarios.loc[usuarios["id"] == usuario_id, "ativo"].iloc[0])
                col4, col5 = st.columns(2)
                if ativo_atual == 1 and col4.button("Bloquear usuário"):
                    executar("UPDATE alunos SET ativo = 0 WHERE id = ?", (int(usuario_id),))
                    _toast_sucesso("Usuário bloqueado com sucesso.")
                    st.rerun()
                if ativo_atual == 0 and col5.button("Reativar usuário"):
                    executar("UPDATE alunos SET ativo = 1, force_troca_senha = 1 WHERE id = ?", (int(usuario_id),))
                    _toast_sucesso("Usuário reativado. Será solicitado troca de senha no próximo acesso.")
                    st.rerun()

    with aba_backup:
        if usuario["perfil"] != "Gestor":
            st.info("Somente gestores podem exportar backups.")
        else:
            st.caption("Exportação lógica para backup e auditoria.")
            tabelas = {
                "alunos": consultar("SELECT id, nome, email, perfil, ativo, force_troca_senha FROM alunos"),
                "disciplinas": consultar("SELECT * FROM disciplinas"),
                "aulas": consultar("SELECT * FROM aulas"),
                "assuntos": consultar("SELECT * FROM assuntos"),
                "tarefas": consultar("SELECT * FROM tarefas"),
                "execucoes": consultar("SELECT * FROM execucoes"),
            }
            for nome_tab, dados in tabelas.items():
                st.download_button(
                    f"Exportar {nome_tab}.csv",
                    data=dados.to_csv(index=False).encode("utf-8"),
                    file_name=f"{nome_tab}_{date.today()}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )


# ─────────────────────────────────────────────
# INICIALIZAÇÃO E MAIN
# ─────────────────────────────────────────────

def inicializar():
    criar_tabelas()
    inserir_admin()
    qtd = consultar("SELECT COUNT(*) AS qtd FROM tarefas").iloc[0]["qtd"]
    if qtd == 0 and PLANILHA_REFERENCIA.exists():
        importar_planilha_referencia(PLANILHA_REFERENCIA, substituir=True)


def main():
    inicializar()

    if "usuario" not in st.session_state:
        tela_login()
        return

    usuario = aluno_logado()
    if int(usuario.get("force_troca_senha") or 0) == 1:
        tela_troca_obrigatoria()
        return

    st.sidebar.title(APP_NAME)
    st.sidebar.write(f"**Usuário:** {usuario['nome']}")
    st.sidebar.write(f"**Perfil:** {usuario['perfil']}")
    if st.sidebar.button("Sair", use_container_width=True):
        st.session_state.clear()
        st.rerun()

    if usuario["perfil"] == "Gestor":
        opcoes = ["Dashboard", "Registro rápido", "Tarefas", "Aulas e assuntos", "Disciplinas", "Alunos", "Importações", "Configurações"]
    else:
        opcoes = ["Dashboard", "Registro rápido", "Tarefas", "Aulas e assuntos", "Configurações"]

    opcao = st.sidebar.radio("Navegação", opcoes)
    paginas = {
        "Dashboard":       dashboard,
        "Registro rápido": tela_registro_rapido,
        "Tarefas":         tela_tarefas,
        "Aulas e assuntos":tela_aulas,
        "Disciplinas":     tela_disciplinas,
        "Alunos":          tela_alunos,
        "Importações":     tela_importacao,
        "Configurações":   tela_configuracoes,
    }
    paginas[opcao]()


if __name__ == "__main__":
    main()
