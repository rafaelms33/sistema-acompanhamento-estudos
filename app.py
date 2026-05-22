"""
Sistema de Acompanhamento de Estudos
Arquivo único — refatorado para legibilidade, layout e regras de negócio do Registro Rápido.
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

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ─────────────────────────────────────────────
# CSS GLOBAL
# ─────────────────────────────────────────────

CSS = """
<style>
  :root { color-scheme: light; }

  /* Layout */
  .block-container {
    padding: 1.25rem 2rem 2rem;
    max-width: 1480px;
  }

  /* Sidebar */
  [data-testid="stSidebar"] {
    background: #0f172a;
  }
  [data-testid="stSidebar"] * { color: #f8fafc !important; }
  [data-testid="stSidebar"] .stRadio label { padding: 6px 0; }

  /* Métricas */
  [data-testid="stMetric"] {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 14px 18px;
    box-shadow: 0 1px 3px rgba(15,23,42,.05);
  }
  [data-testid="stMetricValue"] { font-size: 1.45rem; font-weight: 800; color: #0f172a; }
  [data-testid="stMetricLabel"] { color: #475569; font-size: .85rem; }

  /* DataFrames */
  div[data-testid="stDataFrame"] {
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    overflow: hidden;
  }

  /* Hero banner */
  .hero {
    border: 1px solid #dbeafe;
    background: linear-gradient(135deg, #eff6ff 0%, #f8fafc 50%, #ecfdf5 100%);
    border-radius: 10px;
    padding: 20px 24px;
    margin-bottom: 16px;
  }
  .hero h1 { margin: 0 0 4px; font-size: 1.6rem; color: #0f172a; font-weight: 800; }
  .hero p  { margin: 0; color: #475569; font-size: .95rem; }

  /* Cards de atividade */
  .quick-card {
    border: 1px solid #cbd5e1;
    border-left: 6px solid #94a3b8;
    border-radius: 10px;
    padding: 16px 20px;
    background: #ffffff;
    box-shadow: 0 4px 12px rgba(15,23,42,.06);
    margin-bottom: 16px;
  }
  .quick-card.ok   { border-left-color: #22c55e; }
  .quick-card.warn { border-left-color: #f59e0b; }
  .quick-card.off  { border-left-color: #94a3b8; }
  .quick-card h3   { margin: 4px 0 0; font-size: 1.05rem; color: #0f172a; }

  /* Grid interna do card */
  .quick-grid {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 10px;
    margin-top: 14px;
  }
  @media (max-width: 900px) { .quick-grid { grid-template-columns: repeat(2, 1fr); } }
  @media (max-width: 560px) { .quick-grid { grid-template-columns: 1fr; } }

  .quick-item {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 10px 12px;
  }
  .quick-label {
    color: #64748b;
    font-size: .72rem;
    text-transform: uppercase;
    font-weight: 800;
    letter-spacing: .04em;
  }
  .quick-value {
    color: #0f172a;
    font-weight: 700;
    margin-top: 3px;
    overflow-wrap: anywhere;
    font-size: .93rem;
  }

  /* Badges de status */
  .status-ok, .status-warn, .status-off {
    display: inline-flex;
    align-items: center;
    padding: 4px 12px;
    border-radius: 999px;
    font-weight: 800;
    font-size: .82rem;
    margin-top: 6px;
  }
  .status-ok   { background: #dcfce7; color: #166534; }
  .status-warn { background: #fef3c7; color: #92400e; }
  .status-off  { background: #f1f5f9; color: #475569; }

  /* Utilitários */
  .muted         { color: #64748b; font-size: .88rem; }
  .section-title { font-size: 1rem; font-weight: 800; color: #0f172a; margin: 10px 0 6px; }

  /* Bloco de aviso de regra de negócio */
  .rule-warning {
    background: #fffbeb;
    border: 1px solid #fde68a;
    border-left: 5px solid #f59e0b;
    border-radius: 8px;
    padding: 14px 16px;
    margin: 12px 0;
    color: #78350f;
    font-size: .93rem;
  }
  .rule-error {
    background: #fef2f2;
    border: 1px solid #fecaca;
    border-left: 5px solid #ef4444;
    border-radius: 8px;
    padding: 14px 16px;
    margin: 12px 0;
    color: #7f1d1d;
    font-size: .93rem;
  }

  /* Formulários */
  .stForm { border: 1px solid #e2e8f0 !important; border-radius: 10px !important; padding: 16px !important; }

  /* Botões primários */
  .stButton > button[kind="primary"] {
    background: #0f172a;
    color: white;
    border-radius: 8px;
    font-weight: 700;
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
    classe = {
        STATUS_CONCLUIDA:   "status-ok",
        STATUS_EM_ANDAMENTO:"status-warn",
        STATUS_NAO_INICIADA:"status-off",
    }.get(status, "status-off")
    return f'<span class="{classe}">{escape_html(STATUS_LABELS.get(status, status))}</span>'


def status_card_class(status: str) -> str:
    return {STATUS_CONCLUIDA: "ok", STATUS_EM_ANDAMENTO: "warn", STATUS_NAO_INICIADA: "off"}.get(status, "off")


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

def importar_planilha_referencia(caminho=PLANILHA_REFERENCIA, substituir=True):
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
            if substituir:
                conn.execute("DELETE FROM execucoes")
                conn.execute("DELETE FROM tarefas")
                conn.execute("DELETE FROM assuntos")
                conn.execute("DELETE FROM aulas")
                conn.execute("DELETE FROM disciplinas")

            for aba in xl.sheet_names:
                if aba in abas_ignoradas:
                    continue
                raw = pd.read_excel(caminho, sheet_name=aba, header=None, nrows=1)
                nome_disciplina = limpar_texto(raw.iloc[0, 0]) or aba
                disciplina_id = upsert_disciplina(conn, nome_disciplina)
                aulas_df = pd.read_excel(caminho, sheet_name=aba, header=1).dropna(how="all")
                if "Aula" not in aulas_df.columns or "Assunto" not in aulas_df.columns:
                    continue
                for _, row in aulas_df.iterrows():
                    aula = limpar_texto(row.get("Aula"))
                    assunto = limpar_texto(row.get("Assunto"))
                    if not aula or not assunto:
                        continue
                    aula_id = upsert_aula(
                        conn, disciplina_id, aula,
                        limpar_texto(row.get("Estudada")) or "Não",
                        limpar_texto(row.get("Revisão 24h")) or "Não",
                    )
                    upsert_assunto(conn, aula_id, assunto)

            tarefas = pd.read_excel(caminho, sheet_name="Tarefas", header=2).dropna(how="all")
            for _, row in tarefas.iterrows():
                numero = converter_inteiro(row.get("Tarefa"))
                disciplina = limpar_texto(row.get("Disciplina"))
                if not numero or not disciplina:
                    continue
                disciplina_id = upsert_disciplina(conn, disciplina)
                aula_valor = str(limpar_texto(row.get("Aula")) or "Aula não informada")
                aula_id = upsert_aula(conn, disciplina_id, aula_valor)
                assunto_id = upsert_assunto(conn, aula_id, limpar_texto(row.get("Conteúdo")))
                tipo = normalizar_tipo_estudo(row.get("Tipo"))
                conn.execute(
                    """
                    INSERT INTO tarefas
                        (numero, trilha, disciplina_id, seq_disciplina, aula, qtd_exercicios_previstos,
                         tipo, conteudo, ativo, aula_id, assunto_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
                    ON CONFLICT(numero) DO UPDATE SET
                        trilha = excluded.trilha,
                        disciplina_id = excluded.disciplina_id,
                        seq_disciplina = excluded.seq_disciplina,
                        aula = excluded.aula,
                        qtd_exercicios_previstos = excluded.qtd_exercicios_previstos,
                        tipo = excluded.tipo,
                        conteudo = excluded.conteudo,
                        ativo = 1,
                        aula_id = excluded.aula_id,
                        assunto_id = excluded.assunto_id
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
            st.success("Filtro salvo.")

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


def calcular_metricas(df):
    analisavel   = base_metricas(df)
    concluidas   = analisavel[analisavel["status"] == STATUS_CONCLUIDA]
    andamento    = analisavel[analisavel["status"] == STATUS_EM_ANDAMENTO]
    total        = len(analisavel)
    qtd_conc     = len(concluidas)
    progresso    = qtd_conc / total * 100 if total else 0
    horas        = float(analisavel["ch_efetiva"].sum())
    questoes     = int(analisavel["qtd_questoes_feitas"].sum())
    acertos      = int(analisavel["qtd_acertos"].sum())
    desempenho   = acertos / questoes * 100 if questoes else 0
    dias         = analisavel["data_ref"].dropna().dt.date.nunique() if "data_ref" in analisavel else 0
    return {
        "analisavel": analisavel,
        "concluidas_df": concluidas,
        "andamento_df": andamento,
        "total": total,
        "concluidas": qtd_conc,
        "andamento": len(andamento),
        "nao_iniciadas": len(df[df["status"] == STATUS_NAO_INICIADA]),
        "progresso": progresso,
        "horas": horas,
        "questoes": questoes,
        "acertos": acertos,
        "desempenho": desempenho,
        "media_dia": horas / dias if dias else 0,
        "produtividade": qtd_conc / horas if horas else 0,
    }


def fig_layout(fig):
    fig.update_layout(
        legend_title_text="Legenda",
        margin=dict(l=10, r=10, t=56, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#0f172a"),
    )
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
            st.success("Senha alterada com sucesso.")
            st.rerun()


# ─────────────────────────────────────────────
# TELA: DASHBOARD
# ─────────────────────────────────────────────

def dashboard():
    render_html(f"""
        <div class="hero">
          <h1>{APP_NAME}</h1>
          <p>Visão consolidada, análise individual, evolução e comparação entre alunos.</p>
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

    metricas = calcular_metricas(df_filtrado)
    analisavel = metricas["analisavel"]
    st.caption(
        "Os indicadores e gráficos consideram apenas atividades Em andamento ou Concluídas. "
        "Atividades Não iniciadas aparecem como fila, mas não entram nos percentuais."
    )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Taxa de conclusão", f"{metricas['progresso']:.1f}%")
    c2.metric("Tarefas concluídas", metricas["concluidas"])
    c3.metric("Em andamento", metricas["andamento"])
    c4.metric("Não iniciadas", metricas["nao_iniciadas"])

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Horas estudadas", f"{metricas['horas']:.2f}")
    c6.metric("Média por dia ativo", f"{metricas['media_dia']:.2f}h")
    c7.metric("Desempenho", f"{metricas['desempenho']:.1f}%")
    c8.metric("Produtividade", f"{metricas['produtividade']:.2f} tarefas/h")

    abas = st.tabs(["Visão geral", "Disciplinas", "Tipos de estudo", "Comparação", "Evolução", "Análise inteligente", "Atividades"])

    with abas[0]:
        col_a, col_b = st.columns([1.25, 1])
        status_df = (
            df_filtrado.groupby("status", as_index=False)["tarefa_id"]
            .count()
            .rename(columns={"tarefa_id": "tarefas"})
        )
        status_df["status_label"] = status_df["status"].map(STATUS_LABELS)
        col_a.plotly_chart(
            fig_layout(px.bar(
                status_df, x="status_label", y="tarefas", color="status",
                color_discrete_map=STATUS_CORES, text_auto=True,
                title="Distribuição das atividades por status",
                labels={"status_label": "Status", "tarefas": "Quantidade", "status": "Legenda"},
            )),
            use_container_width=True,
        )
        if not analisavel.empty:
            por_aluno = (
                analisavel.groupby("aluno", as_index=False)
                .agg(horas=("ch_efetiva", "sum"), concluidas=("status", lambda s: (s == STATUS_CONCLUIDA).sum()))
                .sort_values("horas", ascending=False)
            )
            col_b.plotly_chart(
                fig_layout(px.bar(
                    por_aluno, x="aluno", y=["horas", "concluidas"], barmode="group",
                    title="Produtividade consolidada por aluno",
                    labels={"aluno": "Aluno", "value": "Total", "variable": "Indicador"},
                )),
                use_container_width=True,
            )
        else:
            col_b.info("Sem atividades iniciadas ou concluídas para os filtros.")

    with abas[1]:
        if analisavel.empty:
            st.info("Sem dados analisáveis.")
        else:
            disc = (
                analisavel.groupby("disciplina", as_index=False)
                .agg(
                    tarefas=("tarefa_id", "count"),
                    horas=("ch_efetiva", "sum"),
                    questoes=("qtd_questoes_feitas", "sum"),
                    acertos=("qtd_acertos", "sum"),
                    concluidas=("status", lambda s: (s == STATUS_CONCLUIDA).sum()),
                )
            )
            disc["progresso"]  = disc["concluidas"] / disc["tarefas"] * 100
            disc["desempenho"] = disc.apply(lambda r: r["acertos"] / r["questoes"] * 100 if r["questoes"] else 0, axis=1)
            disc["disciplina_label"] = disc["disciplina"].apply(quebrar_texto)
            col_a, col_b = st.columns(2)
            col_a.plotly_chart(
                fig_layout(px.bar(disc.sort_values("progresso"), x="disciplina_label", y="progresso",
                    text_auto=".1f", title="Progresso por disciplina",
                    labels={"disciplina_label": "Disciplina", "progresso": "% concluído"})),
                use_container_width=True,
            )
            col_b.plotly_chart(
                fig_layout(px.bar(disc.sort_values("horas", ascending=False), x="disciplina_label", y="horas",
                    text_auto=".1f", title="Tempo estudado por disciplina",
                    labels={"disciplina_label": "Disciplina", "horas": "Horas efetivas"})),
                use_container_width=True,
            )
            st.dataframe(disc.drop(columns=["disciplina_label"]), use_container_width=True, hide_index=True)

    with abas[2]:
        if analisavel.empty:
            st.info("Sem dados analisáveis.")
        else:
            tipos = (
                analisavel.groupby("tipo", as_index=False)
                .agg(horas=("ch_efetiva", "sum"), tarefas=("tarefa_id", "count"), questoes=("qtd_questoes_feitas", "sum"))
                .sort_values("horas", ascending=False)
            )
            tipos["produtividade"] = tipos.apply(lambda r: r["tarefas"] / r["horas"] if r["horas"] else 0, axis=1)
            col_a, col_b = st.columns(2)
            col_a.plotly_chart(fig_layout(px.pie(tipos, names="tipo", values="horas", title="Tipos de estudo mais utilizados")), use_container_width=True)
            col_b.plotly_chart(
                fig_layout(px.bar(tipos, x="tipo", y="produtividade", text_auto=".2f",
                    title="Produtividade por tipo de estudo",
                    labels={"tipo": "Tipo de estudo", "produtividade": "Tarefas por hora"})),
                use_container_width=True,
            )
            st.dataframe(tipos, use_container_width=True, hide_index=True)

    with abas[3]:
        if analisavel.empty:
            st.info("Sem dados para comparação.")
        else:
            comp = (
                analisavel.groupby("aluno", as_index=False)
                .agg(
                    horas=("ch_efetiva", "sum"),
                    tarefas=("tarefa_id", "count"),
                    concluidas=("status", lambda s: (s == STATUS_CONCLUIDA).sum()),
                    questoes=("qtd_questoes_feitas", "sum"),
                    acertos=("qtd_acertos", "sum"),
                    disciplinas=("disciplina", "nunique"),
                )
            )
            comp["desempenho"]   = comp.apply(lambda r: r["acertos"] / r["questoes"] * 100 if r["questoes"] else 0, axis=1)
            comp["produtividade"] = comp.apply(lambda r: r["concluidas"] / r["horas"] if r["horas"] else 0, axis=1)
            comp = comp.sort_values(["produtividade", "concluidas"], ascending=False)
            st.plotly_chart(
                fig_layout(px.scatter(
                    comp, x="horas", y="concluidas", size="questoes", color="desempenho", hover_name="aluno",
                    title="Comparação entre alunos: tempo, conclusão e desempenho",
                    labels={"horas": "Horas estudadas", "concluidas": "Tarefas concluídas", "desempenho": "Desempenho (%)"},
                )),
                use_container_width=True,
            )
            st.dataframe(comp, use_container_width=True, hide_index=True)

    with abas[4]:
        evolucao = analisavel.dropna(subset=["data_ref"]).copy() if not analisavel.empty else pd.DataFrame()
        if evolucao.empty:
            st.info("Sem datas de execução para evolução.")
        else:
            evolucao["dia"] = evolucao["data_ref"].dt.date
            diario = evolucao.groupby(["dia", "status"], as_index=False).agg(horas=("ch_efetiva", "sum"), tarefas=("tarefa_id", "count"))
            diario["dia_label"] = pd.to_datetime(diario["dia"]).dt.strftime("%d/%m/%Y")
            st.plotly_chart(
                fig_layout(px.bar(diario, x="dia_label", y="horas", color="status", color_discrete_map=STATUS_CORES,
                    title="Evolução diária da carga horária efetiva",
                    labels={"dia_label": "Data", "horas": "Carga horária (h)", "status": "Status"},
                    hover_data={"tarefas": True, "dia": False})),
                use_container_width=True,
            )
            semanal = evolucao.set_index("data_ref").resample("W")["ch_efetiva"].sum().reset_index()
            semanal["semana"] = semanal["data_ref"].dt.strftime("%d/%m/%Y")
            st.plotly_chart(
                fig_layout(px.line(semanal, x="semana", y="ch_efetiva", markers=True,
                    title="Evolução semanal de estudos",
                    labels={"semana": "Semana encerrada em", "ch_efetiva": "Horas efetivas"})),
                use_container_width=True,
            )

    with abas[5]:
        if analisavel.empty:
            st.info("Sem atividades iniciadas ou concluídas para análise.")
        else:
            alunos_lista = sorted(analisavel["aluno"].dropna().unique().tolist()) if visao == "Todos" else [visao]
            for aluno in alunos_lista:
                grupo = analisavel[analisavel["aluno"] == aluno] if aluno != "Todos" else analisavel
                m = calcular_metricas(grupo)
                st.markdown(f"**{aluno}**")
                if m["horas"] == 0:
                    st.warning("Sem carga horária registrada nas atividades iniciadas ou concluídas.")
                elif m["desempenho"] and m["desempenho"] < 70:
                    st.warning(f"Desempenho de {m['desempenho']:.1f}%. Reforce revisão ativa e questões comentadas.")
                elif m["progresso"] < 45:
                    st.info(f"Taxa de conclusão de {m['progresso']:.1f}%. Priorize concluir atividades já iniciadas.")
                else:
                    st.success(f"Ritmo consistente: {m['horas']:.1f}h, {m['concluidas']} tarefas concluídas, {m['produtividade']:.2f} tarefas/h.")
                fracas = (
                    grupo.groupby("disciplina", as_index=False)
                    .agg(questoes=("qtd_questoes_feitas", "sum"), acertos=("qtd_acertos", "sum"), horas=("ch_efetiva", "sum"))
                )
                fracas["desempenho"] = fracas.apply(lambda r: r["acertos"] / r["questoes"] * 100 if r["questoes"] else 0, axis=1)
                fracas = fracas[(fracas["questoes"] > 0) & (fracas["desempenho"] < 70)].sort_values("desempenho")
                if not fracas.empty:
                    st.caption("Disciplinas para atenção: " + ", ".join(fracas["disciplina"].head(4).tolist()))

    with abas[6]:
        tabela = preparar_tabela(df_filtrado)
        colunas = ["aluno", "tarefa", "disciplina", "aula", "assunto", "tipo", "status_label", "data_execucao", "ch_efetiva", "qtd_questoes_feitas", "qtd_acertos", "desempenho", "comentario"]
        colunas_disp = [c for c in colunas if c in tabela.columns]
        st.dataframe(tabela[colunas_disp], use_container_width=True, hide_index=True, row_height=72)


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


def _formulario_registro(tarefa, prefixo="rr", mostrar_confirmacao_concluida=False):
    """Renderiza o formulário de registro e retorna os dados preenchidos."""
    status_atual = str(tarefa.get("status", STATUS_NAO_INICIADA))

    if mostrar_confirmacao_concluida:
        render_html("""
            <div class="rule-warning">
              ⚠️ <strong>Atenção:</strong> Esta tarefa já está <strong>Concluída</strong>.
              Alterar seus dados pode modificar o histórico de estudos.
              Confirme abaixo para prosseguir.
            </div>
        """)

    with st.form(f"form_{prefixo}_{tarefa['tarefa_id']}"):
        col1, col2, col3, col4 = st.columns(4)
        novo_status = col1.selectbox(
            "Status",
            STATUS_VALIDOS,
            index=STATUS_VALIDOS.index(status_atual),
            format_func=lambda v: STATUS_LABELS[v],
        )
        tipo_idx = TIPOS_ESTUDO.index(tarefa["tipo"]) if tarefa["tipo"] in TIPOS_ESTUDO else TIPOS_ESTUDO.index("Outro")
        tipo_estudo  = col2.selectbox("Tipo de estudo", TIPOS_ESTUDO, index=tipo_idx)
        data_estudo  = col3.date_input("Data do estudo", value=date.today())
        ch           = col4.number_input("Tempo gasto (h)", min_value=0.0, value=float(tarefa["ch_efetiva"] or 0), step=0.25)
        questoes     = st.number_input("Questões feitas", min_value=0, value=int(tarefa["qtd_questoes_feitas"] or 0), step=1)
        acertos      = st.number_input("Acertos", min_value=0, value=int(tarefa["qtd_acertos"] or 0), step=1)
        comentario   = st.text_area("Observações", value=tarefa["comentario"] or "")

        if mostrar_confirmacao_concluida:
            confirmado = st.checkbox("✅ Confirmo que desejo alterar uma tarefa já concluída")
        else:
            confirmado = True

        salvar = st.form_submit_button("Registrar atividade", use_container_width=True, type="primary")

    return salvar, novo_status, tipo_estudo, data_estudo, ch, questoes, acertos, comentario, confirmado


def _renderizar_secao_registro(tarefas_df, aluno_id, status_alvo, titulo, key_prefix):
    """Renderiza uma seção de registro para um grupo de tarefas (por status)."""
    grupo = tarefas_df[tarefas_df["status"] == status_alvo].copy()
    if grupo.empty:
        st.info(f"Nenhuma tarefa {STATUS_LABELS[status_alvo].lower()}.")
        return

    opcoes = grupo["tarefa_id"].tolist()
    tarefa_id = st.selectbox(
        f"Tarefa {STATUS_LABELS[status_alvo].lower()}",
        opcoes,
        format_func=lambda v: _label_tarefa(grupo[grupo["tarefa_id"] == v].iloc[0]),
        key=f"{key_prefix}_sel",
    )
    tarefa = grupo[grupo["tarefa_id"] == tarefa_id].iloc[0]

    # Exibe detalhes da tarefa selecionada
    st.dataframe(
        pd.DataFrame([{
            "Disciplina": tarefa["disciplina"],
            "Aula": tarefa["aula"],
            "Assunto": tarefa["assunto"],
            "Tipo": tarefa["tipo"],
            "Conteúdo": tarefa["conteudo"],
            "Exercícios previstos": tarefa["qtd_exercicios_previstos"],
        }]),
        use_container_width=True,
        hide_index=True,
    )

    eh_concluida = (status_alvo == STATUS_CONCLUIDA)

    salvar, novo_status, tipo_estudo, data_estudo, ch, questoes, acertos, comentario, confirmado = (
        _formulario_registro(tarefa, prefixo=key_prefix, mostrar_confirmacao_concluida=eh_concluida)
    )

    if not salvar:
        return

    # Validação: acertos x questões
    if acertos > questoes:
        st.error("Acertos não podem ser maiores que questões feitas.")
        return

    # Validação: tarefa concluída exige confirmação
    if eh_concluida and not confirmado:
        render_html('<div class="rule-error">Marque a confirmação para alterar uma tarefa já concluída.</div>')
        return

    # Validação: regra de andamento por disciplina
    pode, msg_erro = _verificar_regra_andamento(tarefas_df, tarefa, novo_status, int(tarefa["disciplina_id"]))
    if not pode:
        render_html(f'<div class="rule-error">{escape_html(msg_erro)}</div>')
        return

    with conectar() as conn:
        upsert_execucao(conn, aluno_id, int(tarefa_id), str(data_estudo), ch, None, 0, acertos, comentario, questoes, novo_status, tipo_estudo)
    limpar_cache()
    st.success("Atividade registrada com sucesso.")
    st.rerun()


def tela_registro_rapido():
    render_html('<div class="hero"><h1>Registro Rápido</h1><p>Registre o progresso de uma tarefa de forma ágil.</p></div>')

    usuario = aluno_logado()
    alunos  = alunos_ativos()

    if usuario["perfil"] == "Aluno":
        aluno_id = int(usuario["id"])
        st.text_input("Aluno", value=usuario["nome"], disabled=True)
    else:
        if alunos.empty:
            st.info("Cadastre um aluno para registrar atividades.")
            return
        aluno_id = st.selectbox(
            "Aluno",
            alunos["id"].tolist(),
            format_func=lambda v: alunos.loc[alunos["id"] == v, "nome"].iloc[0],
        )

    exibir_card_ultima(ultima_atividade(aluno_id))

    tarefas = carregar_visao_tarefas(aluno_id)
    if tarefas.empty:
        st.info("Cadastre tarefas para registrar estudos.")
        return

    # Separação em três abas por status
    aba_nao, aba_and, aba_conc = st.tabs([
        f"🔘 Não iniciadas ({len(tarefas[tarefas['status'] == STATUS_NAO_INICIADA])})",
        f"🟡 Em andamento ({len(tarefas[tarefas['status'] == STATUS_EM_ANDAMENTO])})",
        f"🟢 Concluídas ({len(tarefas[tarefas['status'] == STATUS_CONCLUIDA])})",
    ])

    with aba_nao:
        st.caption("Uma tarefa não iniciada só pode ser iniciada se não houver outra tarefa em andamento para a mesma disciplina.")
        _renderizar_secao_registro(tarefas, aluno_id, STATUS_NAO_INICIADA, "Não iniciadas", "rr_nao")

    with aba_and:
        _renderizar_secao_registro(tarefas, aluno_id, STATUS_EM_ANDAMENTO, "Em andamento", "rr_and")

    with aba_conc:
        st.caption("⚠️ Alterações em tarefas concluídas exigem confirmação e podem modificar dados históricos.")
        _renderizar_secao_registro(tarefas, aluno_id, STATUS_CONCLUIDA, "Concluídas", "rr_conc")

    # Histórico recente
    recentes = carregar_execucoes()
    recentes = recentes[recentes["aluno_id"] == aluno_id].head(10)
    if not recentes.empty:
        st.markdown("---")
        st.subheader("Histórico recente")
        colunas = ["tarefa", "disciplina", "assunto", "tipo", "status_label", "data_execucao", "ch_efetiva", "comentario"]
        tabela = preparar_tabela(recentes)
        st.dataframe(tabela[[c for c in colunas if c in tabela.columns]], use_container_width=True, hide_index=True, row_height=72)


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
                st.success("Status e registros atualizados.")
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
                    st.success("Tarefa criada.")
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
                st.success("Tarefas atualizadas.")
                st.rerun()
            excluir = col2.selectbox("Excluir tarefa", tarefas["tarefa_id"].tolist(), format_func=lambda v: f"Tarefa {int(tarefas.loc[tarefas['tarefa_id'] == v, 'tarefa'].iloc[0])}")
            if col2.button("Excluir tarefa selecionada"):
                executar("UPDATE tarefas SET ativo = 0 WHERE id = ?", (int(excluir),))
                limpar_cache()
                st.success("Tarefa excluída.")
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
                st.success("Vínculos atualizados.")
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
                    st.success("Aula criada.")
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
                    st.success("Aulas atualizadas.")
                    st.rerun()
                excluir = col2.selectbox("Excluir aula", aulas["id"].tolist(), format_func=lambda v: aulas.loc[aulas["id"] == v, "aula"].iloc[0])
                if col2.button("Excluir aula selecionada"):
                    executar("UPDATE aulas SET ativo = 0 WHERE id = ?", (int(excluir),))
                    limpar_cache()
                    st.success("Aula excluída.")
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
                    st.success("Assunto criado.")
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
                    st.success("Assuntos atualizados.")
                    st.rerun()
                excluir = col2.selectbox("Excluir assunto", assuntos["id"].tolist(), format_func=lambda v: assuntos.loc[assuntos["id"] == v, "assunto"].iloc[0])
                if col2.button("Excluir assunto selecionado"):
                    executar("UPDATE assuntos SET ativo = 0 WHERE id = ?", (int(excluir),))
                    limpar_cache()
                    st.success("Assunto excluído.")
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
                st.success("Disciplina criada.")
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
                st.success("Disciplinas atualizadas.")
                st.rerun()
            excluir = col2.selectbox("Excluir disciplina", df["id"].tolist(), format_func=lambda v: df.loc[df["id"] == v, "nome"].iloc[0])
            if col2.button("Excluir disciplina selecionada"):
                executar("UPDATE disciplinas SET ativo = 0 WHERE id = ?", (int(excluir),))
                limpar_cache()
                st.success("Disciplina excluída.")
                st.rerun()

    st.dataframe(df, use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────
# TELA: ALUNOS
# ─────────────────────────────────────────────

def tela_alunos():
    st.header("Alunos")
    usuario = aluno_logado()
    if usuario["perfil"] != "Gestor":
        st.info("Somente gestores podem administrar alunos.")
        return

    with st.form("novo_aluno", clear_on_submit=True):
        col1, col2 = st.columns(2)
        nome   = col1.text_input("Nome")
        email  = col2.text_input("E-mail")
        salvar = st.form_submit_button("Adicionar aluno")
    if salvar:
        nome  = limpar_texto(nome)
        email = normalizar_email(email) or email_local(nome)
        if not nome:
            st.error("Informe o nome.")
        else:
            try:
                executar(
                    "INSERT INTO alunos (nome, email, senha, perfil, ativo, force_troca_senha) VALUES (?, ?, ?, 'Aluno', 1, 1)",
                    (nome, email, hash_senha("123")),
                )
                limpar_cache()
                st.success("Aluno cadastrado. Senha inicial: 123")
                st.rerun()
            except (IntegrityError, UniqueViolation):
                st.error("Já existe aluno com esse nome ou e-mail.")

    alunos = alunos_ativos(incluir_gestor=False)
    if alunos.empty:
        st.info("Nenhum aluno cadastrado.")
        return

    editado = st.data_editor(alunos, use_container_width=True, hide_index=True, disabled=["id", "perfil"], key="editor_alunos")
    col1, col2 = st.columns(2)
    if col1.button("Salvar alterações de alunos", type="primary"):
        try:
            with conectar() as conn:
                for _, row in editado.iterrows():
                    conn.execute(
                        "UPDATE alunos SET nome = ?, email = ? WHERE id = ? AND perfil = 'Aluno'",
                        (limpar_texto(row["nome"]), normalizar_email(row["email"]), int(row["id"])),
                    )
            limpar_cache()
            st.success("Alunos atualizados.")
            st.rerun()
        except (IntegrityError, UniqueViolation):
            st.error("Nome ou e-mail duplicado.")
    excluir_id = col2.selectbox(
        "Excluir aluno", alunos["id"].tolist(),
        format_func=lambda v: alunos.loc[alunos["id"] == v, "nome"].iloc[0],
    )
    if col2.button("Excluir aluno selecionado"):
        executar("UPDATE alunos SET ativo = 0 WHERE id = ? AND perfil = 'Aluno'", (int(excluir_id),))
        limpar_cache()
        st.success("Aluno excluído.")
        st.rerun()


# ─────────────────────────────────────────────
# TELA: IMPORTAÇÕES
# ─────────────────────────────────────────────

def tela_importacao():
    st.header("Importações")
    st.caption(f"Planilha de referência atual: `{PLANILHA_REFERENCIA}`")
    col1, col2 = st.columns(2)
    with col1:
        arquivo = st.file_uploader("Planilha referência (.xlsx)", type=["xlsx"])
        if arquivo is not None:
            destino = BASE_DIR / "planilha_referencia_importada.xlsx"
            destino.write_bytes(arquivo.getbuffer())
            if st.button("Importar referência enviada", type="primary"):
                if importar_planilha_referencia(destino, substituir=True):
                    st.success("Referência importada.")
                    st.rerun()
        elif st.button("Reimportar referência padrão"):
            if importar_planilha_referencia(PLANILHA_REFERENCIA, substituir=True):
                st.success("Referência padrão importada.")
                st.rerun()
    with col2:
        arquivos = st.file_uploader("Planilhas de alunos", type=["xlsx"], accept_multiple_files=True)
        if arquivos and st.button("Importar execuções dos alunos", type="primary"):
            ok = 0
            for arq in arquivos:
                destino = BASE_DIR / arq.name
                destino.write_bytes(arq.getbuffer())
                ok += int(importar_execucoes_ciclo(destino))
            st.success(f"Planilhas processadas: {ok}.")
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
                st.success("Senha alterada.")

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
                        st.success("Senha redefinida.")
                        st.rerun()
                ativo_atual = int(usuarios.loc[usuarios["id"] == usuario_id, "ativo"].iloc[0])
                col4, col5 = st.columns(2)
                if ativo_atual == 1 and col4.button("Bloquear usuário"):
                    executar("UPDATE alunos SET ativo = 0 WHERE id = ?", (int(usuario_id),))
                    st.success("Usuário bloqueado.")
                    st.rerun()
                if ativo_atual == 0 and col5.button("Reativar usuário"):
                    executar("UPDATE alunos SET ativo = 1, force_troca_senha = 1 WHERE id = ?", (int(usuario_id),))
                    st.success("Usuário reativado com troca de senha obrigatória.")
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
