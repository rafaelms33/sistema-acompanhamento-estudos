"""
Exporta todos os dados do PostgreSQL de produção (Neon) para producao.json.

Uso:
    python exportar_pg.py

A URL de conexão é lida, nesta ordem:
  1. variável de ambiente DATABASE_URL
  2. arquivo neon_url.txt na mesma pasta (uma linha, só a URL)

Manter a URL fora deste arquivo evita versionar a senha do banco por acidente.
Coloque `neon_url.txt` no .gitignore.

Para definir a variável no PowerShell:
    $env:DATABASE_URL = "postgresql://usuario:senha@host/banco?sslmode=require"

Requer:
    pip install "psycopg[binary]"      # a mesma biblioteca usada pelo app
"""

import json
import os
import sys
from pathlib import Path

# Usa a mesma biblioteca do app (psycopg v3, já em requirements.txt).
# Mantém psycopg2 como alternativa para ambientes antigos.
DRIVER = None
try:
    import psycopg
    from psycopg.rows import dict_row
    DRIVER = "psycopg3"
except ImportError:
    try:
        import psycopg2
        import psycopg2.extras
        DRIVER = "psycopg2"
    except ImportError:
        print("Falta a biblioteca de conexão com PostgreSQL. Instale com:")
        print("    .\\.venv\\Scripts\\pip.exe install \"psycopg[binary]\"")
        sys.exit(1)


def conectar_pg(url):
    """Abre a conexão e devolve um cursor que entrega dicionários."""
    if DRIVER == "psycopg3":
        conexao = psycopg.connect(url)
        return conexao, conexao.cursor(row_factory=dict_row)
    conexao = psycopg2.connect(url)
    return conexao, conexao.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

TABELAS = [
    "alunos", "disciplinas", "aulas", "assuntos",
    "tarefas", "execucoes", "sessoes_estudo",
]

SAIDA = Path("producao.json")
ARQUIVO_URL = Path("neon_url.txt")


def obter_url():
    url = os.environ.get("DATABASE_URL")
    if url:
        return url.strip(), "variável DATABASE_URL"
    if ARQUIVO_URL.exists():
        conteudo = ARQUIVO_URL.read_text(encoding="utf-8").strip()
        if conteudo:
            return conteudo, f"arquivo {ARQUIVO_URL}"
    return None, None


def main():
    url, origem = obter_url()
    if not url:
        print("URL de conexão não encontrada.\n")
        print("Escolha uma das opções:")
        print("  1. Defina a variável de ambiente:")
        print('     $env:DATABASE_URL = "postgresql://..."')
        print(f"  2. Crie o arquivo {ARQUIVO_URL} com a URL numa única linha")
        return 1

    print(f"Conectando com {DRIVER} (URL lida da {origem})…")
    try:
        conexao, cursor = conectar_pg(url)
    except Exception as exc:
        print(f"Falha ao conectar: {exc}")
        return 1

    dados = {}
    for tabela in TABELAS:
        try:
            cursor.execute(f"SELECT * FROM {tabela}")
            dados[tabela] = [dict(linha) for linha in cursor.fetchall()]
            print(f"  {tabela}: {len(dados[tabela])} registros")
        except Exception as exc:
            primeira_linha = str(exc).strip().splitlines()[0]
            print(f"  {tabela}: não exportada — {primeira_linha}")
            conexao.rollback()

    conexao.close()

    if not dados:
        print("\nNenhuma tabela exportada. Nada foi gravado.")
        return 1

    SAIDA.write_text(
        json.dumps(dados, ensure_ascii=False, default=str),
        encoding="utf-8",
    )
    total = sum(len(v) for v in dados.values())
    print(f"\n{total} registros salvos em {SAIDA}")
    print("\nPróximo passo:  python montar_db.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
