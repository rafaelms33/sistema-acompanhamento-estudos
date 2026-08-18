"""
Reconstrói o SQLite local (trilha_tjs_ajaa.db) a partir de producao.json.

Uso:
    python montar_db.py

Apaga e recria o banco do zero. Já inclui as constraints UNIQUE que o
importador da planilha exige nas cláusulas ON CONFLICT — sem elas a
importação falha com "Importação cancelada".

Atenção: o banco gerado contém os hashes de senha reais das alunas.
Mantenha `producao.json` e `*.db` no .gitignore.
"""

import json
import os
import sqlite3
import sys
from pathlib import Path

ENTRADA = Path("producao.json")
BANCO = Path("trilha_tjs_ajaa.db")

ESQUEMA = """
CREATE TABLE alunos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL UNIQUE,
    email TEXT UNIQUE,
    senha TEXT NOT NULL,
    perfil TEXT NOT NULL DEFAULT 'Aluno',
    ativo INTEGER DEFAULT 1,
    force_troca_senha INTEGER DEFAULT 0
);
CREATE TABLE disciplinas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL UNIQUE,
    ativo INTEGER DEFAULT 1
);
CREATE TABLE aulas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    disciplina_id INTEGER,
    aula TEXT,
    assunto TEXT,
    estudada_padrao TEXT DEFAULT 'Não',
    revisao_24h_padrao TEXT DEFAULT 'Não',
    ativo INTEGER DEFAULT 1,
    tipo_estudo TEXT DEFAULT 'Outro',
    UNIQUE(disciplina_id, aula)
);
CREATE TABLE assuntos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    aula_id INTEGER,
    titulo TEXT,
    ativo INTEGER DEFAULT 1,
    UNIQUE(aula_id, titulo)
);
CREATE TABLE tarefas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero INTEGER UNIQUE,
    trilha INTEGER,
    disciplina_id INTEGER,
    seq_disciplina INTEGER,
    aula TEXT,
    qtd_exercicios_previstos INTEGER DEFAULT 0,
    tipo TEXT,
    conteudo TEXT,
    ativo INTEGER DEFAULT 1,
    aula_id INTEGER,
    assunto_id INTEGER
);
CREATE TABLE execucoes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    aluno_id INTEGER,
    tarefa_id INTEGER,
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
);
CREATE TABLE sessoes_estudo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    aluno_id INTEGER NOT NULL,
    tarefa_id INTEGER NOT NULL,
    data_sessao TEXT NOT NULL,
    ch_sessao REAL DEFAULT 0,
    qtd_questoes INTEGER DEFAULT 0,
    qtd_acertos INTEGER DEFAULT 0,
    tipo_estudo TEXT DEFAULT 'Outro',
    comentario TEXT,
    criado_em TEXT DEFAULT CURRENT_TIMESTAMP
);
"""

COLUNAS = {
    "alunos": ["id", "nome", "email", "senha", "perfil", "ativo",
               "force_troca_senha"],
    "disciplinas": ["id", "nome", "ativo"],
    "aulas": ["id", "disciplina_id", "aula", "assunto", "estudada_padrao",
              "revisao_24h_padrao", "ativo", "tipo_estudo"],
    "assuntos": ["id", "aula_id", "titulo", "ativo"],
    "tarefas": ["id", "numero", "trilha", "disciplina_id", "seq_disciplina",
                "aula", "qtd_exercicios_previstos", "tipo", "conteudo",
                "ativo", "aula_id", "assunto_id"],
    "execucoes": ["id", "aluno_id", "tarefa_id", "data_execucao", "ch_efetiva",
                  "data_revisao_24h", "ch_revisao", "qtd_acertos", "desempenho",
                  "comentario", "concluida", "atualizado_em",
                  "qtd_questoes_feitas", "status", "tipo_estudo"],
    "sessoes_estudo": ["id", "aluno_id", "tarefa_id", "data_sessao",
                       "ch_sessao", "qtd_questoes", "qtd_acertos",
                       "tipo_estudo", "comentario", "criado_em"],
}


def main():
    if not ENTRADA.exists():
        print(f"{ENTRADA} não encontrado. Rode antes:  python exportar_pg.py")
        return 1

    dados = json.loads(ENTRADA.read_text(encoding="utf-8"))

    if BANCO.exists():
        reserva = BANCO.with_name(BANCO.stem + "_backup.db")
        if reserva.exists():
            reserva.unlink()
        BANCO.rename(reserva)
        print(f"Banco anterior preservado em {reserva}")

    conexao = sqlite3.connect(BANCO)
    conexao.execute("PRAGMA foreign_keys = OFF")
    conexao.executescript(ESQUEMA)

    total = 0
    for tabela, colunas in COLUNAS.items():
        linhas = dados.get(tabela) or []
        if not linhas:
            print(f"  {tabela}: 0")
            continue
        marcadores = ",".join("?" for _ in colunas)
        sql = (f"INSERT OR REPLACE INTO {tabela} ({','.join(colunas)}) "
               f"VALUES ({marcadores})")
        conexao.executemany(
            sql, [[linha.get(c) for c in colunas] for linha in linhas]
        )
        print(f"  {tabela}: {len(linhas)}")
        total += len(linhas)

    conexao.commit()

    historico = conexao.execute("""
        SELECT COUNT(*), ROUND(COALESCE(SUM(ch_efetiva), 0), 2),
               COALESCE(SUM(qtd_questoes_feitas), 0)
        FROM execucoes
        WHERE ch_efetiva > 0 OR qtd_questoes_feitas > 0
           OR status <> 'NAO_INICIADA'
    """).fetchone()
    conexao.close()

    print(f"\n{total} registros gravados em {BANCO}")
    print(f"Histórico de estudo: {historico[0]} execuções · "
          f"{historico[1]}h · {historico[2]} questões")
    print("\nO banco já vem com as constraints UNIQUE que o importador exige.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
