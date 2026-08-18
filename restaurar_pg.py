"""
Restaura o PostgreSQL a partir de um producao.json gerado pelo exportar_pg.py.

É o caminho de volta do backup em JSON, para quando `pg_dump` não está instalado.
Um dump do pg_dump continua sendo melhor (guarda schema, índices e permissões);
este script assume que as tabelas já existem e repõe apenas os DADOS.

Uso:
    python restaurar_pg.py --dry-run          # só mostra o que faria
    python restaurar_pg.py                    # restaura, pedindo confirmação
    python restaurar_pg.py outro_backup.json

A URL de conexão é lida, nesta ordem:
  1. variável de ambiente DATABASE_URL
  2. arquivo neon_url.txt na mesma pasta

OPERAÇÃO DESTRUTIVA: apaga o conteúdo atual das sete tabelas e regrava a partir
do arquivo. Tudo roda dentro de UMA transação — se qualquer passo falhar, nada é
alterado. Exige digitar RESTAURAR para confirmar.

Requer:
    pip install "psycopg[binary]"
"""

import json
import os
import sys
from pathlib import Path

DRIVER = None
try:
    import psycopg
    DRIVER = "psycopg3"
except ImportError:
    try:
        import psycopg2
        DRIVER = "psycopg2"
    except ImportError:
        print("Falta a biblioteca de conexão com PostgreSQL. Instale com:")
        print('    .\\.venv\\Scripts\\pip.exe install "psycopg[binary]"')
        sys.exit(1)

ARQUIVO_URL = Path("neon_url.txt")

# Ordem importa: pais antes dos filhos, por causa das chaves estrangeiras.
# A restauração insere nesta ordem e apaga na ordem inversa.
COLUNAS = {
    "disciplinas": ["id", "nome", "ativo"],
    "alunos": ["id", "nome", "email", "senha", "perfil", "ativo",
               "force_troca_senha"],
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


def obter_url():
    url = os.environ.get("DATABASE_URL")
    if url:
        return url.strip(), "variável DATABASE_URL"
    if ARQUIVO_URL.exists():
        conteudo = ARQUIVO_URL.read_text(encoding="utf-8").strip()
        if conteudo:
            return conteudo, f"arquivo {ARQUIVO_URL}"
    return None, None


def conectar(url):
    if DRIVER == "psycopg3":
        return psycopg.connect(url)
    return psycopg2.connect(url)


def resumo_historico(cur):
    """Números do histórico de estudo, para comparar antes e depois."""
    try:
        cur.execute("""
            SELECT COUNT(*), COALESCE(ROUND(SUM(ch_efetiva)::numeric, 2), 0),
                   COALESCE(SUM(qtd_questoes_feitas), 0),
                   COALESCE(SUM(qtd_acertos), 0)
            FROM execucoes
            WHERE ch_efetiva > 0 OR qtd_questoes_feitas > 0
               OR status <> 'NAO_INICIADA'
        """)
        exec_ = cur.fetchone()
        cur.execute("SELECT COUNT(*) FROM sessoes_estudo")
        sess = cur.fetchone()[0]
        return (int(exec_[0]), float(exec_[1]), int(exec_[2]), int(exec_[3]), int(sess))
    except Exception:
        return None


def resumo_json(dados):
    execs = dados.get("execucoes") or []
    com_estudo = [e for e in execs
                  if (e.get("ch_efetiva") or 0) > 0
                  or (e.get("qtd_questoes_feitas") or 0) > 0
                  or (e.get("status") or "NAO_INICIADA") != "NAO_INICIADA"]
    horas = round(sum(float(e.get("ch_efetiva") or 0) for e in com_estudo), 2)
    quest = sum(int(e.get("qtd_questoes_feitas") or 0) for e in com_estudo)
    acert = sum(int(e.get("qtd_acertos") or 0) for e in com_estudo)
    return (len(com_estudo), horas, quest, acert, len(dados.get("sessoes_estudo") or []))


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dry = "--dry-run" in sys.argv
    entrada = Path(args[0]) if args else Path("producao.json")

    if not entrada.exists():
        print(f"Arquivo não encontrado: {entrada}")
        print("Gere um backup antes com:  python exportar_pg.py")
        return 1

    dados = json.loads(entrada.read_text(encoding="utf-8"))
    faltando = [t for t in COLUNAS if t not in dados]
    if faltando:
        print(f"O arquivo não tem estas tabelas: {', '.join(faltando)}")
        print("Ele não parece ter vindo do exportar_pg.py.")
        return 1

    print(f"Arquivo:  {entrada}")
    for tabela in COLUNAS:
        print(f"  {tabela:<18} {len(dados.get(tabela) or []):>6} registro(s)")
    h = resumo_json(dados)
    print(f"\nHistórico no arquivo: {h[0]} execuções · {h[1]}h · "
          f"{h[2]} questões · {h[3]} acertos · {h[4]} sessões")

    url, origem = obter_url()
    if not url:
        print("\nURL de conexão não encontrada.")
        print('  Defina:  $env:DATABASE_URL = "postgresql://..."')
        print(f"  ou crie o arquivo {ARQUIVO_URL} com a URL numa única linha")
        return 1

    print(f"\nDestino: PostgreSQL (URL lida da {origem})")
    try:
        conexao = conectar(url)
    except Exception as exc:
        print(f"Falha ao conectar: {exc}")
        return 1

    cur = conexao.cursor()
    antes = resumo_historico(cur)
    if antes:
        print(f"Histórico HOJE no banco: {antes[0]} execuções · {antes[1]}h · "
              f"{antes[2]} questões · {antes[3]} acertos · {antes[4]} sessões")

    if dry:
        print("\n--dry-run: nada foi alterado.")
        conexao.close()
        return 0

    print("\n" + "=" * 66)
    print(" Isto APAGA o conteúdo atual das sete tabelas e regrava pelo arquivo.")
    print(" Tudo roda numa transação: se algo falhar, nada muda.")
    print("=" * 66)
    if input(" Digite RESTAURAR para confirmar: ").strip() != "RESTAURAR":
        print("Cancelado. Nada foi alterado.")
        conexao.close()
        return 1

    try:
        # Apaga na ordem inversa (filhos antes dos pais)
        for tabela in reversed(list(COLUNAS)):
            cur.execute(f"DELETE FROM {tabela}")

        total = 0
        for tabela, colunas in COLUNAS.items():
            linhas = dados.get(tabela) or []
            if not linhas:
                print(f"  {tabela}: 0")
                continue
            marcadores = ",".join(["%s"] * len(colunas))
            sql = (f"INSERT INTO {tabela} ({','.join(colunas)}) "
                   f"VALUES ({marcadores})")
            cur.executemany(
                sql, [[linha.get(c) for c in colunas] for linha in linhas]
            )
            print(f"  {tabela}: {len(linhas)}")
            total += len(linhas)

        # Reposiciona as sequências, senão o próximo INSERT colide com id existente
        for tabela in COLUNAS:
            cur.execute(
                f"SELECT setval(pg_get_serial_sequence('{tabela}', 'id'), "
                f"COALESCE((SELECT MAX(id) FROM {tabela}), 1))"
            )

        depois = resumo_historico(cur)
        esperado = resumo_json(dados)
        if depois != esperado:
            conexao.rollback()
            conexao.close()
            print("\nATENÇÃO: o histórico gravado não bate com o do arquivo.")
            print(f"  esperado: {esperado}")
            print(f"  no banco: {depois}")
            print("  Nada foi alterado (rollback).")
            return 1

        conexao.commit()
    except Exception as exc:
        conexao.rollback()
        conexao.close()
        print(f"\nFalhou: {exc}")
        print("Nada foi alterado (rollback).")
        return 1

    conexao.close()
    print(f"\n{total} registros restaurados.")
    print(f"Histórico conferido: {esperado[0]} execuções · {esperado[1]}h · "
          f"{esperado[2]} questões · {esperado[3]} acertos · {esperado[4]} sessões")
    return 0


if __name__ == "__main__":
    sys.exit(main())
