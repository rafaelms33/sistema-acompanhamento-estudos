"""
Realinha o `tipo_estudo` das execuções que ficaram defasadas.

Quando as tarefas são renumeradas (inserção de simulados, por exemplo), a
coluna `execucoes.tipo_estudo` continua com o valor do conteúdo ANTIGO daquele
número. A tarefa 50 vira Simulado, mas a execução ainda diz "Revisão", que era
o tipo do Português que ocupava aquela posição.

Este script copia o tipo da tarefa para a execução, mas SOMENTE onde não há
nenhum estudo registrado:

    status = NAO_INICIADA
    e nenhuma hora lançada
    e nenhuma questão lançada
    e nenhuma sessão de estudo gravada

Qualquer execução com histórico fica intacta — o tipo escolhido pelo aluno na
hora de estudar é dele, não da planilha.

Uso:
    python realinhar_tipos.py [banco.db]
    python realinhar_tipos.py --dry-run          # só mostra, não grava
"""

import sys
import sqlite3
from pathlib import Path

CONDICAO = """
    status = 'NAO_INICIADA'
    AND COALESCE(ch_efetiva, 0) = 0
    AND COALESCE(qtd_questoes_feitas, 0) = 0
    AND NOT EXISTS (
        SELECT 1 FROM sessoes_estudo s
        WHERE s.aluno_id = execucoes.aluno_id
          AND s.tarefa_id = execucoes.tarefa_id
    )
"""


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dry = "--dry-run" in sys.argv
    banco = Path(args[0]) if args else Path("trilha_tjs_ajaa.db")

    if not banco.exists():
        print(f"Banco não encontrado: {banco}")
        return 1

    conn = sqlite3.connect(banco)

    # Fotografa o histórico antes, para provar que nada nele mudou
    def foto_historico():
        return conn.execute("""
            SELECT e.aluno_id, e.tarefa_id, e.tipo_estudo, e.status,
                   e.ch_efetiva, e.qtd_questoes_feitas, e.qtd_acertos
            FROM execucoes e
            WHERE e.ch_efetiva > 0 OR e.qtd_questoes_feitas > 0
               OR e.status <> 'NAO_INICIADA'
            ORDER BY e.aluno_id, e.tarefa_id
        """).fetchall()

    antes = foto_historico()

    defasadas = conn.execute(f"""
        SELECT t.numero, d.nome, execucoes.tipo_estudo, COALESCE(t.tipo, 'Outro')
        FROM execucoes
        JOIN tarefas t ON t.id = execucoes.tarefa_id
        JOIN disciplinas d ON d.id = t.disciplina_id
        WHERE {CONDICAO}
          AND COALESCE(execucoes.tipo_estudo, 'Outro') <> COALESCE(t.tipo, 'Outro')
        ORDER BY t.numero
    """).fetchall()

    protegidas = conn.execute(f"""
        SELECT COUNT(*) FROM execucoes WHERE NOT ({CONDICAO})
    """).fetchone()[0]

    print(f"Banco: {banco}")
    print(f"  Execuções a realinhar:        {len(defasadas)}")
    print(f"  Execuções com histórico:      {protegidas}  (não serão tocadas)")

    if not defasadas:
        print("\nNada a fazer — todos os tipos já estão alinhados.")
        return 0

    print("\n  Amostra das mudanças:")
    vistos = set()
    for numero, disciplina, de, para in defasadas:
        if numero in vistos:
            continue
        vistos.add(numero)
        print(f"    tarefa {numero:<4} {disciplina[:28]:<28} {de or '—':<11} -> {para}")
        if len(vistos) >= 8:
            break

    if dry:
        print("\n--dry-run: nada foi gravado.")
        return 0

    conn.execute(f"""
        UPDATE execucoes SET tipo_estudo = (
            SELECT COALESCE(t.tipo, 'Outro') FROM tarefas t
            WHERE t.id = execucoes.tarefa_id
        )
        WHERE {CONDICAO}
    """)
    conn.commit()

    depois = foto_historico()
    print(f"\n  Realinhadas: {len(defasadas)} execuções")
    if antes == depois:
        print("  Histórico conferido: nenhuma execução com estudo foi alterada.")
    else:
        print("  ATENÇÃO: o histórico mudou — isso não deveria acontecer.")
        conn.close()
        return 1

    conn.close()
    print("\nPronto.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
