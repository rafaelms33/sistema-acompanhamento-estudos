"""
Adiciona as constraints UNIQUE que o importador exige (cláusulas ON CONFLICT).
Rode uma vez sobre um banco SQLite já montado. É idempotente e não altera dados.
"""
import sqlite3

DB = "trilha_tjs_ajaa.db"

conn = sqlite3.connect(DB)
try:
    conn.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS ux_assuntos_aula_titulo "
        "ON assuntos(aula_id, titulo)"
    )
    conn.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS ux_aulas_disc_aula "
        "ON aulas(disciplina_id, aula)"
    )
    conn.commit()
    print("Constraints aplicadas com sucesso.")
    for (nome,) in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='index' AND name LIKE 'ux_%'"
    ):
        print("  -", nome)
except sqlite3.IntegrityError as exc:
    conn.rollback()
    print(f"Falhou: existem duplicatas que impedem a constraint.\n  {exc}")
finally:
    conn.close()