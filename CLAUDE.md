# Sistema de Acompanhamento de Estudos — TJs AJAA

Aplicação Streamlit que acompanha o progresso de duas alunas (Lilian e Jessica)
numa trilha de estudos para concursos de Tribunais de Justiça.

## Ambiente

- **Código:** `app.py` (Streamlit)
- **Banco local:** `trilha_tjs_ajaa.db` (SQLite)
- **Banco de produção:** PostgreSQL no Neon, alcançado pela variável `DATABASE_URL`
- **Deploy:** Render, disparado por `git push origin main`
- **Python:** use sempre `.venv\Scripts\python.exe` (Windows/PowerShell)

O código detecta o ambiente pela presença de `DATABASE_URL`. Por isso **todo SQL
precisa funcionar nos dois bancos**.

## Regras que não podem ser quebradas

### Histórico de estudo é intocável

As tabelas `execucoes` e `sessoes_estudo` guardam o que as alunas efetivamente
estudaram. Nenhum script pode alterar linhas que tenham:

- `status` diferente de `NAO_INICIADA`, ou
- `ch_efetiva > 0`, ou
- `qtd_questoes_feitas > 0`, ou
- qualquer registro em `sessoes_estudo`

Ao escrever qualquer `UPDATE` nessas tabelas, inclua as quatro condições no
`WHERE` e confira o histórico antes e depois para provar que não mudou.

### SQL compatível com SQLite e PostgreSQL

- `ROUND(x, n)` exige `CAST(x AS NUMERIC)` — `REAL` falha no PostgreSQL
- `ORDER BY` no PostgreSQL não aceita apelido de coluna dentro de expressão;
  repita a expressão inteira
- Evite `NULLS LAST`; use `CASE WHEN ... IS NULL THEN 1 ELSE 0 END`
- `ON CONFLICT` exige constraint `UNIQUE` correspondente nos dois bancos

### Planilha de referência

A planilha vem do fornecedor com numeração sequencial. Os simulados são
acréscimo nosso, inseridos a cada 10 tarefas a partir da 50.

**Gere sempre a partir da planilha original do fornecedor, nunca de uma que já
contenha simulados.** Rodar duas vezes sobre o mesmo arquivo empilharia
simulados e deslocaria toda a numeração.

O importador casa as linhas pelo número da tarefa. Se um número passar a
descrever outro conteúdo, o histórico preso a ele fica atribuído à tarefa
errada — em silêncio, sem erro nenhum.

## Documentos de apoio

- `HISTORICO_TECNICO.md` — decisões, bugs já diagnosticados e armadilhas.
  **Leia antes de mexer em SQL, na planilha ou na numeração de tarefas.**
- `GUIA_AMBIENTE_LOCAL.md` — como montar o ambiente e resolver problemas comuns.

## Scripts

| Script | O que faz |
|---|---|
| `exportar_pg.py` | Baixa produção para `producao.json` (precisa de rede) |
| | URL lida de `DATABASE_URL` ou de `neon_url.txt` |
| `montar_db.py` | Reconstrói o SQLite a partir do `producao.json` |
| `corrigir_schema.py` | Cria as constraints `UNIQUE` que o importador exige |
| `inserir_simulados.py` | Insere simulados na planilha do fornecedor |
| `verificar_impacto.py` | Confere se algum estudo seria remapeado |
| `realinhar_tipos.py` | Corrige `tipo_estudo` defasado após renumeração |

## Rotina a cada atualização da planilha

```powershell
.\.venv\Scripts\python.exe montar_db.py
.\.venv\Scripts\python.exe corrigir_schema.py
.\.venv\Scripts\python.exe inserir_simulados.py <planilha_do_fornecedor>.xlsx
.\.venv\Scripts\python.exe verificar_impacto.py <planilha>_com_simulados.xlsx
```

O `verificar_impacto.py` termina em SEGURO ou ATENÇÃO. **Se der ATENÇÃO, pare** e
me avise antes de importar qualquer coisa.

Depois de importar pelo app:

```powershell
.\.venv\Scripts\python.exe realinhar_tipos.py --dry-run
.\.venv\Scripts\python.exe realinhar_tipos.py
```

## O que fazer fora do Cowork

- **`exportar_pg.py`** — precisa alcançar o Neon. Se o egress bloquear, eu rodo
  no PowerShell e você trabalha a partir do `producao.json` gerado.
- **`streamlit run app.py`** — a VM é isolada, o `localhost:8501` dela não abre
  no meu navegador. Teste visual eu faço por fora.
- **`git push`** — deploy e credenciais ficam comigo.

## Como quero que você trabalhe

- Antes de mexer em SQL que toque `execucoes` ou `sessoes_estudo`, mostre o
  `WHERE` e explique por que o histórico está protegido.
- Depois de alterar `app.py`, valide a sintaxe com `python -c "import ast;
  ast.parse(open('app.py').read())"`.
- Ao editar a planilha, rode o recálculo e me diga se sobrou algum erro de
  fórmula.
- Antes de qualquer operação destrutiva num banco, faça uma cópia com sufixo
  `_backup` na mesma pasta.
- Nunca versione `producao.json` nem `*.db` — contêm hashes de senha reais.

## Dados sensíveis

`producao.json` e `trilha_tjs_ajaa.db` trazem os hashes de senha das alunas.
Mantenha os dois no `.gitignore` e não os copie para fora da pasta do projeto.
