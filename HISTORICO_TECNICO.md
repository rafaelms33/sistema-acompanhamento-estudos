# Histórico técnico — decisões e armadilhas

Registro do que foi descoberto ao longo do desenvolvimento. Serve para não
repetir diagnósticos já feitos e para entender **por que** o código está do
jeito que está.

---

## 1. Renumeração de tarefas e histórico de estudo

### O mecanismo

O vínculo entre estudo e tarefa é o `tarefa_id`. O importador casa as linhas da
planilha pelo `numero` (`ON CONFLICT(numero) DO UPDATE`) e sobrescreve
disciplina, tipo e conteúdo de quem tem aquele número.

Se a numeração muda, o número passa a descrever outra coisa — mas o histórico
grudado nele **não se move junto**. O dano é silencioso: nada quebra, nenhum
erro aparece, os totais de horas continuam certos. Só a atribuição por
disciplina fica errada.

### Por que os simulados são seguros

A posição de cada simulado é contada a partir do início da lista (50, 60, 70…).
Enquanto o fornecedor apenas **acrescentar** tarefas no fim, o número final de
toda tarefa existente permanece o mesmo. Verificado:

| Tarefas do fornecedor | Total final | Tarefa 100 vira |
|---|---|---|
| 384 | 422 | 106 |
| 420 | 462 | 106 |
| 500 | 551 | 106 |

Teste feito de ponta a ponta: com histórico gravado nos simulados das tarefas
50 e 70, uma segunda importação com 42 simulados preservou tudo — horas,
questões, acertos, comentários e sessões, com o mesmo `id` de tarefa.

### O que quebraria

- fornecedor inserir ou reordenar tarefas no meio da lista
- mudar `PRIMEIRO_SIMULADO` ou `INTERVALO` no `inserir_simulados.py`
- rodar o gerador sobre uma planilha que já contém simulados

Por isso existe o `verificar_impacto.py`. Rode sempre antes de importar.

---

## 2. `tipo_estudo` defasado após renumeração

### O sintoma

Os simulados apareciam como "Revisão" em vez de "Exercícios", e toda a
sequência de tarefas depois da 50 mostrava o tipo trocado.

### A causa

O tipo fica gravado em **dois lugares**: `tarefas.tipo` (o que a planilha diz) e
`execucoes.tipo_estudo` (o que o aluno escolheu ao estudar). Antes dos
simulados, a tarefa 50 era "Português — Revisão", e a execução guardou
`tipo_estudo = 'Revisão'`. Depois da renumeração a tarefa virou Simulado, mas a
execução continuou dizendo "Revisão" — e era ela que vencia na exibição.

Atingiu 702 execuções.

### A correção

A consulta decide pelo status, não por `COALESCE` simples:

```sql
CASE WHEN COALESCE(e.status, 'NAO_INICIADA') = 'NAO_INICIADA'
     THEN COALESCE(t.tipo, 'Outro')
     ELSE COALESCE(NULLIF(e.tipo_estudo, 'Outro'), t.tipo, 'Outro')
END AS tipo
```

Tarefa não iniciada mostra o tipo da tarefa; tarefa estudada mostra o que o
aluno escolheu. Um valor defasado nunca mais aparece.

> Um `COALESCE(e.tipo_estudo, ...)` puro **não funciona**: a coluna tem
> `DEFAULT 'Outro'`, nunca é `NULL`, então o `COALESCE` sempre para nela e o
> tipo da tarefa jamais aparece. Foi o primeiro diagnóstico, e estava incompleto.

O `realinhar_tipos.py` corrige os dados já gravados.

---

## 3. Ordenação por atividade real, não por `atualizado_em`

`atualizado_em` é carimbado sempre que a linha é tocada — inclusive por uma
importação. Ordenar por ele fazia tarefas nunca estudadas subirem ao topo da
lista de atividades recentes.

A ordenação agora usa a data da sessão de estudo mais recente, com
`data_execucao` como reserva, e joga para o fim quem não tem atividade nenhuma.

---

## 4. Incompatibilidades SQLite × PostgreSQL

O código roda em SQLite local e PostgreSQL em produção. Estas já causaram erro:

| Problema | Sintoma | Solução |
|---|---|---|
| `ROUND(x, n)` com `REAL` | `function round(double precision, integer) does not exist` | `CAST(x AS NUMERIC)` |
| Apelido de coluna em expressão no `ORDER BY` | falha no PostgreSQL | repetir a expressão inteira |
| `NULLS LAST` | não confiável em SQLite antigo | `CASE WHEN ... IS NULL THEN 1 ELSE 0 END` |
| `ON CONFLICT` sem constraint | "Importação cancelada" | criar o índice `UNIQUE` |

O SQLite é mais permissivo. **Um SQL que funciona local pode quebrar em
produção** — foi exatamente o caso do `ROUND`.

---

## 5. Constraints que o importador exige

O importador usa `ON CONFLICT(aula_id, titulo)` e `ON CONFLICT(disciplina_id,
aula)`. Sem as constraints correspondentes o SQLite recusa a cláusula e a
importação falha com a mensagem genérica "Importação cancelada", que esconde a
exceção real.

Para ver a causa real de qualquer falha de importação:

```powershell
$env:DEBUG = "1"
.\.venv\Scripts\python.exe -m streamlit run app.py
```

O `montar_db.py` já cria as constraints. O `corrigir_schema.py` conserta bancos
montados antes dessa correção.

---

## 6. Defeitos da exportação do Google Sheets

A planilha vem do Google. A exportação para `.xlsx` deixa quatro problemas:

| Defeito | Efeito no Excel | Correção |
|---|---|---|
| `INDIRECT` com intervalo aberto (`B$3:B`) | retorna 1 em vez da contagem | fechar em `B$3:B$500` |
| Nome de aba com espaço, sem aspas | referência não resolve | envolver em `'…'` |
| Coluna `AL` com nome completo, aba truncada em 31 caracteres | `INDIRECT` não acha a aba | `AL` guarda o nome real da aba |
| `QUERY()` e `SPARKLINE()` | viram `#NAME?` ao salvar | substituir pelo valor em cache |

Os dois primeiros são **erros silenciosos** — não aparece `#ERRO`, só um número
errado. O `inserir_simulados.py` corrige os quatro a cada execução.

Consequência da última correção: as colunas "Aulas abordadas" e "Total de
Tarefas" das disciplinas antigas viraram números fixos e não se atualizam
sozinhas. As linhas de Simulados usam `COUNTIF`, que é fórmula de verdade.

Divergência pré-existente: o total de "Total de Tarefas" dava 376 contra 384
tarefas reais. O `QUERY` do Google já contava 8 a menos. Não foi introduzida
por nós.

---

## 7. Registro Rápido — formulário unificado

Antes eram dois formulários separados ("Atualizar status" e "Registrar sessão"),
cada um com seu botão. Dava para atualizar o status sem registrar a sessão, e a
ordem das operações confundia.

Agora é um formulário único com um botão "Salvar":

- o status é sempre atualizado
- a sessão só é registrada se houver horas > 0 ou questões > 0
- com os campos zerados, só o status muda

---

## 8. Dados sensíveis

`producao.json` e `trilha_tjs_ajaa.db` contêm os **hashes de senha reais** das
alunas. As credenciais de produção funcionam no ambiente local justamente por
isso.

Mantenha no `.gitignore`:

```
.venv/
producao.json
neon_url.txt
*.db
__pycache__/
```

A URL do Neon inclui a senha do banco. Por isso o `exportar_pg.py` a lê de
variável de ambiente ou de `neon_url.txt`, em vez de tê-la escrita no código.

---

## 9. Estado atual dos dados

- 2 alunas: Lilian e Jessica
- Histórico real concentrado nas tarefas 1 a 26
- 44 execuções com estudo · 152,92h · 3055 questões · 2229 acertos
- 38 simulados nas posições 50, 60, 70 … 420
- 422 tarefas no total (384 do fornecedor + 38 simulados)

O histórico não passar da tarefa 26 é o que torna a renumeração inofensiva
hoje. Quando as alunas passarem da 50, qualquer mudança de cadência dos
simulados vira remapeamento — aí o `verificar_impacto.py` deixa de ser
formalidade.
