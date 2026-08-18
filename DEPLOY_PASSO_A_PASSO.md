# Subir em produção

Cinco passos. Nenhum campo para preencher — todos os comandos são para copiar e
colar como estão.

---

## 1. Backup

**Onde:** navegador, `https://console.neon.tech/`

1. Entre no projeto
2. Menu lateral → **Backup & restore**
3. Botão **Create snapshot**

Aparece na lista como "Manual snapshot". Pronto.

> No plano gratuito só cabe 1 snapshot manual. Se já existir um antigo, apague
> antes de criar o novo.

---

## 2. Conferir a constraint

**Onde:** navegador, mesmo site → menu lateral → **SQL Editor**

Cole isto e clique em **Run**:

```sql
SELECT c.conname AS constraint_nome,
       string_agg(a.attname, ', ' ORDER BY k.ord) AS colunas
FROM pg_constraint c
JOIN LATERAL unnest(c.conkey) WITH ORDINALITY AS k(attnum, ord) ON TRUE
JOIN pg_attribute a ON a.attrelid = c.conrelid AND a.attnum = k.attnum
WHERE c.conrelid = 'execucoes'::regclass AND c.contype IN ('u','p')
GROUP BY c.conname;
```

Tem que aparecer uma linha com `aluno_id, tarefa_id` na coluna `colunas`.

**Se aparecer** → vá para o passo 3.

**Se não aparecer** → rode isto antes de seguir:

```sql
ALTER TABLE execucoes ADD CONSTRAINT ux_exec_aluno_tarefa UNIQUE (aluno_id, tarefa_id);
```

---

## 3. Enviar o código

**Onde:** PowerShell (Menu Iniciar → digite "PowerShell" → Enter)

Cole os cinco blocos abaixo, um de cada vez.

**3.1 — ir para a pasta**

```powershell
cd C:\Users\Rafael.000\Documents\Codex\2026-05-19\import-os-import-re-import-sqlite3
```

**3.2 — marcar os arquivos**

```powershell
git add app.py .gitignore .env.example CLAUDE.md HISTORICO_TECNICO.md DEPLOY_PASSO_A_PASSO.md exportar_pg.py restaurar_pg.py montar_db.py corrigir_schema.py realinhar_tipos.py verificar_impacto.py inserir_simulados.py
git rm --cached --ignore-unmatch .gitignore.txt trilha_tjs_ajaa.db estudos_concursos.db app_teste.py
```

O `--ignore-unmatch` faz o comando não dar erro se algum desses arquivos já tiver
sido tratado antes. Pode rodar quantas vezes quiser.

**3.3 — conferir**

```powershell
git diff --cached --name-only
```

Leia a lista. Ela deve ter só arquivos `.py` e `.md` mais `.gitignore` e
`.env.example`.

**Se aparecer `GUIA_AMBIENTE_LOCAL.md`, `producao.json` ou qualquer `.db`**, tire
antes de continuar:

```powershell
git restore --staged GUIA_AMBIENTE_LOCAL.md producao.json
```

**3.4 — registrar**

```powershell
git commit -m "fix: corrige gravacao de sessoes, metricas por periodo e visualizacao"
```

**3.5 — enviar**

```powershell
git push origin main
```

Se pedir login, use seu usuário do GitHub e um *personal access token* no lugar
da senha.

---

## 4. Acompanhar

**Onde:** navegador, `https://dashboard.render.com/`

1. Clique no seu Web Service
2. Aba **Events** — o deploy aparece em segundos, com o texto do seu commit
3. Espere de 2 a 5 minutos até virar **Live**

Se der **Deploy failed**, a versão antiga continua no ar. Abra a aba **Logs**,
leia o erro, corrija e repita 3.4 e 3.5.

---

## 5. Testar

**Onde:** navegador, na URL pública do app (está no topo da página do Render)

1. Entre como gestor. Se a tela parecer estranha, aperte `Ctrl + Shift + R`
2. **Registro rápido** → escolha uma tarefa → registre 0,25h e 1 questão → Salvar
3. A sessão tem que aparecer no histórico logo abaixo, na hora
4. Vá no Dashboard e veja se o total subiu 15 minutos
5. Volte no Registro rápido e **exclua** a sessão de teste

Se o passo 3 funcionar, o deploy está bom.

---

## Se der errado

**Onde:** Render → seu serviço → aba **Events**

Ache o deploy anterior, clique em **Rollback**, confirme em **Rollback to this
deploy**. Volta ao ar em cerca de 1 minuto.

⚠️ **O rollback desliga o deploy automático.** Depois de corrigir o problema,
reative em **Settings** → **Auto-Deploy**. Se esquecer, seus próximos `git push`
não publicam nada.
