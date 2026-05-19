# Deploy - Sistema de Acompanhamento de Estudos

Este projeto está pronto para rodar localmente com SQLite e em produção com PostgreSQL.

## Variáveis de ambiente

Use `.env.example` como modelo. Em produção configure:

```env
APP_ENV=production
DEBUG=0
DATABASE_URL=postgresql://usuario:senha@host:porta/banco
ESTUDOS_ADMIN_EMAIL=seu-email-admin
ESTUDOS_ADMIN_PASSWORD=uma-senha-temporaria-forte
```

Não coloque credenciais reais no código nem no repositório.

## Executar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se `DATABASE_URL` estiver vazio, a aplicação usa SQLite localmente.

## Criar PostgreSQL gratuito

Opções compatíveis:

- Render PostgreSQL, quando disponível na conta.
- Railway PostgreSQL.
- Neon PostgreSQL gratuito.
- Supabase PostgreSQL gratuito.

Após criar o banco, copie a connection string no formato:

```txt
postgresql://usuario:senha@host:porta/banco
```

Use essa string em `DATABASE_URL`.

## Deploy no Render

1. Suba o projeto para um repositório Git.
2. Crie um PostgreSQL gratuito ou use Neon/Supabase/Railway.
3. Crie um novo Web Service no Render.
4. Configure:
   - Build command: `pip install -r requirements.txt`
   - Start command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0 --server.headless true`
5. Adicione as variáveis de ambiente.
6. Faça o deploy.
7. Acesse a URL HTTPS gerada pelo Render.

## Deploy no Railway

1. Crie um projeto Railway.
2. Adicione um banco PostgreSQL.
3. Adicione o serviço da aplicação a partir do repositório.
4. Configure `DATABASE_URL`, `APP_ENV=production`, `DEBUG=0`, `ESTUDOS_ADMIN_EMAIL` e `ESTUDOS_ADMIN_PASSWORD`.
5. O `Procfile` já informa o comando de inicialização.

## Deploy no Streamlit Cloud

O Streamlit Cloud pode rodar a aplicação, mas não fornece PostgreSQL próprio. Use Neon ou Supabase para o banco.

1. Crie o banco PostgreSQL externo.
2. Publique o repositório no GitHub.
3. Crie o app no Streamlit Cloud apontando para `app.py`.
4. Em Secrets, configure:

```toml
APP_ENV = "production"
DEBUG = "0"
DATABASE_URL = "postgresql://usuario:senha@host:porta/banco"
ESTUDOS_ADMIN_EMAIL = "seu-email-admin"
ESTUDOS_ADMIN_PASSWORD = "senha-temporaria-forte"
```

## Segurança

- O primeiro login de usuários novos exige troca de senha.
- Senhas são armazenadas com hash PBKDF2.
- Alunos não acessam importações, cadastros administrativos nem manutenção.
- Gestores podem redefinir senhas, forçar troca no próximo login, bloquear e reativar usuários.
- Em produção, mantenha `DEBUG=0`.

## HTTPS

Render, Railway e Streamlit Cloud fornecem HTTPS automaticamente. A aplicação não depende de URLs fixas `http://`.

## Backup manual PostgreSQL

Use `pg_dump`:

```bash
pg_dump "$DATABASE_URL" > backup_estudos.sql
```

Restaurar:

```bash
psql "$DATABASE_URL" < backup_estudos.sql
```

## Backup automático gratuito

Estratégias simples:

- Agendar um workflow no GitHub Actions que rode `pg_dump` e salve o arquivo como artifact.
- Usar a rotina de backup nativa do provedor, quando disponível.
- Fazer exportação CSV dentro da aplicação em `Configurações > Backup`.

Exemplo de cron para servidor Linux:

```bash
0 3 * * * pg_dump "$DATABASE_URL" > "/backups/estudos_$(date +\%F).sql"
```

## Teste após deploy

1. Abrir a URL pública HTTPS.
2. Entrar com o gestor.
3. Trocar a senha obrigatória.
4. Criar ou importar alunos.
5. Confirmar que aluno não vê telas administrativas.
6. Registrar uma atividade.
7. Conferir dashboard, comparação entre alunos e backup CSV.

## Checklist final

- [ ] `APP_ENV=production`
- [ ] `DEBUG=0`
- [ ] `DATABASE_URL` configurado
- [ ] senha temporária do admin forte
- [ ] primeiro login força troca de senha
- [ ] alunos criados com troca obrigatória
- [ ] importações visíveis apenas para gestor
- [ ] backup testado
- [ ] URL HTTPS funcionando
