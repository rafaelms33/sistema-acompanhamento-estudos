# Sistema de Acompanhamento de Estudos

Aplicação Streamlit para acompanhamento de estudos, gestão de tarefas educacionais, registro de produtividade, dashboards inteligentes e comparação entre alunos.

## Rodar localmente

```powershell
pip install -r requirements.txt
streamlit run app.py
```

Sem `DATABASE_URL`, o sistema usa SQLite local em `trilha_tjs_ajaa.db`.

Usuário inicial:

- E-mail: `admin@admin.com`
- Senha: `123`

Em novos ambientes, o usuário criado inicialmente deve alterar a senha no primeiro acesso.

## Produção

Para produção gratuita, use PostgreSQL configurando:

```env
APP_ENV=production
DEBUG=0
DATABASE_URL=postgresql://usuario:senha@host:porta/banco
ESTUDOS_ADMIN_EMAIL=seu-email-admin
ESTUDOS_ADMIN_PASSWORD=senha-temporaria-forte
```

Veja o passo a passo completo em [README_DEPLOY.md](README_DEPLOY.md).
