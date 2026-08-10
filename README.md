# Gestor Financeiro de Condomínios — Web

Versão web do Gestor Financeiro de Condomínios v3. Recebe qualquer planilha Excel
de previsão orçamentária, detecta automaticamente a estrutura (grupos numerados,
unidades, taxa atual, fundo de reserva, garantidora) e gera uma **apresentação
PowerPoint premium** no padrão do cliente, pronta para assembleia.

## Arquitetura

```
┌─────────────────────┐        ┌──────────────────────────┐
│  Frontend (Netlify) │  HTTP  │   Backend (Render)       │
│  HTML + CSS + JS    │ ─────► │   Flask API              │
│  + supabase-js      │        │   • condominio_app_v2.py │
└─────────┬───────────┘        │   • powerpoint_generator │
          │                    └────────────┬─────────────┘
          │  Auth + histórico (RLS)         │ valida JWT + grava histórico
          ▼                                 ▼
┌──────────────────────────────────────────────────────────┐
│  Supabase (projeto dedicado)                             │
│  • Auth (e-mail/senha, sem signup público)               │
│  • Tabela relatorios (RLS por usuário)                   │
└──────────────────────────────────────────────────────────┘
```

- **Frontend**: página estática com login (Supabase Auth), upload da planilha,
  pré-visualização da análise e histórico de relatórios gerados.
- **Backend**: API Flask que reutiliza integralmente o analisador e o gerador
  de PPTX da versão desktop. Valida o token do Supabase em cada requisição e
  grava o histórico após cada geração. O `.pptx` é gerado em memória
  (`BytesIO`) — nada é gravado em disco.

## Endpoints da API

| Método | Rota              | Descrição                                  |
|--------|-------------------|--------------------------------------------|
| GET    | `/api/health`     | Ping / aquecimento do servidor             |
| POST   | `/api/analisar`   | Recebe a planilha, devolve JSON da análise |
| POST   | `/api/gerar-pptx` | Recebe a planilha, devolve o `.pptx`       |

Campos do `multipart/form-data` (POST): `planilha` (obrigatório, .xls/.xlsx),
`nome_condominio`, `data_assembleia`, `taxa_atual` (opcionais).

## Deploy

### 1. Supabase (autenticação + histórico)

1. Crie um **projeto Supabase novo e dedicado** a este cliente.
2. No **SQL Editor**, execute o arquivo `supabase/schema.sql` (cria a tabela
   `relatorios` com RLS por usuário).
3. Em **Authentication → Sign In / Up**, desmarque *"Allow new users to sign up"*
   (não há cadastro público — os acessos são criados pela administração).
4. Em **Authentication → Users → Add user**, crie o(s) acesso(s) do cliente
   (e-mail + senha, com *Auto Confirm User* marcado).
5. Anote em **Settings → API**: a *Project URL* e a chave **publishable**
   (`sb_publishable_...`). A chave publishable é pública por design — o RLS
   protege os dados. **Nunca use a service_role/secret key.**

### 2. Backend no Render

1. Suba este repositório para o GitHub.
2. No Render: **New → Blueprint** e aponte para o repositório
   (o `render.yaml` na raiz configura tudo), **ou** crie um Web Service manual:
   - Root Directory: `backend`
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app --timeout 120`
3. Defina as variáveis de ambiente:
   - `SUPABASE_URL` → Project URL do passo 1
   - `SUPABASE_ANON_KEY` → chave publishable do passo 1
4. Anote a URL gerada (ex.: `https://gestor-condominio-api.onrender.com`).

> Sem `SUPABASE_URL` configurada, a API roda em **modo aberto** (sem login) —
> útil apenas para desenvolvimento local. Em produção, configure sempre.

### 3. Frontend no Netlify

1. Edite as três constantes no topo de `frontend/app.js`:
   `API_URL` (Render), `SUPABASE_URL` e `SUPABASE_KEY` (publishable).
2. Faça commit e, no Netlify: **Add new site → Import from Git** apontando
   para o repositório (o `netlify.toml` já define `publish = "frontend"`).

### 4. Travar o CORS (recomendado)

Após o deploy do Netlify, defina no Render a variável de ambiente:

```
CORS_ORIGINS=https://SEUSITE.netlify.app
```

## Desenvolvimento local

```bash
# Backend
cd backend
pip install -r requirements.txt
python app.py                     # http://localhost:5000

# Frontend (em outro terminal)
cd frontend
python -m http.server 8000        # http://localhost:8000
```

`API_URL` em `frontend/app.js` já vem apontando para `http://localhost:5000`.

## Formatos de planilha aceitos

1. **Export Condomínio21 / BalAnual** (formato real do cliente): colunas
   `Rubrica/Despesa | Previsão AAAA Aprovada | 01/AAAA..12/AAAA | Total |
   Realizado Média | Previsão AAAA+1 | Índice`. Detectado automaticamente —
   extrai unidades, fundo de reserva e índice de reajuste do próprio arquivo.
2. **Formato legado** (grupos numerados "1. Encargos Bancários"...): continua
   suportado.

Campos opcionais do formulário: nome do condomínio, data da assembleia,
taxa atual, nº de unidades (sobrepõe o detectado), inadimplência total e
acordos a receber (geram o slide de inadimplência).

## Observações

- **Plano gratuito do Render**: o servidor hiberna após inatividade; a primeira
  requisição pode levar até ~1 minuto. O frontend avisa o usuário e faz um ping
  de aquecimento ao carregar a página.
- **Limite de upload**: 10 MB (configurável em `backend/app.py`).
- **Formatos aceitos**: `.xlsx` (openpyxl) e `.xls` legado (xlrd).
- A versão desktop (Tkinter/`gui_app.py`) segue funcionando de forma
  independente — este repositório contém apenas a versão web.
