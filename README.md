# SP ÁGUAS — Sistema de Gestão de Demandas

## Versão 4 — Power Automate

Esta versão mantém a modernização visual da versão 3, mas muda a arquitetura de persistência:

`Vercel/Flask → Power Automate → Excel Online (Business) → SharePoint/OneDrive`

O Vercel **não acessa diretamente o Excel** e não usa credenciais Microsoft Graph.

### O que foi preservado

- Dashboard moderno com indicadores e próximos prazos.
- Navegação lateral institucional e responsiva.
- Formulário de demanda organizado por seções.
- Badges de status e alertas de prazo.
- Busca/filtro de demandas.
- Histórico da demanda.
- Login, cadastro e aprovação de usuários.
- Exclusão e edição de demandas.
- Diagnósticos somente leitura.

### O que mudou

- Removidas as dependências de `TENANT_ID`, `CLIENT_ID` e `CLIENT_SECRET` no Vercel.
- O Vercel envia apenas JSON ao fluxo Power Automate.
- O Power Automate é responsável por acessar o arquivo corporativo.
- Um Office Script lê/substitui os dados das abas `usuarios`, `demandas` e `historico`.
- O fluxo deve controlar concorrência para evitar gravações simultâneas.

### Arquivos importantes

- `app.py` — aplicação Flask.
- `power_automate/OfficeScript_SP_AGUAS.ts` — script para o Excel.
- `power_automate/FLUXO_POWER_AUTOMATE.md` — passo a passo para criar o fluxo.
- `power_automate/REQUEST_SCHEMA.json` — esquema do POST recebido pelo fluxo.

### Variáveis do Vercel

```text
POWER_AUTOMATE_ENABLED=true
POWER_AUTOMATE_URL=https://SEU-ENDPOINT-DO-POWER-AUTOMATE
POWER_AUTOMATE_SECRET=UM_SEGREDO_FORTE_E_ALEATORIO
POWER_AUTOMATE_TIMEOUT=90
SECRET_KEY=UMA_CHAVE_FORTE_DO_FLASK
```

Não configure `SHAREPOINT_TENANT_ID`, `SHAREPOINT_CLIENT_ID` ou `SHAREPOINT_CLIENT_SECRET` nesta versão.

### Observação

O conector Excel Online (Business) suporta arquivos em OneDrive for Business e SharePoint. A Microsoft também documenta limitações para gravações concorrentes no mesmo workbook; por isso o fluxo deve ser configurado com concorrência controlada.
