# SP ÁGUAS — Sistema de Gestão de Demandas

Versão 3: modernização visual e reforço da segurança do fluxo de sincronização.

## Principais alterações
- Dashboard redesenhado com indicadores e próximos prazos.
- Navegação lateral institucional e responsiva.
- Formulário de demanda organizado por seções.
- Badges de status e alertas de prazo.
- Tela de demandas com busca/filtro e tabela responsiva.
- Histórico da demanda em `/historico/<id>`.
- Login e cadastro com layout modernizado.
- `/diagnostico-sync` transformado em diagnóstico **somente leitura**: não executa upload nem sobrescreve o Excel.
- Mantida a integração Microsoft Graph/OneDrive e os fluxos atômicos existentes.

## Deploy
Suba os arquivos para o projeto Vercel e configure as variáveis do `.env.example` no painel do Vercel.

A integração Graph exige credenciais e permissões administrativas fornecidas pela TI/administrador do Microsoft Entra.
