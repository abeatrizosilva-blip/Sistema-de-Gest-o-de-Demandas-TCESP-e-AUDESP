# SP ÁGUAS — Sistema de Gestão de Demandas TCESP e AUDESP

## Versão limpa — ETC + DOE-TCESP

Aplicação Flask para cadastro e acompanhamento de demandas, com suporte a número ETC e leitura de publicações do DOE-TCESP.

### Persistência

A versão atual **não utiliza Power Automate, Office Scripts, Microsoft Graph ou API do SharePoint**.

Os dados são mantidos no SQLite local (`SQLITE_DATABASE`) e o sistema mantém uma cópia em Excel (`EXCEL_DATABASE`). As alterações feitas pela aplicação são gravadas no Excel e o SQLite é reconstruído para permanecer sincronizado.

> Em Vercel, o sistema de arquivos da função não é um banco persistente. Para produção com persistência entre deploys/execuções, será necessário conectar um banco externo (por exemplo, PostgreSQL/Supabase) em uma etapa posterior.

### DOE-TCESP

Na tela **Nova demanda**, o usuário pode:

- colar o link direto de um PDF oficial do DOE-TCESP; ou
- selecionar um PDF salvo no computador.

Quando o PDF é selecionado no computador, a leitura ocorre **diretamente no navegador com PDF.js**. Isso evita o limite de tamanho de requisição do Vercel e elimina o erro `Request Entity Too Large`.

Quando é informado um link, o servidor baixa o PDF somente de domínios oficiais do TCESP e extrai o texto com `pypdf`.

O sistema procura as palavras-chave cadastradas, informa página e trecho, identifica números de processo quando encontrados e permite **Usar esta ocorrência** para preencher a demanda.

### Dependências

- Flask
- openpyxl
- bcrypt
- pypdf
- requests

### Variáveis de ambiente

```text
SECRET_KEY=UMA_CHAVE_FORTE
EXCEL_DATABASE=sp_aguas.xlsx
SQLITE_DATABASE=sp_aguas.db
```

### Arquivos removidos da arquitetura

Não fazem mais parte desta versão:

- Power Automate
- Office Script
- endpoint de integração Power Automate
- credenciais/segredos Microsoft Graph
- rotinas de sincronização com SharePoint/OneDrive

### Funcionalidades preservadas

- Login e cadastro de usuários
- Aprovação de usuários
- Dashboard
- Cadastro/edição/exclusão de demandas
- Número de processo
- Número ETC correspondente
- Prazos e alertas
- Histórico
- TCE-SP e AUDESP
- Pesquisa de demandas
- Exportação para Excel
- Leitura de DOE-TCESP por PDF/link
