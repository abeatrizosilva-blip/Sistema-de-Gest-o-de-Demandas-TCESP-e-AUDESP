# Sistema de Gestao de Demandas TCE-SP e AUDESP

## Como executar

O sistema requer Python 3.10 ou mais recente.

### Linux, macOS ou terminal do VS Code

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

Abra http://localhost:5000 no navegador. Na primeira execucao, crie o usuario administrador na tela de configuracao inicial.

### Windows

```bat
py -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python app.py
```

Tambem e possivel executar `iniciar_sistema.bat` depois de criar o ambiente virtual e instalar as dependencias.

O banco de dados principal e a planilha `sp_aguas.xlsx`, criada automaticamente na primeira execucao. Ela possui as abas `usuarios`, `demandas` e `historico`. Para usar outro caminho, defina `EXCEL_DATABASE` antes de iniciar:

```bash
EXCEL_DATABASE=/caminho/dados.xlsx python app.py
```

No Windows, se o arquivo `C:\Users\ana.silva\OneDrive - PRODESP\Banco de Dados - Sistema.xlsx` existir, o aplicativo tambem o detecta automaticamente mesmo quando iniciado diretamente com `python app.py`.

O sistema carrega os dados em memoria durante cada operacao e salva a planilha a cada `commit`, substituindo o arquivo de forma atomica. Se `sp_aguas.xlsx` ainda nao existir e `sp_aguas.db` estiver presente, os dados do SQLite sao migrados automaticamente na primeira inicializacao. Depois disso, o arquivo Excel passa a ser a fonte principal.

### OneDrive / Excel Online

Para manter a planilha persistente fora do servidor, o sistema pode sincronizar o arquivo com o OneDrive usando a Microsoft Graph API. Essa configuracao requer uma aplicacao registrada no Microsoft Entra ID com permissao de aplicacao `Files.ReadWrite.All` e consentimento administrativo. Ela e indicada para OneDrive corporativo ou SharePoint; OneDrive pessoal requer um fluxo OAuth delegado.

Defina as variaveis abaixo no ambiente de execucao, sem coloca-las no Git:

```bash
ONEDRIVE_ENABLED=true
ONEDRIVE_TENANT_ID=seu-tenant-id
ONEDRIVE_CLIENT_ID=seu-client-id
ONEDRIVE_CLIENT_SECRET=seu-client-secret
ONEDRIVE_USER=usuario@empresa.gov.br
ONEDRIVE_PATH=SP_AGUAS/sp_aguas.xlsx
```

Ao iniciar, o sistema baixa `ONEDRIVE_PATH`; a cada cadastro, edicao ou exclusao, envia a planilha atualizada de volta ao OneDrive. Se a planilha remota ainda nao existir, o primeiro salvamento cria o arquivo.

### Vercel

O arquivo `vercel.json` configura o Flask como uma funcao Python. Na Vercel, a planilha usa `/tmp` e pode ser apagada entre execucoes. Portanto, essa configuracao e adequada para testes locais ou ambientes com armazenamento persistente; para producao com varios usuarios, recomenda-se um banco externo ou um servico de planilhas com controle de concorrencia.