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

O sistema carrega os dados em memoria durante cada operacao e salva a planilha a cada `commit`, substituindo o arquivo de forma atomica. Se `sp_aguas.xlsx` ainda nao existir e `sp_aguas.db` estiver presente, os dados do SQLite sao migrados automaticamente na primeira inicializacao. Depois disso, o arquivo Excel passa a ser a fonte principal.

### Vercel

O arquivo `vercel.json` configura o Flask como uma funcao Python. Na Vercel, a planilha usa `/tmp` e pode ser apagada entre execucoes. Portanto, essa configuracao e adequada para testes locais ou ambientes com armazenamento persistente; para producao com varios usuarios, recomenda-se um banco externo ou um servico de planilhas com controle de concorrencia.