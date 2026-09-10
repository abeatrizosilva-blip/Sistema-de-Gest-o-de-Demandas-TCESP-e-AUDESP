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

O banco SQLite e criado automaticamente como `sp_aguas.db`. Para usar outro arquivo local, defina a variavel de ambiente `DATABASE` antes de iniciar.

### Vercel

O arquivo `vercel.json` configura o Flask como uma funcao Python. Com `DATABASE_URL` configurada, a Vercel usa o Neon PostgreSQL e os dados ficam persistentes. Sem essa variavel, o aplicativo usa SQLite em `/tmp` apenas para testes; esse armazenamento e temporario e pode ser apagado entre execucoes.

### Neon PostgreSQL

O aplicativo usa Neon quando a variavel `DATABASE_URL` estiver configurada. Na Vercel, adicione essa variavel com a connection string do Neon e publique novamente. Nunca coloque essa string diretamente no codigo ou no Git.