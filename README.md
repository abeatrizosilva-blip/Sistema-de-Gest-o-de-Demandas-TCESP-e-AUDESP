# SP ÁGUAS — Sistema de Gestão de Demandas TCESP e AUDESP

## Persistência em uma única planilha Excel

Esta versão foi preparada para uso **local no computador que possui o OneDrive sincronizado**.

A planilha utilizada pelo sistema é:

```text
C:\Users\ana.silva\OneDrive - PRODESP\SP_AGUAS\Sistema de Gestão de Demandas - SP Aguas.xlsx
```

O sistema grava diretamente nesse arquivo as alterações realizadas pela aplicação. Não há Power Automate, Microsoft Graph, SharePoint API ou banco SQLite em arquivo.

O SQLite que aparece internamente no código é **somente memória RAM**, usado como mecanismo temporário para manter as consultas existentes. Ele é recriado a partir do Excel ao iniciar a aplicação e depois de cada alteração. Nenhum arquivo `.db` é criado ou usado para persistência.

### Dados persistidos no Excel

- `usuarios`: cadastro, aprovação, ativação/bloqueio e demais alterações de usuários.
- `demandas`: cadastro, edição e exclusão de demandas, incluindo ETC e informações do DOE-TCESP.
- `historico`: movimentações e alterações registradas pelo sistema.

### Requisito

O computador precisa estar com o OneDrive sincronizado e a planilha disponível nesse caminho.

Se o arquivo estiver aberto no Excel, o sistema poderá não conseguir substituí-lo em algumas situações. Recomenda-se fechar a planilha durante o uso do sistema.

### Iniciar

Execute `iniciar_sistema.bat`. O arquivo já aponta para o caminho da planilha acima.

### DOE-TCESP

Na tela Nova Demanda, é possível inserir o link de uma publicação/PDF ou selecionar um PDF local. O sistema extrai o texto, procura as palavras-chave cadastradas e permite usar a ocorrência encontrada para preencher os dados da demanda.

### Dependências

- Flask
- openpyxl
- bcrypt
- pypdf
- requests
