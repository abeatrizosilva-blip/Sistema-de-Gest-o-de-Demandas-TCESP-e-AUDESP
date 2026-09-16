# SP ÁGUAS — Excel Local v12

Versão corrigida para planilha Excel no OneDrive local.

## Correção desta versão

Corrigido o erro `StopIteration` que ocorria na inicialização quando uma das abas `usuarios`, `demandas` ou `historico` estava vazia.

O sistema agora:
- cria automaticamente as abas ausentes;
- cria os cabeçalhos quando a aba está vazia;
- acrescenta colunas novas quando a planilha é de uma versão anterior;
- preserva os dados existentes;
- usa somente a planilha configurada como persistência permanente.

Planilha configurada:
`C:\Users\ana.silva\OneDrive - PRODESP\SP_AGUAS\Sistema de Gestão de Demandas - SP Aguas.xlsx`
