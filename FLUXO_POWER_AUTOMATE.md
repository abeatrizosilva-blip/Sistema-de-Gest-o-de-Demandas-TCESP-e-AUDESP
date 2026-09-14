# SP ÁGUAS — fluxo Power Automate

## Arquitetura

`Vercel (Flask) → HTTP POST → Power Automate → Excel Online (Business) / Office Script → Excel no SharePoint/OneDrive`

O Vercel **não** acessa Microsoft Graph, não guarda Tenant ID, Client ID ou Client Secret e não envia o XLSX diretamente para o SharePoint.

O Excel Online (Business) trabalha com arquivos armazenados em OneDrive for Business e Sites do SharePoint. O conector oferece ações de leitura/escrita e também a ação de executar Office Script. Consulte a documentação oficial da Microsoft para detalhes e limitações.

## 1. Preparar o Excel

Mantenha o arquivo:

`SP_AGUAS/Sistema de Gestão de Demandas - SP Aguas.xlsx`

Ele deve conter as três abas com estes cabeçalhos na primeira linha:

- `usuarios`: `id`, `nome`, `usuario`, `email`, `senha_hash`, `perfil`, `ativo`, `aprovado`, `criado_em`
- `demandas`: `id`, `numero_processo`, `origem`, `assunto`, `area`, `responsavel`, `data_recebimento`, `prazo_area`, `prazo_fatal`, `situacao`, `prioridade`, `observacoes`, `criado_por`, `criado_em`, `atualizado_em`
- `historico`: `id`, `demanda_id`, `usuario_id`, `acao`, `descricao`, `data_hora`

Não é obrigatório transformar as abas em Tabelas do Excel nesta arquitetura, porque o Office Script trabalha diretamente nas planilhas.

## 2. Criar o Office Script

No Excel para a Web, abra o arquivo, vá em **Automatizar → Novo Script**, apague o conteúdo e cole:

`OfficeScript_SP_AGUAS.ts`

Salve com o nome, por exemplo:

`SP_AGUAS_API`

O script devolve o snapshot em `GET_SNAPSHOT` e substitui os dados em `REPLACE_SNAPSHOT`.

## 3. Criar o fluxo

Crie um fluxo de nuvem no Power Automate.

### Gatilho

Use o gatilho HTTP para receber um `POST` do Vercel:

**Quando uma solicitação HTTP é recebida**

No esquema JSON, use o conteúdo de `REQUEST_SCHEMA.json`.

> Observação: a disponibilidade/licença do gatilho HTTP pode depender do ambiente da organização. Se o Power Automate indicar que o fluxo é premium, a TI/administrador deverá providenciar a licença adequada.

### Segurança do fluxo

O Vercel envia o segredo no cabeçalho:

`X-SP-AGUAS-SECRET`

Configure uma variável/valor seguro no fluxo com o mesmo segredo e faça uma condição logo após o gatilho:

`triggerOutputs()?['headers']?['x-sp-aguas-secret']`

Se não coincidir, responda HTTP `401` e encerre o fluxo.

Não coloque o segredo na URL pública do fluxo.

### Ação 1 — Executar Script

Adicione:

**Excel Online (Business) → Executar script**

Selecione:

- Localização: OneDrive for Business ou o site do SharePoint onde está o arquivo;
- Arquivo: `SP_AGUAS/Sistema de Gestão de Demandas - SP Aguas.xlsx`;
- Script: `SP_AGUAS_API`;
- `action`: conteúdo dinâmico `action` recebido do gatilho;
- `payloadJson`: use a expressão:

```text
string(triggerBody())
```

O corpo completo é enviado ao Office Script porque o `REPLACE_SNAPSHOT` precisa receber também `expectedVersion` para impedir sobrescrita de uma versão mais nova.

### Ação 2 — Responder a uma solicitação HTTP

Para `GET_SNAPSHOT`, o resultado do script já é o JSON que o Vercel precisa.

Para simplificar o fluxo, você pode usar o mesmo retorno para as duas ações:

- Status Code: `200`
- Header `Content-Type`: `application/json`
- Body: resultado retornado pelo **Executar script**.

O Flask aceita tanto o objeto JSON direto quanto um campo `result` contendo o JSON serializado.

## 4. Recomendações para o fluxo

Ative **controle de concorrência** do gatilho com grau de paralelismo `1`. O Office Script também mantém uma versão interna na aba oculta `_sp_aguas_meta`; se o arquivo mudar entre a leitura e a gravação, o fluxo devolve `412` e o Vercel tenta novamente.

Também recomendamos não editar o mesmo arquivo simultaneamente pelo Excel Desktop enquanto o sistema estiver gravando. A Microsoft documenta limitações de gravações concorrentes no Excel Online (Business).

## 5. Variáveis no Vercel

Configure somente:

```text
POWER_AUTOMATE_ENABLED=true
POWER_AUTOMATE_URL=https://SEU-ENDPOINT-DO-POWER-AUTOMATE
POWER_AUTOMATE_SECRET=UM_SEGREDO_FORTE_E_ALEATORIO
POWER_AUTOMATE_TIMEOUT=90
SECRET_KEY=UMA_CHAVE_FORTE_DO_FLASK
```

Você **não precisa mais** configurar:

- `SHAREPOINT_TENANT_ID`
- `SHAREPOINT_CLIENT_ID`
- `SHAREPOINT_CLIENT_SECRET`
- `SHAREPOINT_SITE_ID`
- `SHAREPOINT_DRIVE_ID`

## 6. Teste

Depois do Deploy no Vercel:

1. entre como administrador;
2. abra `/teste-integracao`;
3. o resultado esperado é:

`Vercel → Power Automate → Excel está funcionando.`

A rota `/diagnostico-sync` continua somente leitura.

## 7. Importante sobre licenciamento

O conector **Excel Online (Business)** é listado pela Microsoft como conector padrão do Power Automate. Já conectores premium/HTTP podem depender da licença do ambiente. Confirme no tenant da organização antes de publicar o fluxo em produção.
