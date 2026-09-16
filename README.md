# SP ÁGUAS — versão 13 (servidor na rede local)

Esta versão mantém a **planilha Excel do OneDrive como única persistência** e permite que outra pessoa, na mesma rede local, use o sistema pelo navegador.

## Como funciona

- O computador da Ana executa o sistema e acessa diretamente:
  `C:\Users\ana.silva\OneDrive - PRODESP\SP_AGUAS\Sistema de Gestão de Demandas - SP Aguas.xlsx`
- O Flask/Waitress fica ouvindo a porta 5000 na rede local.
- O segundo usuário acessa pelo navegador usando o IPv4 do computador da Ana:
  `http://IP-DO-COMPUTADOR:5000`
- O segundo usuário **não precisa instalar Python nem receber o ZIP**.
- O login do próprio sistema continua sendo usado.

## Início

1. No computador que possui o OneDrive e a planilha, extraia esta pasta.
2. Execute `iniciar_sistema.bat`.
3. Aguarde a mensagem **SISTEMA DISPONÍVEL NA REDE LOCAL**.
4. O próprio arquivo mostrará os endereços IPv4.
5. No outro computador, abra o navegador e digite, por exemplo:
   `http://192.168.1.25:5000`

## Se o outro computador não conseguir acessar

No computador servidor:

1. Feche o sistema.
2. Clique com o botão direito em `liberar_porta_5000.bat`.
3. Escolha **Executar como administrador**.
4. Execute novamente `iniciar_sistema.bat`.

Se a empresa bloquear alterações no Firewall, será necessário pedir ao TI a liberação da porta TCP 5000 na rede interna.

## Regras importantes

- O computador servidor precisa ficar ligado e com `iniciar_sistema.bat` aberto.
- A outra pessoa não deve abrir a planilha diretamente no Excel para editar os dados enquanto o sistema estiver em uso.
- O sistema usa um bloqueio interno para serializar as gravações feitas pelos usuários que acessam o servidor.
- Se o OneDrive ou o Excel bloquear o arquivo durante uma gravação, a operação poderá precisar ser repetida após o arquivo ficar disponível.
- Esta solução é para **rede interna**. Não publique a porta 5000 diretamente na internet.

## Vercel

Esta versão não depende do Vercel para acessar o arquivo local. O Vercel continua sem acesso ao caminho `C:\Users\...` do computador servidor.
