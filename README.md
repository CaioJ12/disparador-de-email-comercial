# Disparador de E-mail Comercial

Ferramenta local para enviar e-mails comerciais personalizados, um a um, a partir de uma planilha de empresas. Cada contato precisa ser revisado e aprovado antes do envio.

Tudo roda no seu computador: o painel abre no navegador, mas não fica acessível pela internet.

## O que ele faz

- Lê uma planilha de empresas e separa e-mails, telefones e nomes, mesmo quando estão misturados na mesma célula.
- Sugere a saudação de cada contato, como `Olá, Maria,` ou `Olá, equipe de Compras da Empresa Exemplo,`.
- Bloqueia o envio até cada contato ser aprovado.
- Mostra a prévia do e-mail de cada contato.
- Inclui faixa do catálogo com botão de download, botão de WhatsApp e assinatura.
- Envia pelo Outlook clássico (sem senha) ou por SMTP.
- Copia a equipe comercial em Cc ou Cco.
- Espera de 8 a 10 segundos entre os e-mails e envia até 30 por execução.
- Para sozinho depois de 3 erros seguidos.
- Registra os envios e não repete quem já recebeu.

## Requisitos

- Windows 10 ou 11
- Python 3.10 ou mais novo, com a opção **Add python.exe to PATH** marcada na instalação
- Para enviar pelo Outlook: Outlook clássico configurado (o "Novo Outlook" não funciona)

## Instalação

1. Baixe o projeto em **Code > Download ZIP** e extraia a pasta.
2. Abra o Prompt de Comando dentro da pasta e instale a dependência:

   ```bash
   pip install openpyxl
   ```

3. Se for enviar pelo Outlook, instale também:

   ```bash
   pip install pywin32
   ```

4. Dê dois cliques em `ABRIR_PAINEL.bat`. O painel abre em `http://127.0.0.1:8765`. Deixe a janela preta aberta enquanto usar.

Na primeira execução, o `config.exemplo.txt` é copiado para `config.txt`, onde ficam os seus dados.

## Passo a passo

### 1. Configure o remetente

Na aba **Configurações**, preencha empresa, nome, cargo, e-mail, telefone, WhatsApp e endereço. Use nos campos de remetente, resposta e assinatura o mesmo e-mail da conta que faz o envio.

Para mostrar sua logo, coloque um arquivo chamado `logo.png` na pasta do projeto. Ela é opcional.

### 2. Escolha como enviar

| Modo | Quando usar | Senha |
|---|---|---|
| `OUTLOOK` | Outlook clássico instalado neste computador | Não pede |
| `SMTP` | Gmail, Hostinger, Locaweb, cPanel ou outro servidor | Pedida a cada envio, não fica gravada |

No Gmail, use uma senha de app. Se o Microsoft 365 recusar o login por senha, use o modo `OUTLOOK`. O botão **Testar conexão** ajuda a descobrir o servidor e a porta.

### 3. Monte a planilha de empresas

Crie um arquivo `.xlsx` com uma coluna chamada `EMPRESA` ou `RAZÃO SOCIAL`. O cabeçalho deve estar nas primeiras 15 linhas. As demais colunas podem ter qualquer nome:

| EMPRESA | CONTATO | OBSERVAÇÕES |
|---|---|---|
| Empresa Exemplo S.A. | compras@exemplo.com.br; (11) 3000-0000 | Falar com Maria |
| Grupo Modelo Ltda. | maria.silva@modelo.com.br | |

Importe a planilha pelo painel e gere a lista de revisão. O resultado fica em `revisao_envio.xlsx`.

### 4. Revise os contatos

Na aba **Contatos & Revisão**, confira cada linha:

| Campo | O que significa |
|---|---|
| `TIPO` | `PESSOAL` para uma pessoa, `GRUPO` para um setor |
| `NOME` | Nome da pessoa ou do setor |
| `SAUDACAO` | Primeira linha exata do e-mail |
| `FRASE_PERSONALIZADA` | Frase opcional só para aquela empresa |
| `ASSUNTO` | Assunto opcional, no lugar do padrão |
| `ENVIAR` | `NAO` pula o contato |
| `APROVADO` | `SIM` libera o contato. Sem isso, o envio fica bloqueado |

### 5. Escreva o e-mail

Na aba **Texto do e-mail**, troque todos os campos `[entre colchetes]` do modelo. O envio fica bloqueado enquanto algum deles estiver no texto.

**Texto modelo:**

> Olá, Maria,
>
> Meu nome é **[Seu nome]**, sou [seu cargo] da **[Sua Empresa]**, empresa de [sua cidade – UF] que atende [sua região ou segmento].
>
> Gostaria de apresentar o nosso **catálogo de produtos** e me colocar à disposição como fornecedor da Empresa Exemplo. Nele estão as principais linhas que oferecemos:
>
> - **[Linha de produto 1]:** [breve descrição]
> - **[Linha de produto 2]:** [breve descrição]
> - **[Linha de produto 3]:** [breve descrição]
>
> Trabalhamos com atendimento próximo e retorno rápido de cotações.
>
> Se fizer sentido, você pode me enviar uma lista dos itens que costumam comprar para que eu prepare uma cotação — ou, se preferir, é só me chamar no WhatsApp.
>
> Fico à disposição e agradeço a atenção.
>
> Atenciosamente,

**Campos automáticos:**

| Campo | Vira |
|---|---|
| `{saudacao}` | A saudação revisada do contato |
| `{nome}` | Nome da pessoa ou do setor |
| `{empresa_curta}` | Nome da empresa sem S.A. ou Ltda. |
| `{a_empresa}` / `{da_empresa}` | "a Empresa Exemplo" / "da Empresa Exemplo" |
| `{voces}` / `{podem}` | "você" / "pode", ou "vocês" / "podem" em grupos |
| `{frase_personalizada}` | A frase do contato (some se estiver vazia) |
| `{bloco_catalogo}` | A faixa do catálogo |
| `{botao_whatsapp}` | O botão do WhatsApp |

Use `**texto**` para negrito. Linhas começando com `- ` viram lista.

### 6. Disponibilize o catálogo

O catálogo pode ir por link, em anexo ou das duas formas:

- **Por link:** suba o PDF no OneDrive ou Google Drive, compartilhe com "Qualquer pessoa com o link" e cole em **Texto do e-mail > Link do PDF**. Teste o link em uma aba anônima.
- **Em anexo:** escolha o PDF em **Configurações**. O ideal é até 5 MB.

Um e-mail não consegue apontar para o próprio anexo: para o botão de download funcionar, é preciso usar um link.

### 7. Copie a equipe comercial

Em **Configurações > Cópias do comercial**, informe os e-mails separados por vírgula. Quem está em Cc aparece para o cliente; quem está em Cco não aparece. Cada pessoa recebe uma cópia de cada e-mail do disparo. Os testes não copiam ninguém.

### 8. Teste e envie

1. Na aba **Envio**, mande um teste para você mesmo.
2. Confira o e-mail no computador e no celular, incluindo o link do catálogo e o botão do WhatsApp.
3. Inicie o envio real e digite a confirmação pedida na tela, como `ENVIAR 12`.
4. Acompanhe a barra de progresso. Você pode parar a qualquer momento.

Se houver mais contatos que o limite, rode de novo: quem já recebeu é pulado.

## Comandos no terminal

O painel faz tudo isso, mas os comandos também estão disponíveis:

| Comando | O que faz |
|---|---|
| `python painel.py` | Abre o painel |
| `python disparo.py menu` | Abre o menu no terminal |
| `python disparo.py preparar planilha.xlsx` | Gera `revisao_envio.xlsx` |
| `python disparo.py revisar` | Revisa os contatos linha a linha |
| `python disparo.py previa --abrir` | Gera as prévias e abre a primeira |
| `python disparo.py contatos planilha.xlsx` | Mostra a relação empresa, nome e telefone |
| `python disparo.py enviar` | Simula o envio, sem mandar nada |
| `python disparo.py enviar --real` | Envia de verdade |
| `python disparo.py diagnosticar` | Testa a conexão de e-mail |

## Para evitar a caixa de spam

- Envie apenas para empresas com quem faz sentido falar e respeite quem pedir para sair da lista.
- Use remetente, resposta e assinatura com o mesmo e-mail.
- Peça ao responsável pelo domínio para configurar SPF, DKIM e DMARC.
- Prefira o catálogo por link a um anexo pesado.
- Se um destinatário recusar a mensagem, teste o estilo `SIMPLES` em **Configurações**.
- Não reenvie várias vezes o mesmo e-mail recusado. Ligue e peça para liberar o seu endereço.

Nada garante a entrega: cada empresa tem o próprio filtro. O status `ENVIADO` significa que o servidor aceitou a mensagem, não que ela chegou. Recusas posteriores aparecem na sua caixa de entrada.

## Problemas comuns

| Problema | O que fazer |
|---|---|
| `python` não é reconhecido | Reinstale o Python marcando **Add python.exe to PATH** |
| "A planilha está aberta no Excel" | Feche o Excel e tente de novo |
| "Login recusado" no Microsoft 365 | Use o modo `OUTLOOK` |
| O modo `OUTLOOK` não funciona | Use o Outlook clássico e rode `pip install pywin32` |
| Envio bloqueado por colchetes | Troque os campos `[entre colchetes]` do texto |
| O painel mostra um erro depois de atualizar | Feche e abra o painel de novo, depois aperte Ctrl + F5 |
| O link do catálogo pede login | Compartilhe como "Qualquer pessoa com o link" |

## Seus dados ficam fora do repositório

O `.gitignore` deixa fora do Git as planilhas, o histórico de envios, a configuração, os textos salvos, os catálogos e as imagens:

| Arquivo | Conteúdo |
|---|---|
| `*.xlsx` | Empresas e contatos |
| `log_envios.csv` | Histórico de envios |
| `config.txt` | Dados do remetente e da equipe |
| `textos.json` | Texto do e-mail e link do catálogo |
| `*.pdf`, `*.png` | Catálogo e logo |

Ao usar a ferramenta, siga a LGPD e as regras da sua empresa sobre contatos comerciais.
