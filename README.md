# 📧 Envio automático de convites do curso

Projeto em Google Apps Script para enviar, por e-mail, convites personalizados aos participantes selecionados para uma nova turma. Cada mensagem contém o nome do participante e um botão de acesso ao grupo do WhatsApp.

## Funcionalidades

- Seleção da aba da planilha por uma janela visual.
- Personalização da mensagem com o nome do participante.
- Corpo do e-mail em HTML e texto puro.
- Botão e link direto para o grupo do WhatsApp.
- Limite padrão de 40 convites por execução.
- Verificação da cota diária disponível no Google.
- Validação básica dos endereços de e-mail.
- Registro de status, data de envio e erros na planilha.
- Bloqueio de execuções simultâneas.
- Proteção contra reenvio de linhas marcadas como `SIM`.

## Estrutura da planilha

A primeira linha pode ser usada como cabeçalho. Os participantes devem começar na linha 2:

| Coluna | Conteúdo | Exemplo |
| --- | --- | --- |
| A | Nome do participante | Maria da Silva |
| B | E-mail | maria@example.com |
| C | Status preenchido pelo script | SIM |
| D | Data de envio preenchida pelo script | 19/08/2026 10:00 |
| E | Erro ou observação | Endereço de e-mail inválido |

## Instalação

1. Abra a planilha no Google Sheets.
2. Acesse `Extensões` > `Apps Script`.
3. Copie o conteúdo de `envioEmails.gs` para um arquivo de script.
4. Crie um arquivo HTML chamado `SheetPicker` e copie o conteúdo de `SheetPicker.html`.
5. Salve o projeto.

## Configuração obrigatória

No início da função `enviarConvitesComAba`, ajuste estas constantes:

```javascript
const nomePrograma = "NOME DO PROGRAMA";
const nomeTrilha = "NOME DA TRILHA OU CURSO";
const linkGrupoWhatsApp = "https://chat.whatsapp.com/SEU_CODIGO";
const nomeEquipe = "NOME DA EQUIPE";
const nomeInstituto = "NOME DA INSTITUIÇÃO";
const MAXIMO_POR_EXECUCAO = 40;
```

O script não permite iniciar o envio enquanto `linkGrupoWhatsApp` não contiver um endereço válido.

## Execução

1. Na primeira utilização, selecione a função `autorizarServicos` no editor do Apps Script.
2. Clique em **Executar** e autorize o acesso à planilha e ao envio de e-mails.
3. Depois, selecione e execute a função `enviarConvites`.
4. Escolha a aba que contém os participantes.
5. Confirme o envio.

Ao atingir o limite de 40 convites, execute novamente a função. As linhas já marcadas como `SIM` serão ignoradas.

## Teste recomendado

Antes do envio para toda a turma, crie uma aba de teste com apenas seu nome e seu e-mail. Configure o link do grupo e execute o processo nessa aba para conferir assunto, texto, botão e remetente.

## Observações

- O envio utiliza `MailApp`, sujeito às cotas da conta Google.
- O projeto envia o link do WhatsApp por e-mail; ele não envia mensagens diretamente pelo WhatsApp.
- Não apague o status `SIM` de uma linha já processada, pois isso permitirá um novo envio.
