# Automação de Envio de E-mails

Automação desenvolvida em Python para envio individualizado de
e-mails utilizando a Microsoft Graph API.

## Tecnologias

- Python
- Microsoft Graph API
- Microsoft Authentication Library (MSAL)
- Pandas
- OpenPyXL
- Requests

## Funcionalidades

- Leitura de destinatários a partir de uma planilha Excel
- Validação e remoção de e-mails duplicados
- Autenticação com Microsoft Entra ID
- Envio de e-mails através do Microsoft Graph
- Suporte a imagens inline
- Envio de e-mail de teste antes do disparo
- Confirmação manual antes do envio em massa
- Registro dos resultados em arquivo CSV

## Configuração

Crie um arquivo `.env`:

CLIENT_ID=...
TENANT_ID=...
EMAIL_TESTE=...

Os arquivos da campanha devem ser colocados localmente
na pasta do projeto.

## Execução

Instale as dependências:

pip install -r requirements.txt

Execute:

py disparo_email.py