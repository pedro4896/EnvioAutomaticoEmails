# 📧 Envio Automático de Certificados por E-mail

Este projeto permite o envio automático de certificados em PDF por e-mail utilizando **Google Sheets**, **Google Drive** e **Google Apps Script**.

É uma solução simples e eficiente para escolas, cursos, eventos e instituições que precisam distribuir certificados de forma rápida e organizada.

---

## 🚀 Funcionalidades

- Envio automático de certificados em PDF
- Associação inteligente entre nome do aluno e arquivo
- Tratamento automático de:
  - Acentuação
  - Espaços duplicados
  - Letras maiúsculas/minúsculas
- Corpo de e-mail personalizável
- Variável para nome do curso

---

## 🛠 Tecnologias Utilizadas

- Google Apps Script (.gs)
- Google Sheets
- Google Drive
- Gmail API

---

## 📋 Estrutura Necessária

### 📄 Planilha

A planilha deve conter as seguintes colunas:

| Nome | Email |
|------|-------|
| João Silva | joao@email.com |
| Maria Oliveira | maria@email.com |

> A primeira linha deve conter os títulos das colunas.

---

### 📁 Certificados

Todos os certificados devem estar em uma única pasta no Google Drive e nomeados exatamente assim:<br>
João Silva.pdf<br>
Maria Oliveira.pdf<br>


---

## ⚙️ Como Configurar

1. Abra o Google Sheets
2. Clique em: Extensões → Apps Script
3. Apague qualquer código existente.
4. Cole o código do projeto.

5. Altere a seguinte linha no script:

```javascript
var pastaId = "COLE_AQUI_O_ID_DA_PASTA";
```

## ▶️ Como Executar

No editor do Apps Script:

Selecione a função:
```javascript
enviarCertificados
```

Clique no botão:
```javascript
▶ Executar
```

Na primeira execução:

Autorize o acesso à sua conta Google.

## ✅ Comportamento do Sistema

O script lê cada linha da planilha.

Procura o certificado correspondente na pasta do Drive.

Envia o e-mail automaticamente com o PDF em anexo.

Registra logs no console do Apps Script.

## 🧾 Licença

Este projeto é de uso livre para fins educacionais e organizacionais.
