# Projeto Google Apps Script

Este projeto foi clonado do Google Apps Script utilizando o `clasp`.

## Informações Gerais
- **ID do Script:** 1KCei6a9N-Wu5WpBBmjdXKzeQbf0Zj7w_3V4AjY1synZ7WobNSR_kPAep
- **Diretório dos arquivos:** src/
- **Arquivos principais:**
  - `appsscript.json`: Configuração do projeto Apps Script
  - `Código.js`: Código principal em JavaScript
  - `index.html`: Interface HTML utilizada pelo script

## Como usar o projeto
1. Instale o [clasp](https://github.com/google/clasp) globalmente:
   ```sh
   npm install -g @google/clasp
   ```
2. Faça login com sua conta Google:
   ```sh
   clasp login
   ```
3. Para sincronizar alterações com o Apps Script:
   ```sh
   clasp push
   ```
4. Para baixar as últimas alterações do Apps Script:
   ```sh
   clasp pull
   ```

## Estrutura das Abas da Planilha Google

Abaixo estão as estruturas esperadas para cada aba utilizada pelo projeto:

- **DadosTeste** [Sistema; Núm.PAE/SAJ; Interessado; Entrada; Situação; Assimetria; Observação; UG Origem; Assunto; Sub Assunto; ACI Responsável; Destino; Saída]
- **Config Usuarios** [E-mail; Usuário]
- **Config Assuntos** [Assunto; Sub Assunto]

> Mantenha a ordem e os nomes dos campos conforme acima para garantir o funcionamento correto do script.

## Observações
- Certifique-se de ter permissão de edição no projeto do Apps Script.
- O arquivo `.clasp.json` já está configurado para este projeto.

---

Este repositório serve para facilitar o desenvolvimento e versionamento do seu projeto Google Apps Script localmente.

// deploymentId salvo para deploys futuros do Google Apps Script
// Sempre usar: AKfycby8tVF5st56XfelYzuOOXJETKkMmyT0LwFLxyYHPP93lfZBrO4eRG3xbGiQ6xmZWmY2
