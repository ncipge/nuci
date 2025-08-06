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

## Estrutura da Planilha Google

A planilha utilizada por este projeto deve conter os seguintes campos, nesta ordem:

1. **Status**
2. **Sistema**
3. **Núm.PAE/SAJ**
4. **Interessado**
5. **Entrada**
6. **Situação**
7. **Assimetria**
8. **Observação**
9. **UG Origem**
10. **Assunto**
11. **Sub Assunto**
12. **ACI Responsável**
13. **Destino**
14. **Saída**

- **Cabeçalho:** A primeira linha da planilha deve conter exatamente esses nomes de campos.
- **Aba principal:** Recomenda-se nomear como "Dados" ou conforme a lógica do script.
- **Permissões:** O script precisa de permissão de edição na planilha.

> Mantenha a ordem e os nomes dos campos conforme acima para garantir o funcionamento correto do script.

## Observações
- Certifique-se de ter permissão de edição no projeto do Apps Script.
- O arquivo `.clasp.json` já está configurado para este projeto.

---

Este repositório serve para facilitar o desenvolvimento e versionamento do seu projeto Google Apps Script localmente.
