# Histórico de Versões

<<<<<<< HEAD
## v1.0.8 - 08/08/2025 (Estável)
- Sistema restaurado para o commit 38c7e12: melhorias visuais nos filtros e tabela, agrupamento, cards, cabeçalho fixo, linhas alternadas, responsividade e contraste.
- Versão marcada como estável após validação e commit de reinício.
=======
## v1.0.7 - 07/08/2025 (estável)
- O filtro "UG Origem" agora utiliza um dropdown dinâmico (<select>) populado automaticamente com os valores da coluna SETOR da aba "Config Setor".
- O campo de filtro foi atualizado para <select> e as opções são carregadas via função Apps Script (getUgOrigemOptions).
- Corrigida duplicidade de inicialização (DOMContentLoaded) para evitar conflitos de carregamento.
- Ajuste visual e sintático para garantir funcionamento correto do filtro e do modal.

## v1.0.6 - 07/08/2025
- Alteração visual: sistema agora utiliza tons de verde em toda a interface (background, botões, cards, tabela, footer, badges, etc).

## v1.0.5 - 06/08/2025
- Filtros ajustados para exibir 4 campos por linha em telas grandes.

## v1.0.4 - 06/08/2025
- Adicionado footer destacado: "Desenvolvido por DTIGD © 2025" ao final da página.

## v1.0.3 - 06/08/2025
- Cabeçalho da tabela permanece fixo durante a rolagem dos dados.

## v1.0.2 - 06/08/2025
- Botão "Adicionar Novo Registro" movido para ao lado do botão "Atualizar Dados" acima da tabela.
- Removido botão de adicionar registro da área central acima dos filtros.

## v1.0.1 - 06/08/2025
- Card de filtros recolhível/expansível na interface web.
- Card da tabela ocupa toda a largura da tela.
- Melhorias visuais: agrupamento em cards, responsividade, espaçamento, cabeçalho fixo, linhas alternadas, contraste, hover.
>>>>>>> 0f10501 (v1.0.7 estável: Filtro UG Origem agora é dropdown dinâmico, integração Apps Script, correção de inicialização e sintaxe)

## v1.0.0 - 06/08/2025
- Projeto clonado do Google Apps Script via `clasp`.
- Estrutura inicial criada na pasta `src`.
- Arquivos principais adicionados: `appsscript.json`, `Código.js`, `index.html`.
- README com instruções de uso e informações do projeto.
