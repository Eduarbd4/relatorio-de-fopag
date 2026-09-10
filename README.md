# Relatórios de Pagamento DANE-SE

Sistema estático para importar a planilha de fechamento, identificar todos os colaboradores e gerar relatórios individuais de detalhamento de pagamento.

## Publicar no GitHub Pages

1. Crie um repositório novo no GitHub.
2. Envie **o conteúdo da pasta `dist`** para a raiz do repositório.
3. Abra `Settings` → `Pages`.
4. Em `Build and deployment`, escolha `Deploy from a branch`.
5. Selecione a branch `main`, pasta `/ (root)` e clique em `Save`.

## Acesso

Senha inicial: `DANE-SE2015`

Para alterar, abra `app.js` e edite a primeira linha (`ACCESS_PASSWORD`).

> Observação: por ser um site estático, a senha é uma barreira visual e não substitui autenticação segura em servidor.

## Uso

Após entrar, importe um arquivo `.xlsx`, `.xlsm` ou `.xls`. O sistema seleciona automaticamente a aba de fechamento mais recente, mas permite escolher outra aba e ajustar a competência. Os relatórios podem ser visualizados individualmente ou impressos/salvos em PDF em lote.
