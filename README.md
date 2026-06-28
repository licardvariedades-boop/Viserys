# Fechamento Dropshipping Mercado Livre

Dashboard web estático para conciliar a planilha de pedidos da plataforma de dropshipping com o relatório de vendas do Mercado Livre.

## Como usar

Abra `index.html` no navegador e importe:

1. A planilha de pedidos da plataforma.
2. O relatório de vendas do Mercado Livre.

O app cruza as vendas pelo identificador:

- Pedidos: `Número do Pedido no Canal`
- Mercado Livre: `N.º de venda`

## Indicadores

- `Recebido líquido`: soma de `Total (BRL)` do Mercado Livre.
- `Custo de produtos`: soma de `Custo Total` da plataforma.
- `Lucro líquido`: recebido líquido menos custo de produtos.
- `Margem`: lucro líquido dividido pelo recebido líquido.
- `ROI`: lucro líquido dividido pelo custo de produtos.

Por padrão, os KPIs usam apenas vendas conciliadas e ignoram vendas canceladas. Desative `Sem canceladas` para auditar cancelamentos dentro do fechamento.

O painel `Componentes Mercado Livre` soma as colunas monetárias do relatório: receita por produtos, acréscimos, tarifas, envio, descontos, bônus, cancelamentos e reembolsos. A `Diferença conferência` mostra se a soma dessas colunas diverge do `Total (BRL)`.

## Exportação

Os botões do topo exportam:

- Backup da sessão em `.json`, para guardar ou levar para outro navegador.
- Excel com abas `Resumo`, `Fechamento` e `Divergencias`.
- CSV da aba visível na tabela.

## Acesso e Salvamento

O dashboard exige login e salva os dados no Postgres. Não há mais sessão local nem chave privada por navegador.

Usuário inicial:

- `jones260`

Senha inicial:

- configure na Vercel como `APP_PASSWORD`.

## Banco na Vercel

Para salvar na nuvem, configure em `Vercel > Project > Settings > Environment Variables`:

- `DATABASE_URL`: Service URI do Postgres da Aiven, com `sslmode=require`.
- `APP_USER`: `jones260`
- `APP_PASSWORD`: senha do dashboard.
- `APP_SESSION_SECRET`: qualquer texto longo e aleatório para assinar o cookie de login.
- `PG_CA_CERT`: opcional. Cole o CA certificate da Aiven se quiser validação SSL completa.

Depois de salvar as variáveis, faça um redeploy na Vercel.
