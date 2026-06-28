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

## Salvamento

Depois que uma planilha é importada, a sessão é salva automaticamente no armazenamento local do navegador. Ao abrir o mesmo `index.html` de novo no mesmo navegador, os dados voltam sem precisar importar tudo novamente.

Use o botão de pasta para carregar um backup `.json` e o botão de lixeira para limpar os dados salvos.
