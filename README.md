# Painel de Compras

Gerenciador web de pedidos de loja para supermercados. O sistema importa localmente um Excel do ERP com as abas `SALDO`, `consolidado3meses` e `Zerado`, calcula prioridades de abastecimento e exporta uma fila de pedido em CSV.

## Recursos

- Dashboard de ruptura, risco de cobertura, saldo do depósito e fila de pedidos.
- Conciliação por código e descrição.
- Estoque único para produtos de venda ligados a produtos-base de produção, como `QUE MUC YEMA FATIADO kg` → `QUE MUC YEMA FRACIONADO kg`.
- Cadastro manual de produtos e vínculo manual de estoque, persistidos localmente no navegador.
- Períodos curto e longo de venda configuráveis.
- Histórico local de decisões para aprendizado futuro, exportável pelo botão **Histórico ML**.

## Histórico de aprendizado

O botão **Histórico ML** baixa `historico-aprendizado-orquestra-AAAA-MM-DD.json`. O arquivo contém um contrato versionado com:

- variáveis observadas no momento da recomendação;
- parâmetros usados no cálculo;
- prioridade e quantidade sugerida;
- seleção e quantidade final ajustada pelo comprador;
- exportações realizadas;
- vínculos e cadastros manuais;
- campos de observação futura para venda, ruptura, aceitação, stockout e avaria.

O Excel original e seus dados brutos não são armazenados no histórico nem versionados neste repositório. Para treinamento supervisionado real, os campos de resultado precisam ser preenchidos com o que aconteceu após o pedido.

## Desenvolvimento

```bash
pnpm install
pnpm dev
```

A aplicação roda na porta `3000`. Para validar a build:

```bash
pnpm build
```
