# Painel de Compras

Gerenciador web de pedidos de loja para supermercados. O sistema usa uma base histórica/cadastral com setor e classificação de giro e combina suas informações com as três atualizações operacionais do ERP: `consolidado3meses`, `SALDO` e `Zerado`. Depois calcula prioridades de abastecimento e exporta uma fila de pedido em CSV.

Aplicação publicada: <https://jottacostta.github.io/PaineldeCompras/>

Snapshot anterior: [`versoes_anteriores/PaineldeCompras_v1_antes_importacao_arquivos_separados/`](./versoes_anteriores/PaineldeCompras_v1_antes_importacao_arquivos_separados/)

## Recursos

- Dashboard de ruptura, risco de cobertura, saldo do depósito e fila de pedidos.
- Conciliação exclusivamente pela descrição normalizada do produto; o código permanece apenas como referência visual.
- Estoque único para produtos de venda ligados a produtos-base de produção, como `QUE MUC YEMA FATIADO kg` → `QUE MUC YEMA FRACIONADO kg`.
- Cadastro manual de produtos e vínculo manual de estoque, persistidos localmente no navegador.
- Parâmetros de cálculo fixos e exibidos no painel: venda/dia pelo MDV30, faixa crítica de 0 a 3 dias, meta de 7 dias e teto de 14 dias.
- Histórico local de decisões para aprendizado futuro, exportável pelo botão **Histórico ML**.
- Importação de históricos JSON de outros computadores ou períodos, com união sem duplicação.
- Reconhecimento de `Saldo.xlsx`, `consolidado3meses.xls` tabulado e `zerado.xlsx` selecionados juntos ou individualmente, com aviso quando a carga estiver parcial.
- Venda/dia operacional calculada somente pelo `MDV30` do `consolidado3meses`: `(Qtde Venda 3M ÷ 4) ÷ (Dias Venda 3M ÷ 4)`. O campo `MDV (110)` do ERP não é usado; o `SALDO` continua fornecendo estoque da loja, saldo e pedidos.
- Cobertura operacional em três faixas: 0 a 3 dias crítico, 4 a 6 dias atenção e acima de 6 dias meta. O pedido continua limitado ao saldo disponível do depósito e a uma cobertura máxima de 14 dias.
- A fila `Produtos para pedido` mostra lado a lado `Venda 15D — SALDO`, `Venda 30D — SALDO` e `Venda/dia — MDV30`, permitindo conferir a tendência recente antes de liberar o pedido.
- Uso opcional de `vendastrimestrelaticios_dashboard_pedidos.xlsx` como base histórica: setor e giro vêm dela, enquanto estoque, saldo, loja, pendências e vendas atuais vêm dos três arquivos do ERP selecionados na mesma carga.

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

O histórico corrente é persistido automaticamente em `localStorage` para não desaparecer ao recarregar a página. O JSON exportado funciona como cópia transportável: em outro computador, use **Histórico ML → Importar JSON**. Os históricos importados são mesclados por ID de importação, amostra, vínculo e cadastro manual, evitando duplicação.

## Desenvolvimento

```bash
pnpm install
pnpm dev
```

A aplicação roda na porta `3000`. Para validar a build:

```bash
pnpm build
```


## Publicação no GitHub Pages

O deploy é feito pelo workflow `.github/workflows/pages.yml` após push na branch `main` ou execução manual.

**Configuração inicial obrigatória:** antes da primeira publicação, acesse **Settings → Pages** neste repositório e selecione **Build and deployment → Source → GitHub Actions**. O `GITHUB_TOKEN` usado pelo workflow não tem permissão para criar/habilitar o site Pages automaticamente; por isso, o workflow não tenta executar essa operação. Depois da configuração inicial, o workflow constrói o aplicativo, envia o artefato e publica o site.

A URL esperada é <https://jottacostta.github.io/PaineldeCompras/>. Considere uma publicação concluída somente quando o job **Deploy to GitHub Pages** terminar com sucesso e a URL responder com o aplicativo.
