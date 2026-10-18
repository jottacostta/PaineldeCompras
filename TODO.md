# Gerenciador de Pedidos de Loja — Entregas

- [x] Importar localmente um arquivo Excel do ERP e validar as abas obrigatórias `SALDO`, `consolidado3meses` e `Zerado`.
- [x] Conciliar registros por código e descrição, indicando itens sem correspondência ou com informações incompletas.
- [x] Permitir cadastrar produtos novos manualmente, com persistência local no navegador, sem duplicar códigos importados.
- [x] Permitir configurar os períodos curto e longo de venda usados na demanda diária e recalcular a fila.
- [x] Permitir vincular um produto de venda a um produto-base de estoque, somando a venda ao estoque único sem duplicar ruptura ou saldo.
- [x] Reconhecer automaticamente vínculos provenientes da aba opcional `PAINEL`, coluna `Pai`.
- [x] Exibir um dashboard com estoque de depósito, saldo, venda, cobertura, pendências, trânsito, rupturas e recomendações de pedido.
- [x] Classificar a prioridade por risco de ruptura, demanda e disponibilidade no depósito; evitar sugestão que gere excesso de estoque.
- [x] Criar filtros por busca, pasta/categoria, prioridade e disponibilidade, além de ranking de urgência, produtos mais vendidos, itens sem estoque e rupturas.
- [x] Criar uma fila de pedidos revisável: ajustar quantidade, selecionar itens e acompanhar o total pronto para pedido.
- [x] Exportar a fila final em CSV compatível com o fluxo operacional do comprador.
- [x] Permitir reutilização recorrente com novo arquivo ERP, sem digitação manual e sem armazenar os arquivos importados.
- [x] Registrar automaticamente importações, parâmetros, variáveis dos produtos, decisões do comprador, vínculos, cadastros e exportações em histórico local estruturado.
- [x] Permitir baixar o histórico em JSON com amostras de treinamento e campos de observação futura para venda, ruptura, aceitação e avaria.
