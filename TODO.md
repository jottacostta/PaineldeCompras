# Gerenciador de Pedidos de Loja — Entregas

- [x] Importar localmente um Excel completo ou arquivos ERP separados e identificar as fontes `SALDO`, `consolidado3meses` e `Zerado`.
- [x] Conciliar registros exclusivamente pela descrição normalizada, indicando itens sem correspondência ou com informações incompletas; o código não participa da união.
- [x] Permitir cadastrar produtos novos manualmente, com persistência local no navegador, sem duplicar códigos importados.
- [x] Permitir configurar os períodos curto e longo de venda usados na demanda diária e recalcular a fila.
- [x] Calcular a venda/dia operacional exclusivamente no `SALDO`, combinando venda média de 15 dias e venda média de 30 dias; manter o `consolidado3meses` como histórico trimestral e estoque cheio do depósito.
- [x] Trabalhar com cobertura crítica de 2 dias, meta de 7 dias e limite máximo de 14 dias, limitando o pedido ao saldo disponível do depósito.
- [x] Permitir vincular um produto de venda a um produto-base de estoque, somando a venda ao estoque único sem duplicar ruptura ou saldo.
- [x] Reconhecer automaticamente vínculos provenientes da aba opcional `PAINEL`, coluna `Pai`.
- [x] Exibir um dashboard com estoque de depósito, saldo, venda, cobertura, pendências, trânsito, rupturas e recomendações de pedido.
- [x] Classificar a prioridade por risco de ruptura, demanda e disponibilidade no depósito; evitar sugestão que gere excesso de estoque.
- [x] Criar filtros por busca, pasta/categoria, prioridade e disponibilidade, além de ranking de urgência, produtos mais vendidos, itens sem estoque e rupturas.
- [x] Criar uma fila de pedidos revisável: ajustar quantidade, selecionar itens e acompanhar o total pronto para pedido.
- [x] Exportar a fila final em CSV compatível com o fluxo operacional do comprador.
- [x] Permitir reutilização recorrente com novo arquivo ERP, sem digitação manual e sem armazenar os arquivos importados.
- [x] Registrar automaticamente importações, parâmetros, variáveis dos produtos, decisões do comprador, vínculos, cadastros e exportações em histórico local estruturado.
- [x] Aceitar uma planilha completa com as três abas ou os arquivos separados `Saldo.xlsx`, `consolidado3meses.xls` e `zerado.xlsx`, inclusive export tabulado com extensão `.xls`.
- [x] Aceitar a base histórica `vendastrimestrelaticios_dashboard_pedidos.xlsx` separadamente, usando setor e classificação de giro sem deixar seus números antigos sobrescrever as três atualizações atuais do ERP.
- [x] Permitir baixar o histórico em JSON com amostras de treinamento e campos de observação futura para venda, ruptura, aceitação e avaria.
- [x] Permitir importar históricos JSON de outros computadores/períodos, unir os registros sem duplicação e persistir o resultado no `localStorage`.

- [ ] Configurar uma vez **Settings → Pages → Build and deployment → Source → GitHub Actions**; o workflow não consegue criar o site Pages com o `GITHUB_TOKEN` padrão.
- [ ] Validar que o workflow de publicação conclua os jobs de build e deploy e que a URL pública abra o aplicativo.
- [x] Remover a tentativa de habilitar Pages por `enablement: true` no workflow e documentar o pré-requisito manual.
