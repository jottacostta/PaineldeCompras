import {
  PRIORITY_META,
  createManualProduct,
  exportRows,
  importWorkbook,
  makeCsv,
  recalculateProducts,
} from './engine.js';
import {
  clearHistory,
  downloadHistory,
  historySummary,
  loadHistory,
  mergeImportedHistory,
  recordEvent,
  recordExport,
  recordImport,
  recordManualProduct,
  recordMapping,
  saveHistory,
  updateCurrentDecisions,
} from './history.js';
import './styles.css';

const PAGE_SIZE = 25;
const formatNumber = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 0 });
const formatDecimal = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const formatCurrency = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 });

const state = {
  products: [],
  categories: [],
  stores: [],
  fileName: '',
  currentFile: null,
  history: loadHistory(),
  currentSnapshotId: null,
  importedAt: null,
  stats: null,
  params: { shortPeriod: 15, longPeriod: 30, criticalDays: 2, targetDays: 7, maxDays: 14 },
  mappings: loadStoredMappings(),
  appliedMappings: [],
  manualEntries: loadStoredManualEntries(),
  selectedIds: new Set(),
  adjustments: new Map(),
  currentView: 'dashboard',
  currentPage: 1,
  filters: { search: '', priority: 'ALL', availability: 'ALL', category: 'ALL', store: 'ALL' },
};

const el = {
  importView: document.querySelector('#importView'),
  appView: document.querySelector('#appView'),
  fileInput: document.querySelector('#fileInput'),
  uploadCard: document.querySelector('#uploadCard'),
  newImportButton: document.querySelector('#newImportButton'),
  historyButton: document.querySelector('#historyButton'),
  addProductButton: document.querySelector('#addProductButton'),
  linkProductButton: document.querySelector('#linkProductButton'),
  shortPeriodInput: document.querySelector('#shortPeriodInput'),
  longPeriodInput: document.querySelector('#longPeriodInput'),
  targetDaysInput: document.querySelector('#targetDaysInput'),
  maxDaysInput: document.querySelector('#maxDaysInput'),
  recalculateButton: document.querySelector('#recalculateButton'),
  storeSelect: document.querySelector('#storeSelect'),
  priorityList: document.querySelector('#priorityList'),
  coverageBars: document.querySelector('#coverageBars'),
  conservativeNote: document.querySelector('#conservativeNote'),
  searchInput: document.querySelector('#searchInput'),
  priorityFilter: document.querySelector('#priorityFilter'),
  availabilityFilter: document.querySelector('#availabilityFilter'),
  categoryFilter: document.querySelector('#categoryFilter'),
  clearFiltersButton: document.querySelector('#clearFiltersButton'),
  ordersTableBody: document.querySelector('#ordersTableBody'),
  tableTotal: document.querySelector('#tableTotal'),
  paginationText: document.querySelector('#paginationText'),
  prevPageButton: document.querySelector('#prevPageButton'),
  nextPageButton: document.querySelector('#nextPageButton'),
  selectAllInput: document.querySelector('#selectAllInput'),
  exportButton: document.querySelector('#exportButton'),
  queueNavCount: document.querySelector('#queueNavCount'),
  alertNavCount: document.querySelector('#alertNavCount'),
  importMeta: document.querySelector('#importMeta'),
  pageSubtitle: document.querySelector('#pageSubtitle'),
  rulesDialog: document.querySelector('#rulesDialog'),
  openRulesButton: document.querySelector('#openRulesButton'),
  closeRulesButton: document.querySelector('#closeRulesButton'),
  toastRegion: document.querySelector('#toastRegion'),
  productDialog: document.querySelector('#productDialog'),
  productForm: document.querySelector('#productForm'),
  closeProductButton: document.querySelector('#closeProductButton'),
  cancelProductButton: document.querySelector('#cancelProductButton'),
  mappingDialog: document.querySelector('#mappingDialog'),
  mappingForm: document.querySelector('#mappingForm'),
  closeMappingButton: document.querySelector('#closeMappingButton'),
  cancelMappingButton: document.querySelector('#cancelMappingButton'),
  mappingList: document.querySelector('#mappingList'),
  productCodeList: document.querySelector('#productCodeList'),
  productDescriptionList: document.querySelector('#productDescriptionList'),
  historyDialog: document.querySelector('#historyDialog'),
  closeHistoryButton: document.querySelector('#closeHistoryButton'),
  closeHistoryActionButton: document.querySelector('#closeHistoryActionButton'),
  downloadHistoryButton: document.querySelector('#downloadHistoryButton'),
  clearHistoryButton: document.querySelector('#clearHistoryButton'),
  importHistoryButton: document.querySelector('#importHistoryButton'),
  historyFileInput: document.querySelector('#historyFileInput'),
  historyImports: document.querySelector('#historyImports'),
  historySamples: document.querySelector('#historySamples'),
  historyObserved: document.querySelector('#historyObserved'),
  historyExports: document.querySelector('#historyExports'),
  kpi: {
    rupture: document.querySelector('#kpiRupture'),
    ruptureText: document.querySelector('#kpiRuptureText'),
    risk: document.querySelector('#kpiRisk'),
    riskText: document.querySelector('#kpiRiskText'),
    suggested: document.querySelector('#kpiSuggested'),
    suggestedText: document.querySelector('#kpiSuggestedText'),
    blocked: document.querySelector('#kpiBlocked'),
    blockedText: document.querySelector('#kpiBlockedText'),
    reconciled: document.querySelector('#kpiReconciled'),
    reconciledText: document.querySelector('#kpiReconciledText'),
  },
};

function loadStoredManualEntries() {
  try {
    return JSON.parse(localStorage.getItem('orquestra-manual-products') || '[]');
  } catch {
    return [];
  }
}

function saveStoredManualEntries() {
  localStorage.setItem('orquestra-manual-products', JSON.stringify(state.manualEntries));
}

function loadStoredMappings() {
  try {
    return JSON.parse(localStorage.getItem('orquestra-product-mappings') || '[]');
  } catch {
    return [];
  }
}

function saveStoredMappings() {
  localStorage.setItem('orquestra-product-mappings', JSON.stringify(state.mappings));
}

function persistHistory() {
  const result = saveHistory(state.history);
  state.history = result.history;
  if (!result.persisted) showToast('O histórico ficou apenas nesta sessão; baixe o JSON para não perder os registros.', 'warning');
  if (result.trimmed) showToast('O histórico local foi compactado. Baixe o JSON para manter um arquivo completo.', 'warning');
  return result;
}

function refreshHistoryDialog() {
  const summary = historySummary(state.history);
  el.historyImports.textContent = formatNumber.format(summary.imports);
  el.historySamples.textContent = formatNumber.format(summary.samples);
  el.historyObserved.textContent = formatNumber.format(summary.observedSamples);
  el.historyExports.textContent = formatNumber.format(summary.exports);
}

function openHistoryDialog() {
  refreshHistoryDialog();
  el.historyDialog.showModal();
}

async function importHistoryFile(file) {
  if (!file) return;
  try {
    const text = await file.text();
    const before = historySummary(state.history);
    state.history = mergeImportedHistory(state.history, text);
    persistHistory();
    const after = historySummary(state.history);
    refreshHistoryDialog();
    showToast(`${formatNumber.format(Math.max(0, after.samples - before.samples))} amostras de histórico importadas para a base de aprendizado.`, 'success');
  } catch (error) {
    showToast(error?.message || 'Não foi possível importar este histórico JSON.', 'error');
  } finally {
    el.historyFileInput.value = '';
  }
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function normalizeTextForUi(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function showToast(message, type = 'info') {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  el.toastRegion.appendChild(toast);
  setTimeout(() => toast.classList.add('is-visible'), 20);
  setTimeout(() => {
    toast.classList.remove('is-visible');
    setTimeout(() => toast.remove(), 240);
  }, 4200);
}

function priorityBadge(priority) {
  const meta = PRIORITY_META[priority];
  return `<span class="priority-badge priority-${meta.tone}">${escapeHtml(meta.label)}</span>`;
}

function formatCoverage(value) {
  if (value === null || value === undefined || !Number.isFinite(value)) return '—';
  if (value > 99) return '99+ dias';
  return `${formatDecimal.format(value)} d`;
}

function productQuantity(product) {
  const adjusted = state.adjustments.get(product.id);
  return adjusted === undefined ? product.suggested : adjusted;
}

function selectedProductCount() {
  return [...state.selectedIds].filter((id) => {
    const product = state.products.find((entry) => entry.id === id);
    return product && productQuantity(product) > 0;
  }).length;
}

function currentStoreLabel() {
  return state.filters.store === 'ALL' ? 'Todas as lojas' : state.filters.store;
}

function getFilteredProducts() {
  const query = state.filters.search.trim().toUpperCase();
  return state.products.filter((product) => {
    if (state.filters.store !== 'ALL' && product.store !== state.filters.store) return false;
    if (state.filters.priority !== 'ALL' && product.priority !== state.filters.priority) return false;
    if (state.filters.availability === 'AVAILABLE' && product.depotAvailable <= 0) return false;
    if (state.filters.availability === 'ZERO' && product.depotAvailable > 0) return false;
    if (state.filters.category !== 'ALL' && (product.sector || product.paste) !== state.filters.category) return false;
    if (query && !`${product.code} ${product.description}`.toUpperCase().includes(query)) return false;
    return true;
  });
}

function getViewProducts() {
  if (state.currentView === 'orders') return getFilteredProducts().filter((product) => state.selectedIds.has(product.id));
  if (state.currentView === 'exceptions') return getFilteredProducts().filter((product) => product.warnings.length || product.priority === 'BLOCKED');
  return getFilteredProducts();
}

function updateKpis() {
  const scoped = state.products.filter((product) => state.filters.store === 'ALL' || product.store === state.filters.store);
  const rupture = scoped.filter((product) => product.priority === 'P0');
  const risk = scoped.filter((product) => product.priority === 'P1');
  const blocked = scoped.filter((product) => product.priority === 'BLOCKED');
  const suggested = scoped.reduce((sum, product) => sum + productQuantity(product), 0);
  const reconciled = scoped.filter((product) => product.hasSaldo && product.hasConsolidated).length;
  const percentage = scoped.length ? Math.round((reconciled / scoped.length) * 100) : 0;

  el.kpi.rupture.textContent = formatNumber.format(rupture.length);
  el.kpi.ruptureText.textContent = rupture.length ? `${formatNumber.format(rupture.reduce((sum, product) => sum + product.suggested, 0))} un. para atender` : 'sem ruptura atendível';
  el.kpi.risk.textContent = formatNumber.format(risk.length);
  el.kpi.riskText.textContent = risk.length ? 'cobertura projetada até 2 dias' : 'nenhum risco imediato';
  el.kpi.suggested.textContent = formatNumber.format(suggested);
  el.kpi.suggestedText.textContent = `${formatNumber.format(selectedProductCount())} itens selecionados para exportação`;
  el.kpi.blocked.textContent = formatNumber.format(blocked.length);
  el.kpi.blockedText.textContent = blocked.length ? `${formatNumber.format(blocked.reduce((sum, product) => sum + product.shortage, 0))} un. sem atendimento` : 'saldo disponível para os críticos';
  el.kpi.reconciled.textContent = `${percentage}%`;
  el.kpi.reconciledText.textContent = `${formatNumber.format(reconciled)} de ${formatNumber.format(scoped.length)} com SALDO + venda`;
  el.queueNavCount.textContent = formatNumber.format(selectedProductCount());
  el.alertNavCount.textContent = formatNumber.format(scoped.filter((product) => product.warnings.length || product.priority === 'BLOCKED').length);
  el.exportButton.disabled = selectedProductCount() === 0;
}

function updatePriorityList() {
  const priorities = state.products
    .filter((product) => ['P0', 'P1', 'P2', 'BLOCKED'].includes(product.priority))
    .filter((product) => state.filters.store === 'ALL' || product.store === state.filters.store)
    .slice(0, 6);

  if (!priorities.length) {
    el.priorityList.innerHTML = '<div class="empty-inline"><strong>Nenhuma prioridade crítica.</strong><span>Importe um arquivo ou reduza os filtros para visualizar a fila.</span></div>';
    return;
  }

  el.priorityList.innerHTML = priorities.map((product) => {
    const selected = state.selectedIds.has(product.id) ? 'checked' : '';
    const quantity = productQuantity(product);
    const warning = product.priority === 'BLOCKED' ? `<span class="micro-alert">saldo indisponível</span>` : '';
    return `<article class="priority-item">
      <label class="check-wrap"><input type="checkbox" data-select-id="${escapeHtml(product.id)}" ${selected} ${quantity <= 0 ? 'disabled' : ''}/><span></span></label>
      <div class="priority-main"><div class="priority-topline">${priorityBadge(product.priority)}${warning}<span>${escapeHtml(product.code)}</span></div><strong>${escapeHtml(product.description)}</strong><small>${escapeHtml(product.sector || product.paste)} · ${escapeHtml(product.giro || 'giro não classificado')} · cobertura ${formatCoverage(product.coverage)}</small></div>
      <div class="priority-quantity"><strong>${formatNumber.format(quantity)}</strong><span>un. sugeridas</span></div>
    </article>`;
  }).join('');
}

function updateCoverageBars() {
  const scoped = state.products.filter((product) => state.filters.store === 'ALL' || product.store === state.filters.store);
  const groups = [
    { label: 'Ruptura / 0 dia', test: (product) => product.coverage !== null && product.coverage <= 0, tone: 'danger' },
    { label: 'Até 2 dias · crítico', test: (product) => product.coverage !== null && product.coverage > 0 && product.coverage <= state.params.criticalDays, tone: 'danger' },
    { label: 'De 2 a 7 dias · abaixo da meta', test: (product) => product.coverage !== null && product.coverage > state.params.criticalDays && product.coverage < state.params.targetDays, tone: 'warning' },
    { label: 'De 7 a 14 dias · faixa segura', test: (product) => product.coverage !== null && product.coverage >= state.params.targetDays && product.coverage <= state.params.maxDays, tone: 'teal' },
    { label: 'Acima de 14 dias · excesso', test: (product) => product.coverage !== null && product.coverage > state.params.maxDays, tone: 'neutral' },
  ];
  const largest = Math.max(1, ...groups.map((group) => scoped.filter(group.test).length));
  el.coverageBars.innerHTML = groups.map((group) => {
    const count = scoped.filter(group.test).length;
    const width = Math.max(count ? 4 : 0, (count / largest) * 100);
    return `<div class="coverage-row"><div class="coverage-label"><span>${escapeHtml(group.label)}</span><strong>${formatNumber.format(count)}</strong></div><div class="coverage-track"><span class="coverage-fill fill-${group.tone}" style="width:${width}%"></span></div></div>`;
  }).join('');
  el.conservativeNote.textContent = `Meta: ${state.params.targetDays}d · Máximo: ${state.params.maxDays}d`;
}

function renderTable() {
  const products = getViewProducts();
  const totalPages = Math.max(1, Math.ceil(products.length / PAGE_SIZE));
  state.currentPage = Math.min(state.currentPage, totalPages);
  const start = (state.currentPage - 1) * PAGE_SIZE;
  const slice = products.slice(start, start + PAGE_SIZE);

  el.tableTotal.textContent = `${formatNumber.format(products.length)} ${products.length === 1 ? 'item' : 'itens'}`;
  el.paginationText.textContent = products.length ? `Exibindo ${start + 1}–${Math.min(start + PAGE_SIZE, products.length)} de ${formatNumber.format(products.length)}` : 'Nenhum item encontrado';
  el.prevPageButton.disabled = state.currentPage <= 1;
  el.nextPageButton.disabled = state.currentPage >= totalPages;
  const visibleSelectable = slice.filter((product) => productQuantity(product) > 0 && product.priority !== 'BLOCKED');
  el.selectAllInput.checked = visibleSelectable.length > 0 && visibleSelectable.every((product) => state.selectedIds.has(product.id));
  el.selectAllInput.indeterminate = visibleSelectable.some((product) => state.selectedIds.has(product.id)) && !el.selectAllInput.checked;

  if (!slice.length) {
    el.ordersTableBody.innerHTML = '<tr><td colspan="11"><div class="table-empty"><strong>Nenhum produto atende aos filtros atuais.</strong><span>Limpe filtros ou importe outra planilha para continuar.</span></div></td></tr>';
    return;
  }

  el.ordersTableBody.innerHTML = slice.map((product) => {
    const selected = state.selectedIds.has(product.id) ? 'checked' : '';
    const quantity = productQuantity(product);
    const disabled = product.priority === 'BLOCKED' || quantity <= 0 ? 'disabled' : '';
    const alerts = product.warnings.length ? `<span class="warning-dot" title="${escapeHtml(product.warnings.join(' · '))}">!</span>` : '';
    const sourceTag = product.isRupture ? '<span class="rupture-tag">ERP zerado</span>' : '';
    const reason = product.warnings[0] || product.reason;
    return `<tr data-product-id="${escapeHtml(product.id)}">
      <td class="checkbox-cell"><label class="check-wrap"><input type="checkbox" data-select-id="${escapeHtml(product.id)}" ${selected} ${disabled}/><span></span></label></td>
      <td>${priorityBadge(product.priority)}</td>
      <td><div class="product-cell"><strong>${escapeHtml(product.description)}</strong><span>${escapeHtml(product.code)} · ${escapeHtml(product.sector || product.paste)} · ${escapeHtml(product.giro || 'giro não classificado')} ${sourceTag}</span></div></td>
      <td>${escapeHtml(product.store)}</td>
      <td>${formatDecimal.format(product.demandDaily)}<small> un./dia</small></td>
      <td>${formatNumber.format(product.storeStock)}</td>
      <td>${formatNumber.format(product.inbound)}</td>
      <td><strong class="${product.depotAvailable <= 0 ? 'number-zero' : ''}">${formatNumber.format(product.depotAvailable)}</strong><small>de ${formatNumber.format(product.depotStock ?? 0)}</small></td>
      <td><span class="coverage-chip ${product.coverage !== null && product.coverage <= 2 ? 'coverage-low' : ''}">${formatCoverage(product.coverage)}</span></td>
      <td><input class="qty-input" type="number" data-qty-id="${escapeHtml(product.id)}" min="0" step="${product.pack || 1}" value="${quantity}" ${product.priority === 'BLOCKED' ? 'disabled' : ''}/><small>múlt. ${formatNumber.format(product.pack || 1)}</small></td>
      <td><div class="reason-cell">${alerts}<span title="${escapeHtml(product.warnings.join(' · '))}">${escapeHtml(reason)}</span></div></td>
    </tr>`;
  }).join('');
}

function syncStoreSelect() {
  el.storeSelect.innerHTML = ['ALL', ...state.stores].map((store) => `<option value="${escapeHtml(store)}" ${store === state.filters.store ? 'selected' : ''}>${store === 'ALL' ? 'Todas as lojas' : escapeHtml(store)}</option>`).join('');
}

function syncCategorySelect() {
  el.categoryFilter.innerHTML = `<option value="ALL">Todos os setores</option>${state.categories.map((category) => `<option value="${escapeHtml(category)}" ${category === state.filters.category ? 'selected' : ''}>${escapeHtml(category)}</option>`).join('')}`;
}

function syncProductLists() {
  el.productCodeList.innerHTML = state.products.map((product) => `<option value="${escapeHtml(product.code)}">${escapeHtml(product.description)}</option>`).join('');
  el.productDescriptionList.innerHTML = state.products.map((product) => `<option value="${escapeHtml(product.description)}">${escapeHtml(product.code)}</option>`).join('');
}

function renderMappingList() {
  if (!state.mappings.length && !state.appliedMappings.length) {
    el.mappingList.innerHTML = '<div class="mapping-empty">Nenhum vínculo manual salvo. Se a planilha tiver PAINEL com a coluna Pai, os vínculos serão sugeridos automaticamente.</div>';
    return;
  }
  const allMappings = [...state.mappings.map((mapping) => ({ ...mapping, manual: true })), ...state.appliedMappings.filter((mapping) => mapping.origin === 'PAINEL').map((mapping) => ({ ...mapping, manual: false }))];
  el.mappingList.innerHTML = `<div class="mapping-list-head"><strong>Vínculos reconhecidos</strong><span>${formatNumber.format(allMappings.length)}</span></div>${allMappings.map((mapping, index) => `<div class="mapping-row"><span class="mapping-source">${mapping.manual ? 'Manual' : 'PAINEL'}</span><div><strong>${escapeHtml(mapping.saleDescription || mapping.saleCode)}</strong><small>${escapeHtml(mapping.saleCode)} → ${escapeHtml(mapping.stockDescription || mapping.stockCode)}</small></div>${mapping.manual ? `<button class="icon-button remove-mapping" type="button" data-mapping-index="${index}">×</button>` : '<span class="mapping-lock">auto</span>'}</div>`).join('')}`;
}

function applyManualProducts(baseProducts) {
  const existing = new Set(baseProducts.map((product) => String(product.code || '').trim().toUpperCase()).filter(Boolean));
  const manualProducts = state.manualEntries
    .map((entry) => createManualProduct(entry, state.params))
    .filter((product) => !existing.has(String(product.code || '').trim().toUpperCase()));
  return [...baseProducts, ...manualProducts];
}

function updateViewNarrative() {
  const navigation = document.querySelectorAll('.nav-item');
  navigation.forEach((button) => button.classList.toggle('is-active', button.dataset.view === state.currentView));
  if (state.currentView === 'orders') {
    el.pageSubtitle.textContent = 'Esta visão mostra somente itens selecionados para exportação. Ajuste a quantidade antes de concluir.';
  } else if (state.currentView === 'exceptions') {
    el.pageSubtitle.textContent = 'Itens bloqueados, sem conciliação completa ou com divergências que merecem conferência.';
  } else {
    el.pageSubtitle.textContent = 'Revise prioridades, ajuste quantidades e exporte uma fila conferida.';
  }
}

function renderAll() {
  updateViewNarrative();
  updateKpis();
  updatePriorityList();
  updateCoverageBars();
  renderTable();
}

function setInitialSelection() {
  state.selectedIds = new Set(state.products.filter((product) => ['P0', 'P1', 'P2', 'P3'].includes(product.priority) && product.suggested > 0).map((product) => product.id));
  state.adjustments = new Map();
}

async function handleFile(fileInput) {
  const files = Array.isArray(fileInput) ? fileInput.filter(Boolean) : Array.from(fileInput ?? (fileInput ? [fileInput] : []));
  if (!files.length) return;
  const invalidFile = files.find((file) => !/\.(xlsx|xls|xlsm)$/i.test(file.name));
  const allowed = !invalidFile;
  if (!allowed) {
    showToast(`Arquivo não aceito: ${invalidFile.name}. Use .xlsx, .xls ou .xlsm.`, 'error');
    return;
  }
  el.uploadCard.classList.add('is-loading');
  try {
    const result = await importWorkbook(files, state.params, state.mappings);
    state.products = applyManualProducts(result.products);
    state.categories = [...new Set(state.products.map((product) => product.sector || product.paste).filter(Boolean))].sort();
    state.stores = [...new Set(state.products.map((product) => product.store).filter(Boolean))].sort();
    state.stats = result.stats;
    state.appliedMappings = result.mappings ?? [];
    state.fileName = result.fileName;
    state.currentFile = files.length === 1 ? files[0] : files;
    state.importedAt = result.importedAt;
    const historyImport = recordImport(state.history, {
      fileName: result.fileName,
      importedAt: result.importedAt,
      params: state.params,
      stats: result.stats,
      mappings: result.mappings,
      products: state.products,
    });
    state.history = historyImport.history;
    state.currentSnapshotId = historyImport.snapshotId;
    state.filters = { search: '', priority: 'ALL', availability: 'ALL', category: 'ALL', store: 'ALL' };
    state.currentPage = 1;
    state.currentView = 'dashboard';
    setInitialSelection();
    updateCurrentDecisions(state.history, state.currentSnapshotId, state.products, state.selectedIds, state.adjustments);
    persistHistory();
    syncStoreSelect();
    syncCategorySelect();
    syncProductLists();
    renderMappingList();
    el.importMeta.textContent = `IMPORTADO · ${result.fileName.toUpperCase()}`;
    el.importView.classList.add('hidden');
    el.appView.classList.remove('hidden');
    renderAll();
    const missing = result.stats.missingSheets?.length ? ` Fontes ausentes nesta carga: ${result.stats.missingSheets.join(', ')}.` : '';
    const baseInfo = result.stats.baseMatched ? ` ${formatNumber.format(result.stats.baseMatched)} produtos receberam setor/giro da base histórica.` : '';
    showToast(`${formatNumber.format(state.products.length)} produtos processados. ${formatNumber.format(state.appliedMappings.length)} vínculos de estoque aplicados.${baseInfo}${missing}`, missing ? 'warning' : 'success');
  } catch (error) {
    console.error(error);
    showToast(error?.message || 'Não foi possível ler a planilha. Confira a estrutura do arquivo.', 'error');
  } finally {
    el.uploadCard.classList.remove('is-loading');
    el.fileInput.value = '';
  }
}

function recalculate() {
  const shortPeriod = Math.max(1, Number(el.shortPeriodInput.value) || 15);
  const longPeriod = Math.max(shortPeriod, 30, Number(el.longPeriodInput.value) || 30);
  const criticalDays = 2;
  const targetDays = Math.max(1, Number(el.targetDaysInput.value) || 7);
  const maxDays = Math.max(targetDays, Number(el.maxDaysInput.value) || 14);
  state.params = { shortPeriod, longPeriod, criticalDays, targetDays, maxDays };
  el.shortPeriodInput.value = shortPeriod;
  el.longPeriodInput.value = longPeriod;
  el.targetDaysInput.value = targetDays;
  el.maxDaysInput.value = maxDays;
  const selectedBefore = new Set(state.selectedIds);
  state.products = recalculateProducts(state.products, state.params);
  state.selectedIds = new Set(state.products.filter((product) => selectedBefore.has(product.id) && product.suggested > 0 && product.priority !== 'BLOCKED').map((product) => product.id));
  state.adjustments.clear();
  state.currentPage = 1;
  recordEvent(state.history, {
    type: 'recalculate',
    snapshotId: state.currentSnapshotId,
    parameters: { ...state.params },
    productCount: state.products.length,
  });
  updateCurrentDecisions(state.history, state.currentSnapshotId, state.products, state.selectedIds, state.adjustments);
  persistHistory();
  renderAll();
  showToast(`Fila recalculada usando SALDO de ${shortPeriod}D/${longPeriod}D, risco em ${criticalDays} dias, meta de ${targetDays} dias e teto de ${maxDays} dias.`, 'success');
}

function openProductDialog() {
  el.productForm.reset();
  el.productForm.elements.pack.value = 1;
  el.productForm.elements.store.value = state.filters.store === 'ALL' ? '' : state.filters.store;
  el.productDialog.showModal();
}

function addManualProduct(event) {
  event.preventDefault();
  const form = Object.fromEntries(new FormData(el.productForm).entries());
  form.rupture = el.productForm.elements.rupture.checked;
  const normalizedCode = String(form.code || '').trim().toUpperCase();
  if (!normalizedCode || !form.description?.trim()) {
    showToast('Informe código e descrição do produto.', 'warning');
    return;
  }
  if (state.products.some((product) => String(product.code || '').trim().toUpperCase() === normalizedCode)) {
    showToast('Esse código já existe no painel. Use o vínculo de estoque se for um produto de produção.', 'warning');
    return;
  }
  state.manualEntries.push(form);
  saveStoredManualEntries();
  const product = createManualProduct(form, state.params);
  state.products = [...state.products, product];
  state.history = recordManualProduct(state.history, form, product);
  state.categories = [...new Set(state.products.map((item) => item.sector || item.paste).filter(Boolean))].sort();
  state.stores = [...new Set(state.products.map((item) => item.store).filter(Boolean))].sort();
  syncStoreSelect();
  syncCategorySelect();
  syncProductLists();
  state.selectedIds.add(product.id);
  el.productDialog.close();
  renderAll();
  showToast('Produto cadastrado e incluído no painel de pedidos.', 'success');
}

function openMappingDialog() {
  syncProductLists();
  renderMappingList();
  el.mappingDialog.showModal();
}

function saveMapping(event) {
  event.preventDefault();
  const form = Object.fromEntries(new FormData(el.mappingForm).entries());
  const saleDescription = String(form.saleDescription || '').trim();
  const stockDescription = String(form.stockDescription || '').trim();
  if (!saleDescription || !stockDescription || normalizeTextForUi(saleDescription) === normalizeTextForUi(stockDescription)) {
    showToast('Informe descrições diferentes para venda e estoque-base.', 'warning');
    return;
  }
  const mapping = { ...form, saleCode: '', stockCode: '', saleDescription, stockDescription, origin: 'manual' };
  if (state.mappings.some((item) => normalizeTextForUi(item.saleDescription) === normalizeTextForUi(saleDescription) && normalizeTextForUi(item.stockDescription) === normalizeTextForUi(stockDescription))) {
    showToast('Esse vínculo já está salvo.', 'warning');
    return;
  }
  state.mappings.push(mapping);
  state.history = recordMapping(state.history, mapping);
  persistHistory();
  saveStoredMappings();
  el.mappingForm.reset();
  el.mappingDialog.close();
  if (state.currentFile) {
    showToast('Vínculo salvo. Recalculando a fila com estoque único…', 'success');
    handleFile(state.currentFile);
  } else {
    showToast('Vínculo salvo. Ele será aplicado na próxima importação.', 'success');
    renderMappingList();
  }
}

function persistDecisionState() {
  if (!state.currentSnapshotId) return;
  updateCurrentDecisions(state.history, state.currentSnapshotId, state.products, state.selectedIds, state.adjustments);
  persistHistory();
}

function selectItem(id, selected) {
  if (selected) state.selectedIds.add(id);
  else state.selectedIds.delete(id);
  persistDecisionState();
  updateKpis();
  renderTable();
  updatePriorityList();
}

function updateQuantity(id, rawValue) {
  const product = state.products.find((entry) => entry.id === id);
  if (!product) return;
  const max = Math.max(0, product.depotAvailable);
  let quantity = Math.max(0, Math.floor(Number(rawValue) || 0));
  if (quantity > max) {
    quantity = max;
    showToast(`Quantidade limitada ao saldo disponível do depósito (${formatNumber.format(max)} un.).`, 'warning');
  }
  state.adjustments.set(id, quantity);
  if (quantity > 0 && product.priority !== 'BLOCKED') state.selectedIds.add(id);
  if (quantity === 0) state.selectedIds.delete(id);
  persistDecisionState();
  updateKpis();
  renderTable();
  updatePriorityList();
}

function exportQueue() {
  const rows = exportRows(state.products, state.adjustments, state.selectedIds, state.filters.store);
  if (!rows.length) {
    showToast('Selecione itens com quantidade positiva antes de exportar.', 'warning');
    return;
  }
  state.history = recordExport(state.history, state.currentSnapshotId, state.products, state.selectedIds, state.adjustments, rows.length);
  persistHistory();
  const csv = makeCsv(rows);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  const day = new Date().toISOString().slice(0, 10);
  anchor.href = url;
  anchor.download = `pedido-loja-${day}.csv`;
  anchor.click();
  URL.revokeObjectURL(url);
  showToast(`${formatNumber.format(rows.length)} itens exportados para a fila de pedido.`, 'success');
}

function clearFilters() {
  state.filters.search = '';
  state.filters.priority = 'ALL';
  state.filters.availability = 'ALL';
  state.filters.category = 'ALL';
  el.searchInput.value = '';
  el.priorityFilter.value = 'ALL';
  el.availabilityFilter.value = 'ALL';
  el.categoryFilter.value = 'ALL';
  state.currentPage = 1;
  renderAll();
}

function quickFilter(filter) {
  if (filter === 'urgent') {
    state.filters.priority = 'ALL';
    state.filters.availability = 'ALL';
    el.priorityFilter.value = 'ALL';
    el.availabilityFilter.value = 'ALL';
    state.currentView = 'dashboard';
    state.currentPage = 1;
    document.querySelector('.filters-panel').scrollIntoView({ behavior: 'smooth', block: 'start' });
    renderAll();
  }
}

el.fileInput.addEventListener('change', (event) => handleFile(event.target.files));
el.uploadCard.addEventListener('dragover', (event) => { event.preventDefault(); el.uploadCard.classList.add('is-dragging'); });
el.uploadCard.addEventListener('dragleave', () => el.uploadCard.classList.remove('is-dragging'));
el.uploadCard.addEventListener('drop', (event) => { event.preventDefault(); el.uploadCard.classList.remove('is-dragging'); handleFile(event.dataTransfer.files); });
el.newImportButton.addEventListener('click', () => el.fileInput.click());
el.historyButton.addEventListener('click', openHistoryDialog);
el.closeHistoryButton.addEventListener('click', () => el.historyDialog.close());
el.closeHistoryActionButton.addEventListener('click', () => el.historyDialog.close());
el.importHistoryButton.addEventListener('click', () => el.historyFileInput.click());
el.historyFileInput.addEventListener('change', (event) => importHistoryFile(event.target.files?.[0]));
el.downloadHistoryButton.addEventListener('click', () => {
  downloadHistory(state.history);
  showToast('Histórico JSON gerado para download.', 'success');
});
el.clearHistoryButton.addEventListener('click', () => {
  if (!window.confirm('Apagar o histórico de aprendizado salvo neste navegador? O arquivo Excel não será afetado.')) return;
  state.history = clearHistory();
  state.currentSnapshotId = null;
  refreshHistoryDialog();
  showToast('Histórico local apagado.', 'success');
});
el.addProductButton.addEventListener('click', openProductDialog);
el.linkProductButton.addEventListener('click', openMappingDialog);
el.productForm.addEventListener('submit', addManualProduct);
el.closeProductButton.addEventListener('click', () => el.productDialog.close());
el.cancelProductButton.addEventListener('click', () => el.productDialog.close());
el.mappingForm.addEventListener('submit', saveMapping);
el.closeMappingButton.addEventListener('click', () => el.mappingDialog.close());
el.cancelMappingButton.addEventListener('click', () => el.mappingDialog.close());
el.mappingList.addEventListener('click', (event) => {
  const button = event.target.closest('.remove-mapping');
  if (!button) return;
  const index = Number(button.dataset.mappingIndex);
  if (!Number.isInteger(index)) return;
  const manualMappings = state.mappings;
  manualMappings.splice(index, 1);
  saveStoredMappings();
  renderMappingList();
  showToast('Vínculo manual removido. Reimporte a planilha para recalcular.', 'success');
});
el.recalculateButton.addEventListener('click', recalculate);
el.exportButton.addEventListener('click', exportQueue);
el.storeSelect.addEventListener('change', (event) => { state.filters.store = event.target.value; state.currentPage = 1; renderAll(); });
el.searchInput.addEventListener('input', (event) => { state.filters.search = event.target.value; state.currentPage = 1; renderTable(); });
el.priorityFilter.addEventListener('change', (event) => { state.filters.priority = event.target.value; state.currentPage = 1; renderTable(); });
el.availabilityFilter.addEventListener('change', (event) => { state.filters.availability = event.target.value; state.currentPage = 1; renderTable(); });
el.categoryFilter.addEventListener('change', (event) => { state.filters.category = event.target.value; state.currentPage = 1; renderTable(); });
el.clearFiltersButton.addEventListener('click', clearFilters);
el.prevPageButton.addEventListener('click', () => { state.currentPage -= 1; renderTable(); });
el.nextPageButton.addEventListener('click', () => { state.currentPage += 1; renderTable(); });
el.selectAllInput.addEventListener('change', (event) => {
  const start = (state.currentPage - 1) * PAGE_SIZE;
  const visible = getViewProducts().slice(start, start + PAGE_SIZE).filter((product) => productQuantity(product) > 0 && product.priority !== 'BLOCKED');
  visible.forEach((product) => selectItem(product.id, event.target.checked));
});
el.ordersTableBody.addEventListener('change', (event) => {
  const input = event.target;
  if (input.matches('[data-select-id]')) selectItem(input.dataset.selectId, input.checked);
  if (input.matches('[data-qty-id]')) updateQuantity(input.dataset.qtyId, input.value);
});
el.priorityList.addEventListener('change', (event) => {
  const input = event.target;
  if (input.matches('[data-select-id]')) selectItem(input.dataset.selectId, input.checked);
});
document.querySelectorAll('.nav-item').forEach((button) => button.addEventListener('click', () => {
  state.currentView = button.dataset.view;
  state.currentPage = 1;
  renderAll();
  if (!el.appView.classList.contains('hidden')) document.querySelector('.filters-panel').scrollIntoView({ behavior: 'smooth', block: 'start' });
}));
document.querySelectorAll('[data-quick-filter]').forEach((button) => button.addEventListener('click', () => quickFilter(button.dataset.quickFilter)));
el.openRulesButton.addEventListener('click', () => el.rulesDialog.showModal());
el.closeRulesButton.addEventListener('click', () => el.rulesDialog.close());
el.rulesDialog.addEventListener('click', (event) => { if (event.target === el.rulesDialog) el.rulesDialog.close(); });
