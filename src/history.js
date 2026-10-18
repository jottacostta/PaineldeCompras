const STORAGE_KEY = 'orquestra-ml-history-v1';
const HISTORY_FORMAT = 'orquestra-ml-history';
const HISTORY_SCHEMA_VERSION = 1;
const MAX_STORED_IMPORTS = 6;
const MAX_STORED_EVENTS = 12000;

function nowIso() {
  return new Date().toISOString();
}

function numberOrNull(value) {
  return value === null || value === undefined || !Number.isFinite(Number(value)) ? null : Number(value);
}

function sourceValue(product, field) {
  const sources = Object.values(product.sources ?? {});
  const values = sources.map((source) => source?.[field]).filter((value) => value !== null && value !== undefined && Number.isFinite(Number(value)));
  return values.length ? values[0] : null;
}

function productKey(product) {
  const description = String(product.description ?? '').trim().toUpperCase();
  const store = String(product.store ?? '').trim().toUpperCase();
  return `${description || 'produto sem descrição'}:${store}`;
}

function snapshotId() {
  return `imp_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function emptyHistory() {
  const timestamp = nowIso();
  return {
    format: HISTORY_FORMAT,
    schemaVersion: HISTORY_SCHEMA_VERSION,
    createdAt: timestamp,
    updatedAt: timestamp,
    app: {
      name: 'Orquestra — Gerenciador de Pedidos de Loja',
      version: '1.0.0',
      processing: 'browser-local',
      purpose: 'dataset de decisões de abastecimento para aprendizado supervisionado futuro',
    },
    imports: [],
    events: [],
    mappings: [],
    manualProducts: [],
  };
}

function normalizeHistory(value) {
  const base = emptyHistory();
  if (!value || value.format !== HISTORY_FORMAT) return base;
  return {
    ...base,
    ...value,
    app: { ...base.app, ...(value.app ?? {}) },
    imports: Array.isArray(value.imports) ? value.imports : [],
    events: Array.isArray(value.events) ? value.events : [],
    mappings: Array.isArray(value.mappings) ? value.mappings : [],
    manualProducts: Array.isArray(value.manualProducts) ? value.manualProducts : [],
  };
}

export function loadHistory() {
  try {
    return normalizeHistory(JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null'));
  } catch {
    return emptyHistory();
  }
}

function compactHistory(history) {
  const imports = history.imports.slice(-MAX_STORED_IMPORTS);
  const retainedIds = new Set(imports.map((entry) => entry.id));
  const events = history.events.filter((event) => !event.snapshotId || retainedIds.has(event.snapshotId)).slice(-MAX_STORED_EVENTS);
  return { ...history, imports, events, updatedAt: nowIso() };
}

export function saveHistory(history) {
  const compacted = compactHistory(normalizeHistory(history));
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(compacted));
    return { history: compacted, persisted: true, trimmed: compacted.imports.length < history.imports.length };
  } catch {
    const reduced = {
      ...compacted,
      imports: compacted.imports.slice(-2).map((entry) => ({
        ...entry,
        samples: entry.samples.filter((sample) => sample.features?.demandDaily > 0 || sample.recommendation?.suggestedQty > 0 || sample.features?.isRupture).slice(0, 5000),
      })),
    };
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(reduced));
      return { history: reduced, persisted: true, trimmed: true };
    } catch {
      return { history: compacted, persisted: false, trimmed: true };
    }
  }
}

function sampleFromProduct(product, params, importId) {
  const shortSource = product.sources?.SALDO ?? product.sources?.consolidado3meses ?? {};
  const consolidatedSource = product.sources?.consolidado3meses ?? {};
  const sample = {
    sampleId: `${importId}:${productKey(product)}`,
    snapshotId: importId,
    observedAt: nowIso(),
    entity: {
      productKey: productKey(product),
      code: product.code || null,
      description: product.description || null,
      store: product.store || null,
      category: product.paste || null,
      sector: product.sector || null,
      giro: product.giro || null,
      stockBaseCode: product.code || null,
      aliases: (product.aliases ?? []).map((alias) => ({ code: alias.code ?? null, description: alias.description ?? null })),
    },
    features: {
      salesShortUnits: numberOrNull(shortSource.sales15),
      sales15Units: numberOrNull(shortSource.sales15),
      sales30Units: numberOrNull(shortSource.sales30),
      salesLongUnits: numberOrNull(shortSource.sales30),
      salesThreeMonthsUnits: numberOrNull(consolidatedSource.salesQty),
      salesThreeMonthsDays: numberOrNull(consolidatedSource.daysSale),
      mdv30Calculated: numberOrNull(product.mdv30Calculated),
      mdvDaily: numberOrNull(product.demandDaily),
      mdvErpReference: numberOrNull(consolidatedSource.mdv),
      demandDaily: numberOrNull(product.demandDaily),
      demandSource: product.demandSource || null,
      storeStockUnits: numberOrNull(product.storeStock),
      inboundUnits: numberOrNull(product.inbound),
      pendingUnits: numberOrNull(product.inboundDetails?.pendente),
      transitUnits: numberOrNull(product.inboundDetails?.transito),
      depotAvailableUnits: numberOrNull(product.depotAvailable),
      depotStockUnits: numberOrNull(product.depotStock),
      depotBalanceUnits: numberOrNull(product.depotBalance),
      currentCoverageDays: numberOrNull(product.coverage),
      packUnits: numberOrNull(product.pack),
      targetDays: numberOrNull(params.targetDays),
      maxDays: numberOrNull(params.maxDays),
      criticalDays: numberOrNull(params.criticalDays ?? 2),
      shortPeriodDays: numberOrNull(params.shortPeriod),
      longPeriodDays: numberOrNull(params.longPeriod),
      isRupture: Boolean(product.isRupture),
      hasSaldo: Boolean(product.hasSaldo),
      hasConsolidated: Boolean(product.hasConsolidated),
      hasZerado: Boolean(product.hasZerado),
      baseMetadataMatched: Boolean(product.baseMatch),
      mappedAliasCount: (product.aliases ?? []).length,
      warningCount: (product.warnings ?? []).length,
      reconciliationSources: numberOrNull(product.reconciliation),
    },
    recommendation: {
      priority: product.priority || 'NO_ACTION',
      suggestedQty: numberOrNull(product.suggested) ?? 0,
      coverageAfterSuggestedDays: numberOrNull(product.coverageAfterOrder),
      shortageUnits: numberOrNull(product.shortage),
      reason: product.reason || null,
    },
    decision: {
      selected: false,
      finalQty: 0,
      adjustedByBuyer: false,
      exported: false,
      exportedAt: null,
    },
    outcome: {
      labelStatus: 'awaiting_observation',
      nextSnapshotId: null,
      nextObservedAt: null,
      nextDemandDaily: null,
      nextStoreStockUnits: null,
      nextCoverageDays: null,
      nextPriority: null,
      nextRupture: null,
      acceptedQty: null,
      actualSalesUnits: null,
      avariaUnits: null,
      stockoutAfterOrder: null,
      buyerComment: null,
    },
  };
  return sample;
}

function findSample(importEntry, product) {
  const key = productKey(product);
  return importEntry?.samples?.find((sample) => sample.entity?.productKey === key || sample.entity?.code === product.code) ?? null;
}

function linkPreviousObservation(history, currentImportId, products) {
  const previous = history.imports.at(-1);
  if (!previous || previous.id === currentImportId) return;
  const byKey = new Map(products.map((product) => [productKey(product), product]));
  previous.samples.forEach((sample) => {
    const current = byKey.get(sample.entity?.productKey);
    if (!current) return;
    sample.outcome.nextSnapshotId = currentImportId;
    sample.outcome.nextObservedAt = nowIso();
    sample.outcome.nextDemandDaily = numberOrNull(current.demandDaily);
    sample.outcome.nextStoreStockUnits = numberOrNull(current.storeStock);
    sample.outcome.nextCoverageDays = numberOrNull(current.coverage);
    sample.outcome.nextPriority = current.priority || null;
    sample.outcome.nextRupture = Boolean(current.isRupture);
    sample.outcome.labelStatus = 'partially_observed';
  });
}

export function recordImport(history, { fileName, importedAt, params, stats, mappings, products }) {
  const id = snapshotId();
  linkPreviousObservation(history, id, products);
  const entry = {
    id,
    importedAt: importedAt instanceof Date ? importedAt.toISOString() : String(importedAt || nowIso()),
    recordedAt: nowIso(),
    sourceFile: { name: fileName || null, stored: false, contentStored: false },
    parameters: { ...params },
    stats: { ...stats, productCount: products.length },
    mappingsApplied: (mappings ?? []).map((mapping) => ({
      origin: mapping.origin || 'manual',
      saleCode: mapping.saleCode ?? null,
      saleDescription: mapping.saleDescription ?? null,
      stockCode: mapping.stockCode ?? null,
      stockDescription: mapping.stockDescription ?? null,
    })),
    samples: products.map((product) => sampleFromProduct(product, params, id)),
  };
  history.imports.push(entry);
  history.events.push({ type: 'import', at: entry.recordedAt, snapshotId: id, productCount: products.length, fileName: fileName || null });
  history.updatedAt = nowIso();
  return { history, snapshotId: id };
}

export function updateCurrentDecisions(history, snapshotIdValue, products, selectedIds, adjustments) {
  const current = history.imports.find((entry) => entry.id === snapshotIdValue);
  if (!current) return history;
  products.forEach((product) => {
    const sample = findSample(current, product);
    if (!sample) return;
    const adjusted = adjustments.get(product.id);
    const finalQty = adjusted === undefined ? product.suggested : adjusted;
    sample.decision.selected = selectedIds.has(product.id);
    sample.decision.finalQty = numberOrNull(finalQty) ?? 0;
    sample.decision.adjustedByBuyer = adjusted !== undefined && Number(adjusted) !== Number(product.suggested);
  });
  history.updatedAt = nowIso();
  return history;
}

export function recordExport(history, snapshotIdValue, products, selectedIds, adjustments, rowCount) {
  updateCurrentDecisions(history, snapshotIdValue, products, selectedIds, adjustments);
  const current = history.imports.find((entry) => entry.id === snapshotIdValue);
  const exportedAt = nowIso();
  current?.samples?.forEach((sample) => {
    if (sample.decision.selected && sample.decision.finalQty > 0) {
      sample.decision.exported = true;
      sample.decision.exportedAt = exportedAt;
    }
  });
  history.events.push({ type: 'export_csv', at: exportedAt, snapshotId: snapshotIdValue, rowCount });
  history.updatedAt = exportedAt;
  return history;
}

export function recordEvent(history, event) {
  history.events.push({ ...event, at: event.at || nowIso() });
  history.updatedAt = nowIso();
  return history;
}

export function recordManualProduct(history, form, product) {
  history.manualProducts.push({
    recordedAt: nowIso(),
    code: form.code || null,
    description: form.description || null,
    store: form.store || null,
    category: form.paste || null,
    productId: product.id,
  });
  return recordEvent(history, { type: 'manual_product', productId: product.id, code: form.code || null });
}

export function recordMapping(history, mapping) {
  history.mappings.push({ ...mapping, recordedAt: nowIso() });
  return recordEvent(history, { type: 'mapping_created', saleCode: mapping.saleCode, stockCode: mapping.stockCode });
}

export function historySummary(history) {
  const samples = history.imports.reduce((total, entry) => total + (entry.samples?.length || 0), 0);
  const exports = history.events.filter((event) => event.type === 'export_csv').length;
  const observed = history.imports.reduce((total, entry) => total + (entry.samples ?? []).filter((sample) => sample.outcome?.labelStatus !== 'awaiting_observation').length, 0);
  return {
    imports: history.imports.length,
    samples,
    observedSamples: observed,
    exports,
    mappings: history.mappings.length,
    manualProducts: history.manualProducts.length,
    lastUpdatedAt: history.updatedAt,
  };
}

export function historyJson(history) {
  const payload = normalizeHistory(history);
  payload.exportedAt = nowIso();
  payload.summary = historySummary(payload);
  return JSON.stringify(payload, null, 2);
}

export function mergeImportedHistory(currentHistory, jsonText) {
  let parsed;
  try {
    parsed = JSON.parse(jsonText);
  } catch {
    throw new Error('O arquivo não é um JSON válido.');
  }
  if (!parsed || parsed.format !== HISTORY_FORMAT) {
    throw new Error('Este JSON não é um histórico do Orquestra compatível.');
  }
  if (Number(parsed.schemaVersion) > HISTORY_SCHEMA_VERSION) {
    throw new Error('Este histórico foi criado por uma versão mais nova do sistema.');
  }

  const incoming = normalizeHistory(parsed);
  const merged = normalizeHistory(currentHistory);
  const existingImports = new Map(merged.imports.map((entry) => [entry.id, entry]));
  incoming.imports.forEach((entry) => {
    if (!existingImports.has(entry.id)) existingImports.set(entry.id, entry);
  });
  merged.imports = [...existingImports.values()].sort((a, b) => String(a.importedAt).localeCompare(String(b.importedAt)));

  const eventKeys = new Set(merged.events.map((event) => `${event.type}:${event.at}:${event.snapshotId || ''}`));
  incoming.events.forEach((event) => {
    const key = `${event.type}:${event.at}:${event.snapshotId || ''}`;
    if (!eventKeys.has(key)) {
      merged.events.push(event);
      eventKeys.add(key);
    }
  });
  const mappingKeys = new Set(merged.mappings.map((mapping) => `${mapping.saleCode}:${mapping.stockCode}`));
  incoming.mappings.forEach((mapping) => {
    const key = `${mapping.saleCode}:${mapping.stockCode}`;
    if (!mappingKeys.has(key)) {
      merged.mappings.push(mapping);
      mappingKeys.add(key);
    }
  });
  const manualKeys = new Set(merged.manualProducts.map((product) => `${product.code}:${product.store}`));
  incoming.manualProducts.forEach((product) => {
    const key = `${product.code}:${product.store}`;
    if (!manualKeys.has(key)) {
      merged.manualProducts.push(product);
      manualKeys.add(key);
    }
  });
  merged.events.push({
    type: 'history_imported',
    at: nowIso(),
    importedSnapshots: incoming.imports.length,
    importedSamples: incoming.imports.reduce((total, entry) => total + (entry.samples?.length || 0), 0),
  });
  merged.updatedAt = nowIso();
  return merged;
}

export function downloadHistory(history) {
  const blob = new Blob([historyJson(history)], { type: 'application/json;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  const day = new Date().toISOString().slice(0, 10);
  anchor.href = url;
  anchor.download = `historico-aprendizado-orquestra-${day}.json`;
  anchor.click();
  URL.revokeObjectURL(url);
}

export function clearHistory() {
  localStorage.removeItem(STORAGE_KEY);
  return emptyHistory();
}
