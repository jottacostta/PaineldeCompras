import * as XLSX from 'xlsx';

const REQUIRED_SHEETS = ['SALDO', 'consolidado3meses', 'Zerado'];
const DEFAULT_STORE = 'Loja do arquivo';

export const PRIORITY_META = {
  P0: { label: 'P0 · Ruptura com venda', tone: 'danger', rank: 0 },
  P1: { label: 'P1 · Risco imediato', tone: 'danger', rank: 1 },
  P2: { label: 'P2 · Abaixo do alvo', tone: 'warning', rank: 2 },
  P3: { label: 'P3 · Giro alto', tone: 'teal', rank: 3 },
  BLOCKED: { label: 'Bloqueado', tone: 'dark', rank: 4 },
  NO_ACTION: { label: 'Sem pedido', tone: 'neutral', rank: 5 },
};

const FIELD_ALIASES = {
  code: ['codigo', 'chave', 'codigoproduto', 'cod'],
  description: ['descricaodoproduto', 'descricao', 'produto'],
  distribution: ['distrib'],
  centralStock: ['estcentral', 'estoquecentral'],
  balance: ['saldo'],
  storeStock: ['estloja', 'estoqueun'],
  sales15: ['vda15d', 'venda15d'],
  sales30: ['vda30d', 'venda30d'],
  coverage: ['cobertura'],
  storeOrder: ['pedidoloja'],
  orderedQty: ['qtdpedida', 'qtdpedido'],
  paste: ['pasta'],
  salesValue: ['valorvenda'],
  salesQty: ['qtdevenda'],
  daysSale: ['diasvenda'],
  mdv: ['mdv', 'mdv107'],
  pending: ['qtdepedpend', 'qtdepedpendente'],
  transit: ['qtdeesttransito', 'transito'],
  depositStock: ['estdepositoun', 'estoquedepositoun'],
  lastSale: ['dtultvenda', 'ultimavenda'],
  store: ['loja', 'filial', 'unidade'],
};

function normalizeText(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function normalizeHeader(value) {
  return normalizeText(value).replace(/\s/g, '').toLowerCase();
}

function rawText(value) {
  return String(value ?? '').trim();
}

function codeExact(value) {
  return rawText(value).toUpperCase().replace(/\s+/g, '');
}

function codeNumeric(value) {
  const digits = codeExact(value).replace(/\D/g, '');
  if (!digits) return '';
  return digits.replace(/^0+(?=\d)/, '');
}

function toNumberOrNull(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const clean = String(value)
    .replace(/R\$/gi, '')
    .replace(/\s/g, '')
    .replace(/\.(?=\d{3}(?:\D|$))/g, '')
    .replace(',', '.');
  const parsed = Number(clean);
  return Number.isFinite(parsed) ? parsed : null;
}

function toNonNegative(value) {
  const numeric = toNumberOrNull(value);
  return numeric === null ? null : Math.max(0, numeric);
}

function parsePack(distribution) {
  const match = String(distribution ?? '').match(/\/(\d+(?:[.,]\d+)?)/);
  const value = match ? Number(match[1].replace(',', '.')) : 1;
  return Number.isFinite(value) && value > 0 ? value : 1;
}

function parseErpQuantity(value, pack = 1) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return Math.max(0, value);
  const text = String(value).trim().toLowerCase();
  if (!text || text === '-' || text === '_*_' || text === 'nd') return null;
  const matches = [...text.matchAll(/(-?\d+(?:[.,]\d+)?)/g)].map((match) => Number(match[1].replace(',', '.')));
  if (!matches.length || matches.some((valueNumber) => !Number.isFinite(valueNumber))) return null;
  if (matches.length === 1) return Math.max(0, matches[0]);
  // ERP commonly writes "9+3un" for 9 cartons plus 3 loose units.
  return Math.max(0, matches[0]) * pack + Math.max(0, matches[1]);
}

function getField(row, aliases) {
  const normalized = Object.fromEntries(Object.entries(row).map(([key, value]) => [normalizeHeader(key), value]));
  for (const alias of aliases) {
    if (Object.hasOwn(normalized, alias)) return normalized[alias];
  }
  return null;
}

function sheetRows(workbook, sheetName) {
  const worksheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(worksheet, { defval: null, raw: true, dateNF: 'dd/mm/yyyy' });
}

function normalizedRow(row, source) {
  const code = getField(row, FIELD_ALIASES.code);
  const description = getField(row, FIELD_ALIASES.description);
  const distribution = getField(row, FIELD_ALIASES.distribution);
  const pack = parsePack(distribution);
  const storeValue = getField(row, FIELD_ALIASES.store);

  return {
    source,
    raw: row,
    code: rawText(code),
    codeExact: codeExact(code),
    codeNumeric: codeNumeric(code),
    description: rawText(description),
    descriptionKey: normalizeText(description),
    store: rawText(storeValue) || DEFAULT_STORE,
    pack,
    distribution: rawText(distribution),
    centralStock: toNonNegative(getField(row, FIELD_ALIASES.centralStock)),
    balance: toNonNegative(getField(row, FIELD_ALIASES.balance)),
    storeStock: parseErpQuantity(getField(row, FIELD_ALIASES.storeStock), pack),
    sales15: parseErpQuantity(getField(row, FIELD_ALIASES.sales15), pack),
    sales30: parseErpQuantity(getField(row, FIELD_ALIASES.sales30), pack),
    coverage: toNonNegative(getField(row, FIELD_ALIASES.coverage)),
    storeOrder: rawText(getField(row, FIELD_ALIASES.storeOrder)),
    orderedQty: parseErpQuantity(getField(row, FIELD_ALIASES.orderedQty), pack),
    paste: rawText(getField(row, FIELD_ALIASES.paste)) || 'Sem pasta',
    salesValue: toNonNegative(getField(row, FIELD_ALIASES.salesValue)),
    salesQty: toNonNegative(getField(row, FIELD_ALIASES.salesQty)),
    daysSale: toNonNegative(getField(row, FIELD_ALIASES.daysSale)),
    mdv: toNonNegative(getField(row, FIELD_ALIASES.mdv)),
    pending: toNonNegative(getField(row, FIELD_ALIASES.pending)),
    transit: toNonNegative(getField(row, FIELD_ALIASES.transit)),
    depositStock: toNonNegative(getField(row, FIELD_ALIASES.depositStock)),
    lastSale: getField(row, FIELD_ALIASES.lastSale),
  };
}

function valuesEqual(a, b) {
  return a !== null && b !== null && Math.abs(a - b) < 0.0001;
}

function uniqueCandidate(candidates) {
  const unique = [...new Set(candidates)];
  return unique.length === 1 ? unique[0] : null;
}

function buildIndex(rows) {
  const exact = new Map();
  const numeric = new Map();
  const description = new Map();
  rows.forEach((row, index) => {
    if (row.codeExact) {
      if (!exact.has(row.codeExact)) exact.set(row.codeExact, []);
      exact.get(row.codeExact).push(index);
    }
    if (row.codeNumeric) {
      if (!numeric.has(row.codeNumeric)) numeric.set(row.codeNumeric, []);
      numeric.get(row.codeNumeric).push(index);
    }
    if (row.descriptionKey) {
      if (!description.has(row.descriptionKey)) description.set(row.descriptionKey, []);
      description.get(row.descriptionKey).push(index);
    }
  });
  return { exact, numeric, description };
}

function findMatch(product, incoming, sourceIndex) {
  const index = sourceIndex[incoming.source];
  if (!index) return { record: null, method: 'new' };
  const byExact = incoming.codeExact ? uniqueCandidate(index.exact.get(incoming.codeExact) ?? []) : null;
  if (byExact !== null) return { record: product[byExact], method: 'código exato' };

  const byNumeric = incoming.codeNumeric ? uniqueCandidate(index.numeric.get(incoming.codeNumeric) ?? []) : null;
  if (byNumeric !== null) return { record: product[byNumeric], method: 'código numérico' };

  const byDescription = incoming.descriptionKey
    ? uniqueCandidate(index.description.get(incoming.descriptionKey) ?? [])
    : null;
  if (byDescription !== null) return { record: product[byDescription], method: 'descrição' };

  return { record: null, method: 'new' };
}

function recordKey(row, order) {
  return row.codeExact || row.descriptionKey || `${row.source}-${order}`;
}

function combineSources(saldoRows, consolidatedRows, zeradoRows) {
  const products = [];
  const sourceIndex = {
    SALDO: { exact: new Map(), numeric: new Map(), description: new Map() },
    consolidado3meses: { exact: new Map(), numeric: new Map(), description: new Map() },
    Zerado: { exact: new Map(), numeric: new Map(), description: new Map() },
  };
  const allRows = [
    ...saldoRows.map((row) => ({ ...row, source: 'SALDO' })),
    ...consolidatedRows.map((row) => ({ ...row, source: 'consolidado3meses' })),
    ...zeradoRows.map((row) => ({ ...row, source: 'Zerado' })),
  ];

  const rebuiltIndex = () => {
    for (const source of Object.keys(sourceIndex)) {
      sourceIndex[source] = { exact: new Map(), numeric: new Map(), description: new Map() };
    }
    products.forEach((product, index) => {
      Object.keys(product.sources).forEach((source) => {
        const row = product.sources[source];
        if (row.codeExact) {
          if (!sourceIndex[source].exact.has(row.codeExact)) sourceIndex[source].exact.set(row.codeExact, []);
          sourceIndex[source].exact.get(row.codeExact).push(index);
        }
        if (row.codeNumeric) {
          if (!sourceIndex[source].numeric.has(row.codeNumeric)) sourceIndex[source].numeric.set(row.codeNumeric, []);
          sourceIndex[source].numeric.get(row.codeNumeric).push(index);
        }
        if (row.descriptionKey) {
          if (!sourceIndex[source].description.has(row.descriptionKey)) sourceIndex[source].description.set(row.descriptionKey, []);
          sourceIndex[source].description.get(row.descriptionKey).push(index);
        }
      });
    });
  };

  allRows.forEach((row, index) => {
    const matches = Object.values(sourceIndex)
      .map((indexBySource) => {
        const exact = row.codeExact ? uniqueCandidate(indexBySource.exact.get(row.codeExact) ?? []) : null;
        if (exact !== null) return { index: exact, method: 'código exato' };
        const numeric = row.codeNumeric ? uniqueCandidate(indexBySource.numeric.get(row.codeNumeric) ?? []) : null;
        if (numeric !== null) return { index: numeric, method: 'código numérico' };
        const description = row.descriptionKey ? uniqueCandidate(indexBySource.description.get(row.descriptionKey) ?? []) : null;
        if (description !== null) return { index: description, method: 'descrição' };
        return null;
      })
      .filter(Boolean);
    const matchedIndexes = [...new Set(matches.map((match) => match.index))];

    let product;
    if (matchedIndexes.length === 1) {
      product = products[matchedIndexes[0]];
      product.matchMethods.push(matches[0].method);
      if (matches.some((match) => match.method === 'descrição')) product.warnings.push('Conciliação feita pela descrição: confira código e embalagem.');
    } else {
      product = {
        id: recordKey(row, index),
        code: row.code || 'Sem código',
        description: row.description || 'Produto sem descrição',
        store: row.store || DEFAULT_STORE,
        paste: row.paste || 'Sem pasta',
        sources: {},
        matchMethods: [],
        warnings: matchedIndexes.length > 1 ? ['Possível conciliação ambígua: validar cadastro.'] : [],
      };
      products.push(product);
    }

    if (product.sources[row.source]) {
      product.warnings.push(`Há mais de uma linha em ${row.source}; valores não foram somados automaticamente.`);
    } else {
      product.sources[row.source] = row;
    }

    if (product.code === 'Sem código' && row.code) product.code = row.code;
    if (product.description === 'Produto sem descrição' && row.description) product.description = row.description;
    if (product.paste === 'Sem pasta' && row.paste !== 'Sem pasta') product.paste = row.paste;
    if (product.store === DEFAULT_STORE && row.store !== DEFAULT_STORE) product.store = row.store;
    rebuiltIndex();
  });

  return products;
}

function weightedDemand({ sales15, sales30, mdv, params }) {
  const shortPeriod = Math.max(1, Number(params?.shortPeriod) || 15);
  const longPeriod = Math.max(shortPeriod, Number(params?.longPeriod) || 30);
  const components = [
    { value: sales15 === null ? null : sales15 / shortPeriod, weight: 0.5 },
    { value: sales30 === null ? null : sales30 / longPeriod, weight: 0.3 },
    { value: mdv, weight: 0.2 },
  ].filter((component) => component.value !== null && Number.isFinite(component.value));

  if (!components.length) return { demand: 0, source: 'Sem histórico de venda' };
  const totalWeight = components.reduce((sum, component) => sum + component.weight, 0);
  const blended = components.reduce((sum, component) => sum + component.value * component.weight, 0) / totalWeight;
  const recent = components[0]?.value ?? 0;
  const demand = recent > blended * 1.35 ? Math.min(recent, blended * 1.5) : blended;
  const source = components.length === 3 ? `${shortPeriod}D + ${longPeriod}D + MDV` : `${components.length} fonte${components.length > 1 ? 's' : ''} de venda`;
  return { demand: Math.max(0, demand), source };
}

function minConservative(values) {
  const known = values.filter((value) => value !== null && value !== undefined && Number.isFinite(value)).map((value) => Math.max(0, value));
  return known.length ? Math.min(...known) : 0;
}

function ratioDifference(values) {
  const known = values.filter((value) => value !== null && value !== undefined && Number.isFinite(value) && value >= 0);
  if (known.length < 2) return false;
  const min = Math.min(...known);
  const max = Math.max(...known);
  return max > 0 && (max - min) / max > 0.2;
}

function numberOrZero(value) {
  return value === null || value === undefined ? 0 : value;
}

function calculateProduct(product, params) {
  const saldo = product.sources.SALDO;
  const consolidated = product.sources.consolidado3meses;
  const zerado = product.sources.Zerado;
  const warningSet = new Set(product.warnings);
  const pack = saldo?.pack || zerado?.pack || 1;
  const demandInfo = weightedDemand({
    sales15: saldo?.sales15 ?? zerado?.sales15 ?? null,
    sales30: saldo?.sales30 ?? zerado?.sales30 ?? null,
    mdv: consolidated?.mdv ?? saldo?.mdv ?? null,
    params,
  });

  const storeStock = numberOrZero(consolidated?.storeStock ?? saldo?.storeStock ?? zerado?.storeStock);
  const pending = Math.max(0, numberOrZero(consolidated?.pending), numberOrZero(consolidated?.transit), numberOrZero(saldo?.orderedQty));
  const inboundDetails = {
    pendente: numberOrZero(consolidated?.pending),
    transito: numberOrZero(consolidated?.transit),
    pedidoLoja: numberOrZero(saldo?.orderedQty),
  };
  const depositMeasures = [saldo?.centralStock ?? null, saldo?.balance ?? null, consolidated?.depositStock ?? null];
  const depotAvailable = minConservative(depositMeasures);
  const depotStock = saldo?.centralStock ?? consolidated?.depositStock ?? null;
  const depotBalance = saldo?.balance ?? null;
  const projected = storeStock + pending;
  const coverage = demandInfo.demand > 0 ? projected / demandInfo.demand : (consolidated?.coverage ?? saldo?.coverage ?? null);
  const needed = demandInfo.demand > 0 ? Math.max(0, Math.ceil(demandInfo.demand * params.targetDays - projected)) : 0;
  const maxQuantity = demandInfo.demand > 0 ? Math.max(0, Math.floor(demandInfo.demand * params.maxDays - projected)) : 0;
  const unrounded = Math.min(needed, maxQuantity, depotAvailable);
  const availableUnits = Math.max(0, Math.floor(unrounded));
  const roundedPack = pack > 1 ? Math.floor(availableUnits / pack) * pack : availableUnits;
  let suggested = demandInfo.demand > 0 ? (roundedPack > 0 ? roundedPack : availableUnits) : 0;
  const isRupture = Boolean(zerado) || (storeStock <= 0 && demandInfo.demand > 0);
  const hasDemand = demandInfo.demand > 0.02;

  if (!saldo) warningSet.add('Sem registro na aba SALDO: disponibilidade do depósito pode estar incompleta.');
  if (!consolidated) warningSet.add('Sem histórico no consolidado: demanda pode estar incompleta.');
  if (!product.code || product.code === 'Sem código') warningSet.add('Produto sem código: não liberar pedido automático.');
  if (ratioDifference(depositMeasures)) warningSet.add('Divergência entre estoque central, saldo e/ou depósito: usado menor valor como limite.');
  if (consolidated && saldo && coverage !== null && consolidated.coverage !== null && Math.abs(coverage - consolidated.coverage) > Math.max(4, consolidated.coverage * 0.5)) warningSet.add('Cobertura recalculada diverge da informada no consolidado.');
  if (demandInfo.demand === 0 && Boolean(zerado)) warningSet.add('Ruptura sem venda conhecida: item mantido para revisão, sem pedido automático.');
  if (pending > 0) warningSet.add(`Entrada considerada de ${pending.toFixed(0)} un. (maior valor entre pendência/trânsito/pedido).`);
  if (pack > 1 && availableUnits > 0 && roundedPack === 0) warningSet.add(`Sugestão menor que o múltiplo de ${pack} un.; confirmar se o depósito pode separar unidades.`);

  let priority = 'NO_ACTION';
  let reason = 'Cobertura suficiente ou sem demanda para pedido automático.';
  if (hasDemand && depotAvailable <= 0 && (isRupture || (coverage !== null && coverage < params.targetDays))) {
    priority = 'BLOCKED';
    reason = 'Há demanda, mas não existe saldo liberado no depósito para atender a loja.';
  } else if (hasDemand && isRupture && depotAvailable > 0) {
    priority = 'P0';
    reason = 'Ruptura com venda: atender primeiro enquanto há saldo no depósito.';
  } else if (hasDemand && coverage !== null && coverage <= 2 && suggested > 0) {
    priority = 'P1';
    reason = 'Cobertura projetada de até 2 dias: risco imediato de falta.';
  } else if (hasDemand && coverage !== null && coverage < params.targetDays && suggested > 0) {
    priority = 'P2';
    reason = `Cobertura abaixo da meta de ${params.targetDays} dias.`;
  } else if (hasDemand && suggested > 0) {
    priority = 'P3';
    reason = 'Demanda ativa e oportunidade de reposição no próximo ciclo.';
  }

  if (warningSet.has('Produto sem código: não liberar pedido automático.')) {
    priority = 'NO_ACTION';
    reason = 'Cadastro incompleto: validar código antes de criar pedido.';
  }

  const finalSuggested = priority === 'NO_ACTION' || priority === 'BLOCKED' ? 0 : suggested;
  const projectedAfterOrder = demandInfo.demand > 0 ? (projected + finalSuggested) / demandInfo.demand : null;
  const reconciliation = [saldo, consolidated].filter(Boolean).length;

  return {
    ...product,
    pack,
    isRupture,
    demandDaily: demandInfo.demand,
    demandSource: demandInfo.source,
    salesValue: numberOrZero(consolidated?.salesValue),
    salesQty: numberOrZero(consolidated?.salesQty),
    storeStock,
    inbound: pending,
    inboundDetails,
    depotStock,
    depotBalance,
    depotAvailable,
    coverage,
    coverageAfterOrder: projectedAfterOrder,
    suggested: finalSuggested,
    originalSuggested: finalSuggested,
    shortage: Math.max(0, needed - depotAvailable),
    priority,
    reason,
    warnings: [...warningSet],
    reconciliation,
    hasSaldo: Boolean(saldo),
    hasConsolidated: Boolean(consolidated),
    hasZerado: Boolean(zerado),
    lastSale: consolidated?.lastSale ?? null,
  };
}

function validateWorkbook(workbook) {
  const names = workbook.SheetNames.map((name) => name.trim().toLocaleLowerCase('pt-BR'));
  const missing = REQUIRED_SHEETS.filter((required) => !names.includes(required.toLocaleLowerCase('pt-BR')));
  if (missing.length) {
    throw new Error(`Não encontramos a(s) aba(s) obrigatória(s): ${missing.join(', ')}.`);
  }
}

function actualSheetName(workbook, expected) {
  return workbook.SheetNames.find((name) => name.trim().toLocaleLowerCase('pt-BR') === expected.toLocaleLowerCase('pt-BR'));
}

function productKeys(product) {
  const codes = [product.code, ...Object.values(product.sources ?? {}).map((source) => source.code)];
  const descriptions = [product.description, ...Object.values(product.sources ?? {}).map((source) => source.description)];
  return {
    exactCodes: new Set(codes.map(codeExact).filter(Boolean)),
    numericCodes: new Set(codes.map(codeNumeric).filter(Boolean)),
    descriptions: new Set(descriptions.map((description) => normalizeText(description)).filter(Boolean)),
  };
}

function findProductByMapping(products, code, description) {
  const exact = codeExact(code);
  const numeric = codeNumeric(code);
  const descriptionKey = normalizeText(description);
  const exactMatches = products.filter((product) => productKeys(product).exactCodes.has(exact));
  if (exact && exactMatches.length === 1) return exactMatches[0];
  const numericMatches = products.filter((product) => numeric && productKeys(product).numericCodes.has(numeric));
  if (numeric && numericMatches.length === 1) return numericMatches[0];
  const descriptionMatches = products.filter((product) => descriptionKey && productKeys(product).descriptions.has(descriptionKey));
  return descriptionMatches.length === 1 ? descriptionMatches[0] : null;
}

function addNullable(base, extra) {
  if (base === null || base === undefined) return extra ?? null;
  if (extra === null || extra === undefined) return base;
  return base + extra;
}

function mergeSalesRow(base, extra) {
  return {
    ...base,
    sales15: addNullable(base.sales15, extra.sales15),
    sales30: addNullable(base.sales30, extra.sales30),
    salesValue: addNullable(base.salesValue, extra.salesValue),
    salesQty: addNullable(base.salesQty, extra.salesQty),
    mdv: addNullable(base.mdv, extra.mdv),
    daysSale: Math.max(numberOrZero(base.daysSale), numberOrZero(extra.daysSale)),
    pending: Math.max(numberOrZero(base.pending), numberOrZero(extra.pending)),
    transit: Math.max(numberOrZero(base.transit), numberOrZero(extra.transit)),
    lastSale: base.lastSale && extra.lastSale ? (new Date(base.lastSale) >= new Date(extra.lastSale) ? base.lastSale : extra.lastSale) : (base.lastSale ?? extra.lastSale ?? null),
  };
}

function applyProductMappings(products, mappings) {
  const applied = [];
  for (const mapping of mappings ?? []) {
    const saleProduct = findProductByMapping(products, mapping.saleCode, mapping.saleDescription);
    const stockProduct = findProductByMapping(products, mapping.stockCode, mapping.stockDescription);
    if (!saleProduct || !stockProduct || saleProduct === stockProduct) continue;

    const saleDescription = saleProduct.description;
    const stockDescription = stockProduct.description;
    const saleConsolidated = saleProduct.sources.consolidado3meses;
    if (saleConsolidated) {
      if (stockProduct.sources.consolidado3meses) stockProduct.sources.consolidado3meses = mergeSalesRow(stockProduct.sources.consolidado3meses, saleConsolidated);
      else stockProduct.sources.consolidado3meses = saleConsolidated;
    }
    if (saleProduct.sources.Zerado) {
      stockProduct.warnings.push(`O produto de venda ${saleDescription} está zerado; o alerta foi mantido sem duplicar o estoque-base.`);
    }
    if (saleProduct.sources.SALDO) {
      stockProduct.warnings.push(`O SALDO de ${saleDescription} foi ignorado no estoque porque o vínculo aponta para ${stockDescription}.`);
    }
    stockProduct.aliases = [...(stockProduct.aliases ?? []), { code: saleProduct.code, description: saleDescription }];
    stockProduct.warnings.push(`Estoque único: a venda de ${saleDescription} foi incorporada a ${stockDescription}.`);
    const index = products.indexOf(saleProduct);
    if (index >= 0) products.splice(index, 1);
    applied.push({ ...mapping, saleDescription, stockDescription });
  }
  return applied;
}

function panelMappings(workbook) {
  const panelName = actualSheetName(workbook, 'PAINEL');
  if (!panelName) return [];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[panelName], { range: 1, defval: null, raw: true });
  return rows.map((row) => ({
    saleCode: row.Chave,
    saleDescription: row.Descrição,
    stockCode: row.Host,
    stockDescription: row.Pai ?? row['Descrição Pai'] ?? row['Produto Pai'] ?? '',
    origin: 'PAINEL',
  })).filter((mapping) => rawText(mapping.saleCode) && rawText(mapping.stockCode) && codeExact(mapping.saleCode) !== codeExact(mapping.stockCode));
}

export function createManualProduct(form, params) {
  const code = rawText(form.code);
  const description = rawText(form.description) || 'Produto cadastrado manualmente';
  const pack = Math.max(1, Number(form.pack) || 1);
  const base = {
    id: `manual:${codeExact(code) || normalizeText(description)}:${normalizeText(form.store || DEFAULT_STORE)}`,
    code: code || 'Sem código',
    description,
    store: rawText(form.store) || DEFAULT_STORE,
    paste: rawText(form.paste) || 'Cadastro manual',
    sources: {
      SALDO: {
        source: 'SALDO', code, codeExact: codeExact(code), codeNumeric: codeNumeric(code), description, descriptionKey: normalizeText(description),
        store: rawText(form.store) || DEFAULT_STORE, pack, distribution: `cx/${pack}`, centralStock: toNonNegative(form.depotStock), balance: toNonNegative(form.balance ?? form.depotStock),
        storeStock: toNonNegative(form.storeStock), sales15: toNonNegative(form.sales15), sales30: toNonNegative(form.sales30), coverage: null, storeOrder: '', orderedQty: toNonNegative(form.orderedQty), paste: rawText(form.paste) || 'Cadastro manual', salesValue: 0, salesQty: 0, daysSale: null, mdv: toNonNegative(form.mdv), pending: toNonNegative(form.pending), transit: toNonNegative(form.transit), depositStock: null, lastSale: null,
      },
      consolidado3meses: {
        source: 'consolidado3meses', code, codeExact: codeExact(code), codeNumeric: codeNumeric(code), description, descriptionKey: normalizeText(description),
        store: rawText(form.store) || DEFAULT_STORE, pack, distribution: `cx/${pack}`, centralStock: null, balance: null, storeStock: toNonNegative(form.storeStock), sales15: toNonNegative(form.sales15), sales30: toNonNegative(form.sales30), coverage: null, storeOrder: '', orderedQty: null, paste: rawText(form.paste) || 'Cadastro manual', salesValue: 0, salesQty: addNullable(toNonNegative(form.sales15), toNonNegative(form.sales30)), daysSale: null, mdv: toNonNegative(form.mdv), pending: toNonNegative(form.pending), transit: toNonNegative(form.transit), depositStock: toNonNegative(form.depotStock), lastSale: null,
      },
    },
    matchMethods: ['cadastro manual'],
    warnings: ['Produto cadastrado manualmente; confira o vínculo com o ERP na próxima importação.'],
    manualForm: { ...form, code, description, pack },
  };
  if (form.rupture) base.sources.Zerado = { ...base.sources.SALDO, source: 'Zerado' };
  return calculateProduct(base, params);
}

export async function importWorkbook(file, params, mappings = []) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });
  validateWorkbook(workbook);

  const saldoRows = sheetRows(workbook, actualSheetName(workbook, 'SALDO')).map((row) => normalizedRow(row, 'SALDO'));
  const consolidatedRows = sheetRows(workbook, actualSheetName(workbook, 'consolidado3meses')).map((row) => normalizedRow(row, 'consolidado3meses'));
  const zeradoRows = sheetRows(workbook, actualSheetName(workbook, 'Zerado')).map((row) => normalizedRow(row, 'Zerado'));

  const rawProducts = combineSources(saldoRows, consolidatedRows, zeradoRows);
  const importedMappings = panelMappings(workbook);
  const appliedMappings = applyProductMappings(rawProducts, [...importedMappings, ...(mappings ?? [])]);
  const products = rawProducts
    .map((product) => calculateProduct(product, params))
    .sort((a, b) => {
      const priorityDifference = PRIORITY_META[a.priority].rank - PRIORITY_META[b.priority].rank;
      if (priorityDifference !== 0) return priorityDifference;
      const coverageA = a.coverage === null ? Number.POSITIVE_INFINITY : a.coverage;
      const coverageB = b.coverage === null ? Number.POSITIVE_INFINITY : b.coverage;
      if (coverageA !== coverageB) return coverageA - coverageB;
      return b.demandDaily - a.demandDaily;
    });

  const stores = [...new Set(products.map((product) => product.store).filter(Boolean))].sort();
  const categories = [...new Set(products.map((product) => product.paste).filter(Boolean))].sort();
  const stats = {
    sourceRows: { saldo: saldoRows.length, consolidated: consolidatedRows.length, zerado: zeradoRows.length },
    unmatched: products.filter((product) => !product.hasSaldo || !product.hasConsolidated).length,
    matched: products.filter((product) => product.hasSaldo && product.hasConsolidated).length,
    descriptionMatches: products.filter((product) => product.matchMethods.includes('descrição')).length,
    mappedProducts: appliedMappings.length,
  };

  return { products, stores: stores.length ? stores : [DEFAULT_STORE], categories, stats, mappings: appliedMappings, fileName: file.name, importedAt: new Date() };
}

export function recalculateProducts(products, params) {
  return products.map((product) => calculateProduct({
    id: product.id,
    code: product.code,
    description: product.description,
    store: product.store,
    paste: product.paste,
    sources: product.sources,
    matchMethods: product.matchMethods,
    manualForm: product.manualForm,
    aliases: product.aliases,
    warnings: product.warnings.filter((warning) => !warning.startsWith('Entrada considerada') && !warning.startsWith('Divergência entre estoque') && !warning.startsWith('Cobertura recalculada')),
  }, params));
}

export function csvEscape(value) {
  const content = String(value ?? '');
  return `"${content.replaceAll('"', '""')}"`;
}

export function exportRows(products, adjustments, selectedIds, storeName) {
  return products
    .filter((product) => selectedIds.has(product.id))
    .map((product) => {
      const adjusted = adjustments.get(product.id);
      const quantity = adjusted === undefined ? product.suggested : adjusted;
      return {
        loja: storeName === 'ALL' ? product.store : storeName,
        codigo: product.code,
        produto: product.description,
        pasta: product.paste,
        prioridade: PRIORITY_META[product.priority].label,
        pedido: Math.max(0, Number(quantity) || 0),
        estoqueLoja: product.storeStock,
        entradas: product.inbound,
        saldoDeposito: product.depotAvailable,
        coberturaAtual: product.coverage === null ? '' : product.coverage.toFixed(1),
        coberturaProjetada: product.coverageAfterOrder === null ? '' : product.coverageAfterOrder.toFixed(1),
        motivo: product.reason,
        alertas: product.warnings.join(' | '),
      };
    })
    .filter((row) => row.pedido > 0);
}

export function makeCsv(rows) {
  const headers = ['Loja', 'Código', 'Produto', 'Pasta', 'Prioridade', 'Quantidade para pedido', 'Estoque loja', 'Entradas consideradas', 'Saldo depósito', 'Cobertura atual (dias)', 'Cobertura projetada (dias)', 'Motivo', 'Alertas'];
  const data = rows.map((row) => [row.loja, row.codigo, row.produto, row.pasta, row.prioridade, row.pedido, row.estoqueLoja, row.entradas, row.saldoDeposito, row.coberturaAtual, row.coberturaProjetada, row.motivo, row.alertas]);
  return `\uFEFF${[headers, ...data].map((row) => row.map(csvEscape).join(';')).join('\n')}`;
}
