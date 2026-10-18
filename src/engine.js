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
  if (!worksheet) return [];
  return XLSX.utils.sheet_to_json(worksheet, { defval: null, raw: true, dateNF: 'dd/mm/yyyy' });
}

function decodeLegacyText(data) {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  try {
    return new TextDecoder('windows-1252').decode(bytes);
  } catch {
    return new TextDecoder().decode(bytes);
  }
}

function looksLikeDelimitedText(data) {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const sample = bytes.slice(0, 4096);
  if (sample.includes(0)) return false;
  const text = decodeLegacyText(sample);
  const firstLine = text.split(/\r?\n/, 1)[0];
  return text.includes('\n') && (firstLine.includes('\t') || firstLine.includes(';'));
}

function inferStandaloneSheetName(workbook, fileName) {
  const fileKey = normalizeHeader(fileName);
  if (fileKey.includes('zerado')) return 'Zerado';
  if (fileKey.includes('saldo')) return 'SALDO';
  if (fileKey.includes('consolidado3meses') || fileKey.includes('consolidado')) return 'consolidado3meses';

  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const header = XLSX.utils.sheet_to_json(firstSheet, { header: 1, range: 0, raw: true, defval: '' })[0] ?? [];
  const headerKey = normalizeHeader(header.join(' '));
  if (headerKey.includes('zerado')) return 'Zerado';
  if (headerKey.includes('estcentral') || headerKey.includes('pedidodeloja')) return 'SALDO';
  if (headerKey.includes('valorvenda') || headerKey.includes('qtdevenda') || headerKey.includes('cobertura')) return 'consolidado3meses';
  return null;
}

function renameSingleSheet(workbook, nextName) {
  const currentName = workbook.SheetNames[0];
  if (!currentName || currentName === nextName) return;
  workbook.Sheets[nextName] = workbook.Sheets[currentName];
  delete workbook.Sheets[currentName];
  workbook.SheetNames = [nextName];
}

function readErpWorkbook(data, fileName) {
  const workbook = looksLikeDelimitedText(data)
    ? XLSX.read(decodeLegacyText(data), { type: 'string', FS: decodeLegacyText(data).split(/\r?\n/, 1)[0].includes('\t') ? '\t' : ';', raw: true, cellDates: true })
    : XLSX.read(data, { type: 'array', cellDates: true });
  if (workbook.SheetNames.length === 1) {
    const inferred = inferStandaloneSheetName(workbook, fileName);
    if (inferred) renameSingleSheet(workbook, inferred);
  }
  return workbook;
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
  const description = new Map();
  rows.forEach((row, index) => {
    if (row.descriptionKey) {
      if (!description.has(row.descriptionKey)) description.set(row.descriptionKey, []);
      description.get(row.descriptionKey).push(index);
    }
  });
  return { description };
}

function findMatch(product, incoming, sourceIndex) {
  const index = sourceIndex[incoming.source];
  if (!index) return { record: null, method: 'new' };
  const byDescription = incoming.descriptionKey
    ? uniqueCandidate(index.description.get(incoming.descriptionKey) ?? [])
    : null;
  if (byDescription !== null) return { record: product[byDescription], method: 'descrição' };

  return { record: null, method: 'new' };
}

function recordKey(row, order) {
  return row.descriptionKey || `${row.source}-${order}`;
}

function combineSources(saldoRows, consolidatedRows, zeradoRows) {
  const products = [];
  const sourceIndex = {
    SALDO: { description: new Map() },
    consolidado3meses: { description: new Map() },
    Zerado: { description: new Map() },
  };
  const allRows = [
    ...saldoRows.map((row) => ({ ...row, source: 'SALDO' })),
    ...consolidatedRows.map((row) => ({ ...row, source: 'consolidado3meses' })),
    ...zeradoRows.map((row) => ({ ...row, source: 'Zerado' })),
  ];

  const rebuiltIndex = () => {
    for (const source of Object.keys(sourceIndex)) {
      sourceIndex[source] = { description: new Map() };
    }
    products.forEach((product, index) => {
      Object.keys(product.sources).forEach((source) => {
        const row = product.sources[source];
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

function weightedDemand({ sales15, sales30, params }) {
  const shortPeriod = Math.max(1, Number(params?.shortPeriod) || 15);
  const saldoPeriod = Math.max(shortPeriod, 30, Number(params?.longPeriod) || 30);
  const components = [
    { value: sales15 === null ? null : sales15 / shortPeriod, weight: 0.6, label: `${shortPeriod}D SALDO` },
    { value: sales30 === null ? null : sales30 / saldoPeriod, weight: 0.4, label: `${saldoPeriod}D SALDO` },
  ].filter((component) => component.value !== null && Number.isFinite(component.value));

  if (!components.length) return { demand: 0, source: 'Sem histórico de venda' };
  const totalWeight = components.reduce((sum, component) => sum + component.weight, 0);
  const blended = components.reduce((sum, component) => sum + component.value * component.weight, 0) / totalWeight;
  const recent = sales15 !== null && Number.isFinite(sales15) ? sales15 / shortPeriod : components[0]?.value ?? 0;
  const demand = recent > blended * 1.35 ? Math.min(recent, blended * 1.5) : blended;
  const source = components.map((component) => component.label).join(' + ');
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
    params,
  });

  const storeStock = numberOrZero(saldo?.storeStock ?? zerado?.storeStock ?? consolidated?.storeStock);
  const saldoOrder = numberOrZero(saldo?.orderedQty);
  const consolidatedPending = numberOrZero(consolidated?.pending);
  const transit = numberOrZero(consolidated?.transit);
  const pending = Math.max(saldoOrder, consolidatedPending) + transit;
  const inboundDetails = {
    pendente: consolidatedPending,
    transito: transit,
    pedidoLoja: saldoOrder,
  };
  const depotStock = consolidated?.depositStock ?? saldo?.centralStock ?? zerado?.centralStock ?? null;
  const depotBalance = saldo?.balance ?? zerado?.balance ?? null;
  const depotAvailable = depotBalance ?? minConservative([saldo?.centralStock ?? null, zerado?.centralStock ?? null, consolidated?.depositStock ?? null]);
  const criticalDays = Math.max(1, Number(params?.criticalDays) || 2);
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
  if (depotBalance !== null && depotStock !== null && depotBalance > depotStock) warningSet.add('Saldo disponível maior que o estoque cheio do depósito: conferir a extração do ERP.');
  if (consolidated && saldo && coverage !== null && consolidated.coverage !== null && Math.abs(coverage - consolidated.coverage) > Math.max(4, consolidated.coverage * 0.5)) warningSet.add('Cobertura recalculada diverge da informada no consolidado.');
  if (demandInfo.demand === 0 && Boolean(zerado)) warningSet.add('Ruptura sem venda conhecida: item mantido para revisão, sem pedido automático.');
  if (pending > 0) warningSet.add(`Entrada considerada de ${pending.toFixed(0)} un. (pedido/pendência mais trânsito).`);
  if (pack > 1 && availableUnits > 0 && roundedPack === 0) warningSet.add(`Sugestão menor que o múltiplo de ${pack} un.; confirmar se o depósito pode separar unidades.`);

  let priority = 'NO_ACTION';
  let reason = 'Cobertura suficiente ou sem demanda para pedido automático.';
  if (hasDemand && depotAvailable <= 0 && (isRupture || (coverage !== null && coverage < params.targetDays))) {
    priority = 'BLOCKED';
    reason = 'Há demanda, mas não existe saldo liberado no depósito para atender a loja.';
  } else if (hasDemand && isRupture && depotAvailable > 0) {
    priority = 'P0';
    reason = 'Ruptura com venda: atender primeiro enquanto há saldo no depósito.';
  } else if (hasDemand && coverage !== null && coverage <= criticalDays && suggested > 0) {
    priority = 'P1';
    reason = `Cobertura projetada de até ${criticalDays} dias: risco imediato de falta.`;
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

function validateWorkbook(workbook, fileName) {
  const names = workbook.SheetNames.map((name) => name.trim().toLocaleLowerCase('pt-BR'));
  const recognized = REQUIRED_SHEETS.filter((required) => names.includes(required.toLocaleLowerCase('pt-BR')));
  if (!recognized.length) {
    throw new Error(`Não identificamos a fonte do arquivo “${fileName}”. Use um arquivo com nome SALDO, consolidado3meses ou Zerado, ou uma planilha com as três abas.`);
  }
  return {
    recognized,
    missing: REQUIRED_SHEETS.filter((required) => !recognized.includes(required)),
  };
}

function actualSheetName(workbook, expected) {
  return workbook.SheetNames.find((name) => name.trim().toLocaleLowerCase('pt-BR') === expected.toLocaleLowerCase('pt-BR'));
}

function productKeys(product) {
  const descriptions = [product.description, ...Object.values(product.sources ?? {}).map((source) => source.description)];
  return {
    descriptions: new Set(descriptions.map((description) => normalizeText(description)).filter(Boolean)),
  };
}

function findProductByMapping(products, _code, description) {
  const descriptionKey = normalizeText(description);
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
  })).filter((mapping) => rawText(mapping.saleDescription) && rawText(mapping.stockDescription) && normalizeText(mapping.saleDescription) !== normalizeText(mapping.stockDescription));
}

function rowsWithHeader(workbook, sheetName, requiredHeaders) {
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) return [];
  const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: true });
  const headerIndex = matrix.findIndex((row) => requiredHeaders.every((expected) => row.some((value) => normalizeHeader(value) === expected)));
  return headerIndex < 0 ? [] : XLSX.utils.sheet_to_json(worksheet, { range: headerIndex, defval: null, raw: true });
}

function isBaseWorkbook(workbook, fileName) {
  const fileKey = normalizeHeader(fileName);
  return fileKey.includes('dashboardpedidos') || Boolean(actualSheetName(workbook, '06_BASE_ANALITICA'));
}

function baseMetadataRows(workbook) {
  const metadata = [];
  const analyticalName = actualSheetName(workbook, '06_BASE_ANALITICA');
  if (analyticalName) {
    rowsWithHeader(workbook, analyticalName, ['chave', 'produto', 'setor', 'giro']).forEach((row) => {
      const code = getField(row, ['chave']);
      const description = getField(row, ['produto']);
      metadata.push({
        source: '06_BASE_ANALITICA',
        code: rawText(code),
        codeExact: codeExact(code),
        codeNumeric: codeNumeric(code),
        description: rawText(description),
        descriptionKey: normalizeText(description),
        sector: rawText(getField(row, ['setor'])),
        giro: rawText(getField(row, ['giro', 'classificacaodogiro'])),
        category: rawText(getField(row, ['pasta', 'categoria'])),
      });
    });
  }

  const baseName = actualSheetName(workbook, 'BASE');
  if (baseName) {
    rowsWithHeader(workbook, baseName, ['codigo', 'descricaocompleta']).forEach((row) => {
      const code = getField(row, ['codigo']);
      const description = getField(row, ['descricaocompleta']);
      metadata.push({
        source: 'BASE',
        code: rawText(code),
        codeExact: codeExact(code),
        codeNumeric: codeNumeric(code),
        description: rawText(description),
        descriptionKey: normalizeText(description),
        sector: rawText(getField(row, ['consolidado', 'setor'])),
        giro: '',
        category: rawText(getField(row, ['pasta'])),
      });
    });
  }
  return metadata.filter((row) => row.codeExact || row.codeNumeric || row.descriptionKey);
}

function applyBaseMetadata(products, metadataRows) {
  if (!metadataRows.length) return 0;
  const index = buildIndex(metadataRows);
  let matched = 0;
  products.forEach((product) => {
    const keys = productKeys(product);
    const candidates = [...keys.descriptions].flatMap((key) => index.description.get(key) ?? []);
    const unique = [...new Set(candidates)];
    const metadata = unique.length ? metadataRows[unique[0]] : null;
    if (!metadata) return;
    product.sector = metadata.sector || product.sector || '';
    product.giro = metadata.giro || product.giro || '';
    if (product.paste === 'Sem pasta' && metadata.category) product.paste = metadata.category;
    product.baseMatch = metadata.source;
    matched += 1;
  });
  return matched;
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
  const files = Array.isArray(file) ? file.filter(Boolean) : [file].filter(Boolean);
  if (!files.length) throw new Error('Selecione pelo menos um arquivo do ERP.');

  const sourceRows = { SALDO: [], consolidado3meses: [], Zerado: [] };
  const importedMappings = [];
  const recognizedSheets = new Set();
  const baseMetadata = [];
  const baseFiles = [];
  for (const currentFile of files) {
    const data = await currentFile.arrayBuffer();
    const workbook = readErpWorkbook(data, currentFile.name);
    if (isBaseWorkbook(workbook, currentFile.name)) {
      baseFiles.push(currentFile.name);
      baseMetadata.push(...baseMetadataRows(workbook));
      importedMappings.push(...panelMappings(workbook));
      continue;
    }
    const validation = validateWorkbook(workbook, currentFile.name);
    validation.recognized.forEach((source) => {
      recognizedSheets.add(source);
      const rows = sheetRows(workbook, actualSheetName(workbook, source)).map((row) => normalizedRow(row, source));
      sourceRows[source].push(...rows);
    });
    importedMappings.push(...panelMappings(workbook));
  }

  if (!Object.values(sourceRows).some((rows) => rows.length)) {
    if (baseFiles.length) throw new Error('A base histórica foi reconhecida, mas nenhuma atualização do ERP foi selecionada. Adicione consolidado3meses, SALDO e Zerado para montar o pedido atual.');
    throw new Error('Os arquivos foram reconhecidos, mas não há registros de produtos para processar.');
  }

  const saldoRows = sourceRows.SALDO;
  const consolidatedRows = sourceRows.consolidado3meses;
  const zeradoRows = sourceRows.Zerado;
  const rawProducts = combineSources(saldoRows, consolidatedRows, zeradoRows);
  const appliedMappings = applyProductMappings(rawProducts, [...importedMappings, ...(mappings ?? [])]);
  const baseMatched = applyBaseMetadata(rawProducts, baseMetadata);
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
    recognizedSheets: [...recognizedSheets],
    missingSheets: REQUIRED_SHEETS.filter((source) => !recognizedSheets.has(source)),
    unmatched: products.filter((product) => !product.hasSaldo || !product.hasConsolidated).length,
    matched: products.filter((product) => product.hasSaldo && product.hasConsolidated).length,
    descriptionMatches: products.filter((product) => product.matchMethods.includes('descrição')).length,
    mappedProducts: appliedMappings.length,
    baseRows: baseMetadata.length,
    baseMatched,
    baseFiles,
  };

  return { products, stores: stores.length ? stores : [DEFAULT_STORE], categories, stats, mappings: appliedMappings, fileName: files.map((currentFile) => currentFile.name).join(' + '), fileNames: files.map((currentFile) => currentFile.name), baseFileNames: baseFiles, importedAt: new Date() };
}

export function recalculateProducts(products, params) {
  return products.map((product) => calculateProduct({
    id: product.id,
    code: product.code,
    description: product.description,
    store: product.store,
    paste: product.paste,
    sector: product.sector,
    giro: product.giro,
    baseMatch: product.baseMatch,
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
