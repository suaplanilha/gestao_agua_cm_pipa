/**
 * WATER ERP SAE - Backend Central
 * Google Apps Script (V8) | SAE (Sistema Apollo Enterprise)
 */

const APP_ID = 'WATER_ERP_SAE_001';
const SCHEMA_VERSION = '1.1.0';
const TZ = 'America/Sao_Paulo';
const SETUP_FLAG_KEY = 'DB_SETUP_DONE';
const SETUP_FLAG_VERSION_KEY = 'DB_SETUP_SCHEMA_VERSION';
const ENABLE_SHEET_LOGS_KEY = 'ENABLE_SHEET_LOGS';
const CACHE_TTL_SECONDS = 300;
const ENTITY_SCHEMAS = {
  entregas: [
    'uuid',
    'created_at',
    'updated_at',
    'data_iso',
    'local',
    'volume_m3',
    'km_inic',
    'km_fim',
    'km_delta',
    'h_inic',
    'h_fim',
    'h_delta',
    'status',
    'audit_user'
  ],
  abastecimentos: [
    'uuid',
    'created_at',
    'updated_at',
    'data_iso',
    'litros',
    'valor_total',
    'valor_litro',
    'km_atual',
    'h_atual',
    'posto',
    'observacao',
    'audit_user'
  ],
  despesas: [
    'uuid',
    'created_at',
    'updated_at',
    'data_iso',
    'categoria',
    'descricao',
    'valor',
    'centro_custo',
    'audit_user'
  ]
};

/**
 * WebApp entrypoint
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('WaterERP SAE')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Setup idempotente do banco SAE em Sheets.
 */
function setupDatabase() {
  const props = PropertiesService.getScriptProperties();
  const alreadySetup = props.getProperty(SETUP_FLAG_KEY) === '1';
  const currentVersion = props.getProperty(SETUP_FLAG_VERSION_KEY);

  if (alreadySetup && currentVersion === SCHEMA_VERSION) {
    return {
      ok: true,
      data: {
        appId: APP_ID,
        entities: Object.keys(ENTITY_SCHEMAS),
        setupSkipped: true
      }
    };
  }

  return withWriteLock_(() => {
    const ss = getSpreadsheet_();

    Object.keys(ENTITY_SCHEMAS).forEach((entity) => {
      const headers = ENTITY_SCHEMAS[entity];
      let sheet = ss.getSheetByName(entity);
      if (!sheet) {
        sheet = ss.insertSheet(entity);
      }
      ensureHeaders_(sheet, headers);
    });

    let logsSheet = ss.getSheetByName('SYS_LOGS');
    if (!logsSheet) {
      logsSheet = ss.insertSheet('SYS_LOGS');
    }
    ensureHeaders_(logsSheet, ['timestamp', 'app_id', 'acao', 'detalhes', 'usuario']);

    let metaSheet = ss.getSheetByName('SYS_META');
    if (!metaSheet) {
      metaSheet = ss.insertSheet('SYS_META');
    }
    ensureHeaders_(metaSheet, ['key', 'value', 'updated_at']);
    upsertMeta_(metaSheet, 'app_id', APP_ID);
    upsertMeta_(metaSheet, 'schema_version', SCHEMA_VERSION);
    upsertMeta_(metaSheet, 'last_setup_at', nowIso_());

    props.setProperty(SETUP_FLAG_KEY, '1');
    props.setProperty(SETUP_FLAG_VERSION_KEY, SCHEMA_VERSION);

    writeLog_('SETUP_DATABASE', 'Setup idempotente executado com sucesso.');

    return {
      ok: true,
      data: {
        appId: APP_ID,
        entities: Object.keys(ENTITY_SCHEMAS)
      }
    };
  });
}

function apiBootstrap(month) {
  return withApiResponse_('BOOTSTRAP', () => {
    setupDatabase();
    const refMonth = sanitizeMonth_(month || currentMonth_());

    return {
      appId: APP_ID,
      month: refMonth,
      dashboard: buildDashboard_(refMonth),
      entregas: listEntregas_({ month: refMonth, limit: 200 }),
      abastecimentos: listAbastecimentos_({ month: refMonth, limit: 200 }),
      despesas: listDespesas_({ month: refMonth, limit: 200 }),
      relatorio: buildMonthlyReport_(refMonth)
    };
  });
}

function apiCreateEntrega(payload) {
  return withApiResponse_('CREATE_ENTREGA', () => {
    setupDatabase();
    const input = payload || {};

    const dataIso = normalizeDate_(input.data || input.data_iso);
    const local = String(input.local || '').trim();
    const volume = normalizeNumber_(input.volume_m3 ?? input.volume);
    const kmInic = normalizeNumber_(input.km_inic ?? input.kmInicial);
    const kmFim = normalizeNumber_(input.km_fim ?? input.kmFinal);
    const hInic = normalizeNumber_(input.h_inic ?? input.hInicial);
    const hFim = normalizeNumber_(input.h_fim ?? input.hFinal);

    assert_(local, 'Local é obrigatório.');
    assert_(volume > 0, 'Volume deve ser maior que zero.');
    assert_(kmFim >= kmInic, 'KM final deve ser maior ou igual ao KM inicial.');
    assert_(hFim >= hInic, 'Horímetro final deve ser maior ou igual ao inicial.');

    const row = {
      uuid: Utilities.getUuid(),
      created_at: nowIso_(),
      updated_at: nowIso_(),
      data_iso: dataIso,
      local,
      volume_m3: round2_(volume),
      km_inic: round2_(kmInic),
      km_fim: round2_(kmFim),
      km_delta: round2_(kmFim - kmInic),
      h_inic: round2_(hInic),
      h_fim: round2_(hFim),
      h_delta: round2_(hFim - hInic),
      status: 'concluido',
      audit_user: getActiveUserEmail_()
    };

    appendEntityRow_('entregas', row);
    invalidateDashboardCache_(row.data_iso);
    writeLog_('CREATE_ENTREGA', `Entrega ${row.uuid} criada.`);

    return row;
  });
}

function apiCreateAbastecimento(payload) {
  return withApiResponse_('CREATE_ABASTECIMENTO', () => {
    setupDatabase();
    const input = payload || {};

    const litros = normalizeNumber_(input.litros);
    const valorTotal = normalizeNumber_(input.valor_total ?? input.valorTotal);
    const kmAtual = normalizeNumber_(input.km_atual ?? input.kmAtual);
    const hAtual = normalizeNumber_(input.h_atual ?? input.hAtual);

    assert_(litros > 0, 'Litros deve ser maior que zero.');
    assert_(valorTotal > 0, 'Valor total deve ser maior que zero.');

    const row = {
      uuid: Utilities.getUuid(),
      created_at: nowIso_(),
      updated_at: nowIso_(),
      data_iso: normalizeDate_(input.data || input.data_iso),
      litros: round2_(litros),
      valor_total: round2_(valorTotal),
      valor_litro: round2_(valorTotal / litros),
      km_atual: round2_(kmAtual),
      h_atual: round2_(hAtual),
      posto: String(input.posto || '').trim(),
      observacao: String(input.observacao || '').trim(),
      audit_user: getActiveUserEmail_()
    };

    appendEntityRow_('abastecimentos', row);
    invalidateDashboardCache_(row.data_iso);
    writeLog_('CREATE_ABASTECIMENTO', `Abastecimento ${row.uuid} criado.`);

    return row;
  });
}

function apiCreateDespesa(payload) {
  return withApiResponse_('CREATE_DESPESA', () => {
    setupDatabase();
    const input = payload || {};

    const valor = normalizeNumber_(input.valor);
    assert_(valor > 0, 'Valor da despesa deve ser maior que zero.');

    const row = {
      uuid: Utilities.getUuid(),
      created_at: nowIso_(),
      updated_at: nowIso_(),
      data_iso: normalizeDate_(input.data || input.data_iso),
      categoria: String(input.categoria || 'geral').trim() || 'geral',
      descricao: String(input.descricao || '').trim(),
      valor: round2_(valor),
      centro_custo: String(input.centro_custo || input.centroCusto || '').trim(),
      audit_user: getActiveUserEmail_()
    };

    appendEntityRow_('despesas', row);
    invalidateDashboardCache_(row.data_iso);
    writeLog_('CREATE_DESPESA', `Despesa ${row.uuid} criada.`);

    return row;
  });
}

function apiGetLastEntrega() {
  return withApiResponse_('GET_LAST_ENTREGA', () => {
    const rows = readEntityRows_('entregas')
      .sort((a, b) => String(b.created_at || '').localeCompare(String(a.created_at || '')));

    if (!rows.length) {
      return {
        kmFinal: 0,
        hFinal: 0
      };
    }

    const last = rows[0];
    return {
      kmFinal: Number(last.km_fim || 0),
      hFinal: Number(last.h_fim || 0),
      data: toIsoDate_(last.data_iso)
    };
  });
}

function apiUpdateEntrega(payload) {
  return withApiResponse_('UPDATE_ENTREGA', () => {
    const input = payload || {};
    const uuid = String(input.uuid || '').trim();
    assert_(uuid, 'UUID da entrega é obrigatório para edição.');

    const current = findEntityByUuid_('entregas', uuid);
    assert_(current && current.rowNumber, 'Entrega não encontrada.');

    const dataIso = normalizeDate_(input.data || input.data_iso || current.record.data_iso);
    const local = String((input.local ?? current.record.local) || '').trim();
    const volume = normalizeNumber_(input.volume_m3 ?? input.volume ?? current.record.volume_m3);
    const kmInic = normalizeNumber_(input.km_inic ?? input.kmInicial ?? current.record.km_inic);
    const kmFim = normalizeNumber_(input.km_fim ?? input.kmFinal ?? current.record.km_fim);
    const hInic = normalizeNumber_(input.h_inic ?? input.hInicial ?? current.record.h_inic);
    const hFim = normalizeNumber_(input.h_fim ?? input.hFinal ?? current.record.h_fim);

    assert_(local, 'Local é obrigatório.');
    assert_(volume > 0, 'Volume deve ser maior que zero.');
    assert_(kmFim >= kmInic, 'KM final deve ser maior ou igual ao KM inicial.');
    assert_(hFim >= hInic, 'Horímetro final deve ser maior ou igual ao inicial.');

    const updated = {
      ...current.record,
      updated_at: nowIso_(),
      data_iso: dataIso,
      local,
      volume_m3: round2_(volume),
      km_inic: round2_(kmInic),
      km_fim: round2_(kmFim),
      km_delta: round2_(kmFim - kmInic),
      h_inic: round2_(hInic),
      h_fim: round2_(hFim),
      h_delta: round2_(hFim - hInic),
      audit_user: getActiveUserEmail_()
    };

    updateEntityRow_('entregas', current.rowNumber, updated);
    invalidateDashboardCache_(updated.data_iso);
    writeLog_('UPDATE_ENTREGA', `Entrega ${uuid} atualizada.`);
    return { uuid };
  });
}

function apiDeleteEntrega(uuid) {
  return withApiResponse_('DELETE_ENTREGA', () => {
    const id = String(uuid || '').trim();
    assert_(id, 'UUID da entrega é obrigatório para exclusão.');
    const row = findEntityByUuid_('entregas', id);
    assert_(row && row.rowNumber, 'Entrega não encontrada.');
    invalidateDashboardCache_(row.record.data_iso);
    deleteEntityRow_('entregas', row.rowNumber);
    writeLog_('DELETE_ENTREGA', `Entrega ${id} removida.`);
    return { uuid: id };
  });
}

function apiUpdateDespesa(payload) {
  return withApiResponse_('UPDATE_DESPESA', () => {
    const input = payload || {};
    const uuid = String(input.uuid || '').trim();
    assert_(uuid, 'UUID da despesa é obrigatório para edição.');

    const current = findEntityByUuid_('despesas', uuid);
    assert_(current && current.rowNumber, 'Despesa não encontrada.');

    const valor = normalizeNumber_(input.valor ?? current.record.valor);
    assert_(valor > 0, 'Valor da despesa deve ser maior que zero.');

    const updated = {
      ...current.record,
      updated_at: nowIso_(),
      data_iso: normalizeDate_(input.data || input.data_iso || current.record.data_iso),
      categoria: String((input.categoria ?? current.record.categoria) || 'geral').trim() || 'geral',
      descricao: String((input.descricao ?? current.record.descricao) || '').trim(),
      valor: round2_(valor),
      centro_custo: String((input.centro_custo ?? input.centroCusto ?? current.record.centro_custo) || '').trim(),
      audit_user: getActiveUserEmail_()
    };

    updateEntityRow_('despesas', current.rowNumber, updated);
    invalidateDashboardCache_(updated.data_iso);
    writeLog_('UPDATE_DESPESA', `Despesa ${uuid} atualizada.`);
    return { uuid };
  });
}

function apiUpdateAbastecimento(payload) {
  return withApiResponse_('UPDATE_ABASTECIMENTO', () => {
    const input = payload || {};
    const uuid = String(input.uuid || '').trim();
    assert_(uuid, 'UUID do abastecimento é obrigatório para edição.');

    const current = findEntityByUuid_('abastecimentos', uuid);
    assert_(current && current.rowNumber, 'Abastecimento não encontrado.');

    const litros = normalizeNumber_(input.litros ?? current.record.litros);
    const valorTotal = normalizeNumber_(input.valor_total ?? input.valorTotal ?? current.record.valor_total);
    const kmAtual = normalizeNumber_(input.km_atual ?? input.kmAtual ?? current.record.km_atual);
    const hAtual = normalizeNumber_(input.h_atual ?? input.hAtual ?? current.record.h_atual);
    assert_(litros > 0, 'Litros deve ser maior que zero.');
    assert_(valorTotal > 0, 'Valor total deve ser maior que zero.');

    const updated = {
      ...current.record,
      updated_at: nowIso_(),
      data_iso: normalizeDate_(input.data || input.data_iso || current.record.data_iso),
      litros: round2_(litros),
      valor_total: round2_(valorTotal),
      valor_litro: round2_(valorTotal / litros),
      km_atual: round2_(kmAtual),
      h_atual: round2_(hAtual),
      posto: String((input.posto ?? current.record.posto) || '').trim(),
      observacao: String((input.observacao ?? current.record.observacao) || '').trim(),
      audit_user: getActiveUserEmail_()
    };

    updateEntityRow_('abastecimentos', current.rowNumber, updated);
    invalidateDashboardCache_(updated.data_iso);
    writeLog_('UPDATE_ABASTECIMENTO', `Abastecimento ${uuid} atualizado.`);
    return { uuid };
  });
}

function apiDeleteAbastecimento(uuid) {
  return withApiResponse_('DELETE_ABASTECIMENTO', () => {
    const id = String(uuid || '').trim();
    assert_(id, 'UUID do abastecimento é obrigatório para exclusão.');
    const row = findEntityByUuid_('abastecimentos', id);
    assert_(row && row.rowNumber, 'Abastecimento não encontrado.');
    invalidateDashboardCache_(row.record.data_iso);
    deleteEntityRow_('abastecimentos', row.rowNumber);
    writeLog_('DELETE_ABASTECIMENTO', `Abastecimento ${id} removido.`);
    return { uuid: id };
  });
}

function apiDeleteDespesa(uuid) {
  return withApiResponse_('DELETE_DESPESA', () => {
    const id = String(uuid || '').trim();
    assert_(id, 'UUID da despesa é obrigatório para exclusão.');
    const row = findEntityByUuid_('despesas', id);
    assert_(row && row.rowNumber, 'Despesa não encontrada.');
    invalidateDashboardCache_(row.record.data_iso);
    deleteEntityRow_('despesas', row.rowNumber);
    writeLog_('DELETE_DESPESA', `Despesa ${id} removida.`);
    return { uuid: id };
  });
}

function apiListEntregas(filters) {
  return withApiResponse_('LIST_ENTREGAS', () => listEntregas_(filters || {}));
}

function apiListDespesas(filters) {
  return withApiResponse_('LIST_DESPESAS', () => listDespesas_(filters || {}));
}

function apiListAbastecimentos(filters) {
  return withApiResponse_('LIST_ABASTECIMENTOS', () => listAbastecimentos_(filters || {}));
}

function apiGetDashboard(month) {
  return withApiResponse_('GET_DASHBOARD', () => {
    const refMonth = sanitizeMonth_(month || currentMonth_());
    return getDashboardCached_(refMonth);
  });
}

function apiGetRelatorioMensal(month) {
  return withApiResponse_('GET_RELATORIO_MENSAL', () => buildMonthlyReport_(sanitizeMonth_(month || currentMonth_())));
}

/** =========================
 * Private helpers
 * =======================*/

function withApiResponse_(action, handler) {
  try {
    const data = handler();
    return {
      ok: true,
      data,
      error: null,
      meta: {
        ts: nowIso_(),
        action
      }
    };
  } catch (error) {
    const message = error && error.message ? error.message : String(error);
    writeLog_(`ERROR_${action}`, message);
    return {
      ok: false,
      data: null,
      error: {
        code: 'VALIDATION_OR_RUNTIME_ERROR',
        message
      },
      meta: {
        ts: nowIso_(),
        action
      }
    };
  }
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const fixedId = props.getProperty('SPREADSHEET_ID');
  if (fixedId) return SpreadsheetApp.openById(fixedId);

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  throw new Error('Planilha não encontrada. Defina SPREADSHEET_ID em Script Properties.');
}

function ensureHeaders_(sheet, headers) {
  const hasData = sheet.getLastRow() > 0;
  if (!hasData) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    return;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const mismatch = headers.some((h, i) => currentHeaders[i] !== h);
  if (mismatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.setFrozenRows(1);
}

function upsertMeta_(metaSheet, key, value) {
  const values = metaSheet.getDataRange().getValues();
  const idx = values.findIndex((r, i) => i > 0 && r[0] === key);
  const now = nowIso_();

  if (idx === -1) {
    metaSheet.appendRow([key, value, now]);
    return;
  }

  const row = idx + 1;
  metaSheet.getRange(row, 1, 1, 3).setValues([[key, value, now]]);
}

function appendEntityRow_(entity, rowObject) {
  withWriteLock_(() => {
    const ss = getSpreadsheet_();
    const headers = ENTITY_SCHEMAS[entity];
    assert_(headers, `Entidade inválida: ${entity}`);

    const sheet = ss.getSheetByName(entity);
    assert_(sheet, `Aba da entidade não encontrada: ${entity}`);

    const row = headers.map((header) => rowObject[header] ?? '');
    sheet.appendRow(row);
  });
}

function updateEntityRow_(entity, rowNumber, rowObject) {
  withWriteLock_(() => {
    const ss = getSpreadsheet_();
    const headers = ENTITY_SCHEMAS[entity];
    assert_(headers, `Entidade inválida: ${entity}`);
    const sheet = ss.getSheetByName(entity);
    assert_(sheet, `Aba da entidade não encontrada: ${entity}`);
    const row = headers.map((header) => rowObject[header] ?? '');
    sheet.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
  });
}

function deleteEntityRow_(entity, rowNumber) {
  withWriteLock_(() => {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(entity);
    assert_(sheet, `Aba da entidade não encontrada: ${entity}`);
    sheet.deleteRow(rowNumber);
  });
}

function readEntityRows_(entity) {
  const ss = getSpreadsheet_();
  const headers = ENTITY_SCHEMAS[entity];
  assert_(headers, `Entidade inválida: ${entity}`);

  const sheet = ss.getSheetByName(entity);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  return values.map((row) => {
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = row[idx];
    });
    return obj;
  });
}

function findEntityByUuid_(entity, uuid) {
  const ss = getSpreadsheet_();
  const headers = ENTITY_SCHEMAS[entity];
  assert_(headers, `Entidade inválida: ${entity}`);
  const sheet = ss.getSheetByName(entity);
  assert_(sheet, `Aba da entidade não encontrada: ${entity}`);
  if (sheet.getLastRow() <= 1) return null;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  const uuidIdx = headers.indexOf('uuid');
  for (let i = 0; i < data.length; i += 1) {
    if (String(data[i][uuidIdx]) === uuid) {
      const record = {};
      headers.forEach((h, idx) => {
        record[h] = data[i][idx];
      });
      return {
        rowNumber: i + 2,
        record
      };
    }
  }
  return null;
}

function listEntregas_(filters) {
  const month = sanitizeMonth_(filters.month || currentMonth_());
  const limit = Number(filters.limit || 200);

  const rows = readEntityRows_('entregas')
    .filter((r) => monthFromValue_(r.data_iso) === month)
    .sort((a, b) => toIsoDate_(b.data_iso).localeCompare(toIsoDate_(a.data_iso)))
    .slice(0, limit)
    .map((r) => ({
      id: r.uuid,
      uuid: r.uuid,
      data: toIsoDate_(r.data_iso),
      local: r.local,
      volume: Number(r.volume_m3 || 0),
      kmInicial: Number(r.km_inic || 0),
      kmFinal: Number(r.km_fim || 0),
      kmDelta: Number(r.km_delta || 0),
      hInicial: Number(r.h_inic || 0),
      hFinal: Number(r.h_fim || 0),
      hDelta: Number(r.h_delta || 0),
      status: r.status || 'concluido'
    }));

  return rows;
}

function listDespesas_(filters) {
  const month = sanitizeMonth_(filters.month || currentMonth_());
  const limit = Number(filters.limit || 200);

  const rows = readEntityRows_('despesas')
    .filter((r) => monthFromValue_(r.data_iso) === month)
    .sort((a, b) => toIsoDate_(b.data_iso).localeCompare(toIsoDate_(a.data_iso)))
    .slice(0, limit)
    .map((r) => ({
      id: r.uuid,
      data: toIsoDate_(r.data_iso),
      categoria: r.categoria,
      descricao: r.descricao,
      valor: Number(r.valor || 0),
      centroCusto: r.centro_custo
    }));

  return rows;
}

function listAbastecimentos_(filters) {
  const month = sanitizeMonth_(filters.month || currentMonth_());
  const limit = Number(filters.limit || 200);

  return readEntityRows_('abastecimentos')
    .filter((r) => monthFromValue_(r.data_iso) === month)
    .sort((a, b) => toIsoDate_(b.data_iso).localeCompare(toIsoDate_(a.data_iso)))
    .slice(0, limit)
    .map((r) => ({
      id: r.uuid,
      data: toIsoDate_(r.data_iso),
      litros: Number(r.litros || 0),
      valorTotal: Number(r.valor_total || 0),
      valorLitro: Number(r.valor_litro || 0),
      kmAtual: Number(r.km_atual || 0),
      hAtual: Number(r.h_atual || 0),
      posto: r.posto || '',
      observacao: r.observacao || ''
    }));
}

function buildDashboard_(month) {
  const entregas = listEntregas_({ month, limit: 10000 });
  const despesas = listDespesas_({ month, limit: 10000 });

  const totalVolume = round2_(entregas.reduce((acc, e) => acc + Number(e.volume || 0), 0));
  const totalKm = round2_(entregas.reduce((acc, e) => acc + Number(e.kmDelta || 0), 0));
  const totalHoras = round2_(entregas.reduce((acc, e) => acc + Number(e.hDelta || 0), 0));
  const totalDespesas = round2_(despesas.reduce((acc, d) => acc + Number(d.valor || 0), 0));
  const totalViagens = entregas.length;
  const custoPorM3 = totalVolume > 0 ? round2_(totalDespesas / totalVolume) : 0;

  const custosPorCategoriaObj = despesas.reduce((acc, d) => {
    const key = (d.categoria || 'geral').toLowerCase();
    acc[key] = round2_((acc[key] || 0) + Number(d.valor || 0));
    return acc;
  }, {});

  const custosPorCategoria = Object.keys(custosPorCategoriaObj).map((k) => ({
    categoria: k,
    valor: custosPorCategoriaObj[k]
  }));

  const volumeByWeek = [0, 0, 0, 0, 0];
  entregas.forEach((e) => {
    const day = Number(String(e.data).slice(8, 10) || 1);
    const weekIdx = Math.min(4, Math.floor((day - 1) / 7));
    volumeByWeek[weekIdx] += Number(e.volume || 0);
  });

  return {
    month,
    totalVolume,
    totalKm,
    totalHoras,
    totalDespesas,
    totalViagens,
    custoPorM3,
    custosPorCategoria,
    volumeSeries: {
      labels: ['Semana 1', 'Semana 2', 'Semana 3', 'Semana 4', 'Semana 5'],
      data: volumeByWeek.map((n) => round2_(n))
    }
  };
}

function getDashboardCached_(month) {
  const cache = CacheService.getScriptCache();
  const key = `dashboard:${month}`;
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  const computed = buildDashboard_(month);
  cache.put(key, JSON.stringify(computed), CACHE_TTL_SECONDS);
  return computed;
}

function invalidateDashboardCache_(dateValue) {
  const cache = CacheService.getScriptCache();
  const month = monthFromValue_(dateValue);
  if (!month) return;
  cache.remove(`dashboard:${month}`);
}

function buildMonthlyReport_(month) {
  const entregas = listEntregas_({ month, limit: 10000 });
  const abastecimentos = listAbastecimentos_({ month, limit: 10000 });

  const kmInic = entregas.length ? Math.min.apply(null, entregas.map((e) => e.kmInicial)) : 0;
  const kmFim = entregas.length ? Math.max.apply(null, entregas.map((e) => e.kmFinal)) : 0;
  const hInic = entregas.length ? Math.min.apply(null, entregas.map((e) => e.hInicial)) : 0;
  const hFim = entregas.length ? Math.max.apply(null, entregas.map((e) => e.hFinal)) : 0;

  const totalCombustivel = abastecimentos.reduce((acc, a) => acc + Number(a.valorTotal || 0), 0);
  const totalLitros = abastecimentos.reduce((acc, a) => acc + Number(a.litros || 0), 0);
  const mediaValorLitro = totalLitros > 0 ? round2_(totalCombustivel / totalLitros) : 0;

  return {
    month,
    generatedAt: nowIso_(),
    resumo: {
      viagens: entregas.length,
      volumeTotal: round2_(entregas.reduce((acc, e) => acc + e.volume, 0)),
      kmInicial: round2_(kmInic),
      kmFinal: round2_(kmFim),
      kmTotal: round2_(kmFim - kmInic),
      hInicial: round2_(hInic),
      hFinal: round2_(hFim),
      hTotal: round2_(hFim - hInic),
      totalCombustivel: round2_(totalCombustivel),
      totalLitros: round2_(totalLitros),
      mediaValorLitro
    },
    combustivel: abastecimentos.map((a) => ({
      data: a.data,
      posto: a.posto || 'Abastecimento',
      kmAtual: round2_(a.kmAtual),
      hAtual: round2_(a.hAtual),
      litros: round2_(a.litros),
      valorLitro: round2_(a.valorLitro),
      valorTotal: round2_(a.valorTotal)
    }))
  };
}

function writeLog_(action, details) {
  console.log(`[${APP_ID}] ${action} | ${details}`);
  const props = PropertiesService.getScriptProperties();
  const enableSheetLogs = props.getProperty(ENABLE_SHEET_LOGS_KEY);
  if (enableSheetLogs === '0') {
    return;
  }

  const ss = getSpreadsheet_();
  let logSheet = ss.getSheetByName('SYS_LOGS');
  if (!logSheet) {
    logSheet = ss.insertSheet('SYS_LOGS');
    ensureHeaders_(logSheet, ['timestamp', 'app_id', 'acao', 'detalhes', 'usuario']);
  }

  logSheet.appendRow([nowIso_(), APP_ID, action, details, getActiveUserEmail_()]);
}

function withWriteLock_(handler) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return handler();
  } finally {
    lock.releaseLock();
  }
}

function getActiveUserEmail_() {
  const email = Session.getActiveUser().getEmail();
  return email || 'usuario_nao_identificado';
}

function normalizeNumber_(value) {
  if (typeof value === 'number') return value;
  const parsed = Number(String(value || '0').replace(',', '.'));
  if (Number.isNaN(parsed)) return 0;
  return parsed;
}

function normalizeDate_(value) {
  if (!value) return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const str = String(value).trim();
  const match = /^\d{4}-\d{2}-\d{2}$/.test(str);
  assert_(match, 'Data deve estar em formato ISO yyyy-MM-dd.');
  return str;
}

function toIsoDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }

  const str = String(value || '').trim();
  if (!str) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;

  const asDate = new Date(str);
  if (!Number.isNaN(asDate.getTime())) {
    return Utilities.formatDate(asDate, TZ, 'yyyy-MM-dd');
  }

  return str;
}

function monthFromValue_(value) {
  const iso = toIsoDate_(value);
  return iso ? iso.slice(0, 7) : '';
}

function sanitizeMonth_(value) {
  const month = String(value || '').trim();
  assert_(/^\d{4}-\d{2}$/.test(month), 'Mês deve estar em formato yyyy-MM.');
  return month;
}

function currentMonth_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
}

function nowIso_() {
  return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function round2_(value) {
  return Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100;
}

function assert_(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}
