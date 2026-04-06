/**
 * WATER ERP SAE - Backend Central
 * Google Apps Script (V8) | SAE (Sistema Apollo Enterprise)
 */

const APP_ID = 'WATER_ERP_SAE_001';
const TZ = 'America/Sao_Paulo';
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
  upsertMeta_(metaSheet, 'schema_version', '1.0.0');
  upsertMeta_(metaSheet, 'last_setup_at', nowIso_());

  writeLog_('SETUP_DATABASE', 'Setup idempotente executado com sucesso.');

  return {
    ok: true,
    data: {
      appId: APP_ID,
      entities: Object.keys(ENTITY_SCHEMAS)
    }
  };
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
      posto: String(input.posto || '').trim(),
      observacao: String(input.observacao || '').trim(),
      audit_user: getActiveUserEmail_()
    };

    appendEntityRow_('abastecimentos', row);
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
    writeLog_('CREATE_DESPESA', `Despesa ${row.uuid} criada.`);

    return row;
  });
}

function apiListEntregas(filters) {
  return withApiResponse_('LIST_ENTREGAS', () => listEntregas_(filters || {}));
}

function apiListDespesas(filters) {
  return withApiResponse_('LIST_DESPESAS', () => listDespesas_(filters || {}));
}

function apiGetDashboard(month) {
  return withApiResponse_('GET_DASHBOARD', () => buildDashboard_(sanitizeMonth_(month || currentMonth_())));
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
  const ss = getSpreadsheet_();
  const headers = ENTITY_SCHEMAS[entity];
  assert_(headers, `Entidade inválida: ${entity}`);

  const sheet = ss.getSheetByName(entity);
  assert_(sheet, `Aba da entidade não encontrada: ${entity}`);

  const row = headers.map((header) => rowObject[header] ?? '');
  sheet.appendRow(row);
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

function listEntregas_(filters) {
  const month = sanitizeMonth_(filters.month || currentMonth_());
  const limit = Number(filters.limit || 200);

  const rows = readEntityRows_('entregas')
    .filter((r) => String(r.data_iso || '').slice(0, 7) === month)
    .sort((a, b) => String(b.data_iso).localeCompare(String(a.data_iso)))
    .slice(0, limit)
    .map((r) => ({
      id: r.uuid,
      uuid: r.uuid,
      data: r.data_iso,
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
    .filter((r) => String(r.data_iso || '').slice(0, 7) === month)
    .sort((a, b) => String(b.data_iso).localeCompare(String(a.data_iso)))
    .slice(0, limit)
    .map((r) => ({
      id: r.uuid,
      data: r.data_iso,
      categoria: r.categoria,
      descricao: r.descricao,
      valor: Number(r.valor || 0),
      centroCusto: r.centro_custo
    }));

  return rows;
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

function buildMonthlyReport_(month) {
  const entregas = listEntregas_({ month, limit: 10000 });
  const despesas = listDespesas_({ month, limit: 10000 });
  const abastecimentos = readEntityRows_('abastecimentos')
    .filter((r) => String(r.data_iso || '').slice(0, 7) === month);

  const kmInic = entregas.length ? Math.min.apply(null, entregas.map((e) => e.kmInicial)) : 0;
  const kmFim = entregas.length ? Math.max.apply(null, entregas.map((e) => e.kmFinal)) : 0;
  const hInic = entregas.length ? Math.min.apply(null, entregas.map((e) => e.hInicial)) : 0;
  const hFim = entregas.length ? Math.max.apply(null, entregas.map((e) => e.hFinal)) : 0;

  const totalDespesas = despesas.reduce((acc, d) => acc + Number(d.valor || 0), 0);
  const totalAbastecimentos = abastecimentos.reduce((acc, a) => acc + normalizeNumber_(a.valor_total || 0), 0);

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
      totalDespesas: round2_(totalDespesas + totalAbastecimentos)
    },
    custos: [
      ...abastecimentos.map((a) => ({
        categoria: 'abastecimento',
        descricao: a.posto || 'Abastecimento',
        quantidade: `${round2_(normalizeNumber_(a.litros || 0))} L`,
        valor: round2_(normalizeNumber_(a.valor_total || 0))
      })),
      ...despesas.map((d) => ({
        categoria: d.categoria || 'despesa',
        descricao: d.descricao || '-',
        quantidade: d.centroCusto || '-',
        valor: round2_(d.valor || 0)
      }))
    ]
  };
}

function writeLog_(action, details) {
  const ss = getSpreadsheet_();
  let logSheet = ss.getSheetByName('SYS_LOGS');
  if (!logSheet) {
    logSheet = ss.insertSheet('SYS_LOGS');
    ensureHeaders_(logSheet, ['timestamp', 'app_id', 'acao', 'detalhes', 'usuario']);
  }

  logSheet.appendRow([nowIso_(), APP_ID, action, details, getActiveUserEmail_()]);
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
