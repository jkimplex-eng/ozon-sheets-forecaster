const OZON_MENU = 'Ozon';
const OZON_SELLER_BASE_URL = 'https://api-seller.ozon.ru';
const OZON_PERFORMANCE_BASE_URL = 'https://api-performance.ozon.ru';
const SHEETS_SCHEMA = {
  Config: ['key', 'value', 'updated_at', 'notes'],
  Raw_Funnel_Daily: ['date', 'sku', 'offer_id', 'product_id', 'category', 'impressions', 'product_page_views', 'add_to_cart', 'orders', 'ordered_units', 'gross_revenue', 'cancellations', 'returns', 'buyout_units', 'ctr', 'cart_conversion', 'order_conversion', 'buyout_rate', 'source_endpoint', 'loaded_at'],
  Raw_Ads_Daily: ['date', 'campaign_id', 'campaign_name', 'campaign_type', 'placement', 'sku', 'offer_id', 'impressions', 'clicks', 'ctr', 'spend', 'avg_cpc', 'orders', 'revenue', 'drr_acos', 'roas', 'search_phrase', 'source_endpoint', 'loaded_at'],
  Ad_Campaigns: ['loaded_at', 'endpoint', 'query_params', 'campaign_id', 'campaign_name', 'campaign_type', 'adv_object_type', 'state', 'status', 'payment_type', 'daily_budget', 'budget', 'start_date', 'end_date', 'raw_json'],
  Unit_Costs: ['offer_id', 'unit_cost', 'currency', 'updated_at', 'notes'],
  Products: ['sku', 'offer_id', 'product_id', 'name', 'category', 'brand', 'status', 'price', 'discounted_price', 'rating', 'reviews_count', 'questions_count', 'content_score', 'moderation_status', 'loaded_at'],
  Stocks: ['date', 'sku', 'offer_id', 'warehouse_id', 'warehouse_name', 'available_stock', 'reserved_stock', 'days_of_stock', 'turnover_rate', 'stockout_risk', 'loaded_at'],
  Finance: ['date', 'sku', 'offer_id', 'orders', 'gross_revenue', 'discounts', 'commission', 'logistics', 'returns_cost', 'advertising_spend', 'net_revenue', 'estimated_margin', 'loaded_at'],
  Algorithm_Signals: ['date', 'sku', 'signal_group', 'signal_name', 'observed_change', 'likely_cause', 'confidence', 'supporting_metrics', 'recommended_test', 'priority'],
  Forecasts: ['generated_at', 'base_date', 'horizon_days', 'sku', 'metric', 'low', 'base', 'high', 'assumptions', 'reliability'],
  Recommendations: ['generated_at', 'priority', 'area', 'sku_or_campaign', 'issue', 'action', 'expected_impact', 'effort', 'owner', 'status'],
  Summary: ['section', 'metric', 'value', 'notes'],
  Diagnostics: ['run_at', 'severity', 'component', 'message', 'affected_endpoint', 'affected_tab', 'resolution'],
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(OZON_MENU)
    .addItem('1. Создать вкладки', 'createSheets')
    .addItem('2. Настроить ключи Ozon', 'setupOzonKeys')
    .addItem('3. Проверить подключение Ozon', 'testOzonConnection')
    .addItem('Проверить Performance API', 'testPerformanceConnection')
    .addItem('Диагностика рекламных кампаний', 'diagnosePerformanceCampaigns')
    .addItem('4. Обновить товары', 'updateProducts')
    .addItem('5. Обновить остатки', 'updateStocks')
    .addItem('6. Обновить данные за вчера', 'runYesterday')
    .addItem('7. Обновить последние 30 дней', 'runLast30Days')
    .addItem('Обновить рекламу', 'updateAdsDailyMenu')
    .addItem('8. Построить прогноз', 'buildForecasts')
    .addItem('9. Создать ежедневный автозапуск', 'createDailyTrigger')
    .addItem('Обновить Summary', 'updateSummary')
    .addItem('Показать статус настроек', 'showSettingsStatus')
    .addToUi();
}

function createSheets() {
  Object.keys(SHEETS_SCHEMA).forEach(function (name) {
    ensureSheetWithHeader_(name, SHEETS_SCHEMA[name]);
  });
  seedUnitCostsIfEmpty_();
  writeDiagnostic('info', 'Sheets', 'Вкладки и заголовки проверены/созданы.', '', '', 'Готово');
  SpreadsheetApp.getUi().alert('Вкладки созданы');
}

function setupOzonKeys() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const keys = ['OZON_SELLER_CLIENT_ID', 'OZON_SELLER_API_KEY', 'OZON_PERFORMANCE_CLIENT_ID', 'OZON_PERFORMANCE_CLIENT_SECRET'];

  keys.forEach(function (key) {
    const response = ui.prompt('Настройка Ozon', 'Введите значение для ' + key + ':', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK) {
      const value = (response.getResponseText() || '').trim();
      if (value) props.setProperty(key, value);
    }
  });

  writeDiagnostic('info', 'Security', 'Ключи Ozon сохранены в Script Properties.', '', 'Config', 'Проверьте статус настроек');
  SpreadsheetApp.getUi().alert('Ключи сохранены');
}

function showSettingsStatus() {
  const props = PropertiesService.getScriptProperties();
  const keys = ['OZON_SELLER_CLIENT_ID', 'OZON_SELLER_API_KEY', 'OZON_PERFORMANCE_CLIENT_ID', 'OZON_PERFORMANCE_CLIENT_SECRET'];
  const lines = keys.map(function (key) { return key + ': ' + (props.getProperty(key) ? 'есть' : 'нет'); });
  writeDiagnostic('info', 'Security', 'Проверен статус ключей Ozon.', '', 'Config', 'Заполните отсутствующие ключи');
  SpreadsheetApp.getUi().alert('Статус настроек:\n\n' + lines.join('\n'));
}

function testOzonConnection() {
  createSheets();
  const ui = SpreadsheetApp.getUi();
  try {
    const probe = getProductListPage_('', 1, true);
    writeDiagnostic('info', 'Seller API', 'Подключение к Ozon Seller API успешно. Рабочий endpoint: ' + probe.path, probe.path, 'Diagnostics', 'Можно запускать обновление данных');
    ui.alert('Успешно: подключение к Ozon работает. Endpoint: ' + probe.path);
  } catch (err) {
    const message = sanitizeError_(err);
    writeDiagnostic('error', 'Seller API', message, 'product/list fallback', 'Diagnostics', 'Проверьте Client-Id, Api-Key, права API и актуальные Ozon endpoints');
    ui.alert('Ошибка подключения к Ozon.\n\nПроверьте вкладку Diagnostics.');
  }
}

function updateProducts() {
  createSheets();
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName('Products');
  const loadedAt = nowIso_();
  const rows = [];
  let lastId = '';
  let selectedPath = '';

  try {
    while (true) {
      const response = getProductListPage_(lastId, 100, false, selectedPath);
      selectedPath = response.path;
      const data = response.response || {};
      const items = extractItems_(data);
      const normalized = items.map(normalizeProductListItem_);
      enrichProductsByInfoList_(normalized, selectedPath);

      normalized.forEach(function (product) {
        rows.push([
          product.sku,
          product.offer_id,
          product.product_id,
          product.name,
          product.category,
          product.brand,
          product.status,
          '', '', '', '', '', '', '',
          loadedAt,
        ]);
      });

      const nextLastId = extractLastId_(data);
      if (!nextLastId || nextLastId === lastId) break;
      lastId = nextLastId;
    }

    clearDataRows_(sheet);
    if (rows.length > 0) sheet.getRange(2, 1, rows.length, SHEETS_SCHEMA.Products.length).setValues(rows);
    writeDiagnostic('info', 'Products', 'Товары обновлены: ' + rows.length + ' строк. Рабочий endpoint: ' + (selectedPath || 'не определен'), selectedPath || '', 'Products', 'Готово');
    ui.alert('Обновление завершено');
  } catch (err) {
    writeDiagnostic('error', 'Products', sanitizeError_(err), 'product/list fallback', 'Products', 'Проверьте endpoint, права API и ключи');
    ui.alert('Ошибка: смотри вкладку Diagnostics');
  }
}

function getProductListPage_(lastId, limit, logFallbackWarnings, preferredPath) {
  const payload = { filter: { visibility: 'ALL' }, limit: limit || 100, last_id: lastId || '' };
  const candidates = preferredPath ? [{ path: preferredPath, payload: payload }] : [
    { path: '/v3/product/list', payload: payload },
    { path: '/v2/product/list', payload: payload },
  ];
  return ozonSellerPostWithFallback_(candidates, {
    component: 'Seller API',
    targetTab: 'Diagnostics',
    logFallbackWarnings: !!logFallbackWarnings,
  });
}

function enrichProductsByInfoList_(products, listPath) {
  const needInfo = products.filter(function (p) {
    return !!p.product_id && (!p.name || !p.status);
  });
  if (!needInfo.length) return;

  const productIds = needInfo.map(function (p) { return p.product_id; });
  const map = fetchProductInfoMap_(productIds);
  if (!map) return;

  products.forEach(function (product) {
    const info = map[String(product.product_id)];
    if (!info) return;
    if (!product.name) product.name = safe_(info.name || (info.sources && info.sources[0] && info.sources[0].name));
    if (!product.status) {
      product.status = safe_(
        info.visibility ||
        (info.status && (info.status.state || info.status.status)) ||
        info.state
      );
    }
    if (!product.offer_id) product.offer_id = safe_(info.offer_id);
  });
  writeDiagnostic('info', 'Products', 'Дополнены name/status через /v3/product/info/list: ' + needInfo.length + ' товаров.', listPath, 'Products', 'Готово');
}

function fetchProductInfoMap_(productIds) {
  try {
    const chunks = chunkArray_(productIds, 1000);
    const resultMap = {};
    chunks.forEach(function (chunk) {
      const payload = { product_id: chunk };
      const raw = ozonSellerPostRaw_('/v3/product/info/list', payload);
      if (raw.code >= 200 && raw.code < 300) {
        const body = raw.json || {};
        const items = extractInfoItems_(body);
        items.forEach(function (item) {
          const pid = safe_(item.product_id || item.id);
          if (pid) resultMap[String(pid)] = item;
        });
      } else if (raw.code === 404 || raw.code === 400) {
        writeDiagnostic('warning', 'Products', '/v3/product/info/list вернул HTTP ' + raw.code + '. Продолжаем без enrichment name/status.', '/v3/product/info/list', 'Products', 'Проверьте доступность endpoint и формат payload');
      } else {
        throw new Error('HTTP ' + raw.code + ' на /v3/product/info/list');
      }
    });
    return resultMap;
  } catch (err) {
    writeDiagnostic('warning', 'Products', 'Не удалось обогатить товары через /v3/product/info/list: ' + sanitizeError_(err), '/v3/product/info/list', 'Products', 'Проверьте права API и endpoint');
    return null;
  }
}

function ozonSellerPostWithFallback_(candidateRequests, options) {
  const opts = options || {};
  const errors = [];
  for (let i = 0; i < candidateRequests.length; i++) {
    const req = candidateRequests[i];
    const raw = ozonSellerPostRaw_(req.path, req.payload || {});
    if (raw.code >= 200 && raw.code < 300) {
      return { path: req.path, response: raw.json || {} };
    }

    const short = 'HTTP ' + raw.code + ': ' + shorten_(raw.text || '', 220);
    errors.push(req.path + ' -> ' + short);

    if (raw.code === 404 || raw.code === 400) {
      if (i < candidateRequests.length - 1 && opts.logFallbackWarnings) {
        const next = candidateRequests[i + 1].path;
        writeDiagnostic('warning', opts.component || 'Seller API', req.path + ' returned ' + raw.code + ', trying ' + next, req.path, opts.targetTab || 'Diagnostics', 'Проверка fallback endpoint');
      }
      continue;
    }

    if (raw.code === 401 || raw.code === 403) {
      throw new Error('HTTP ' + raw.code + ': проверьте Client-Id, Api-Key и права API. Endpoint: ' + req.path);
    }
    if (raw.code === 429) {
      throw new Error('HTTP 429: лимит запросов Ozon. Endpoint: ' + req.path);
    }
    if (raw.code >= 500) {
      throw new Error('HTTP ' + raw.code + ': временная ошибка Ozon. Endpoint: ' + req.path);
    }
  }
  throw new Error('Не удалось получить данные product list. Проверены endpoint: ' + errors.join(' | '));
}

function ozonSellerPostRaw_(path, payload) {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('OZON_SELLER_CLIENT_ID');
  const apiKey = props.getProperty('OZON_SELLER_API_KEY');
  if (!clientId || !apiKey) throw new Error('Отсутствуют ключи Seller API. Заполните OZON_SELLER_CLIENT_ID и OZON_SELLER_API_KEY.');

  const response = UrlFetchApp.fetch(OZON_SELLER_BASE_URL + path, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Client-Id': clientId, 'Api-Key': apiKey },
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true,
  });

  const text = response.getContentText() || '';
  let json = null;
  try { json = text ? JSON.parse(text) : null; } catch (e) { json = null; }
  return { code: response.getResponseCode(), text: text, json: json };
}

function updateStocks() {
  createSheets();
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName('Stocks');
  const loadedAt = nowIso_();
  const date = formatDate_(new Date(), 'yyyy-MM-dd');

  try {
    const stockResponse = getStocksPage_(1000, true);
    const endpoint = stockResponse.path;
    const items = extractStockItems_(stockResponse.response || {});
    const rows = items.map(function (item) {
      const normalized = normalizeStockItem_(item, date, loadedAt);
      return [
        normalized.date,
        normalized.sku,
        normalized.offer_id,
        normalized.warehouse_id,
        normalized.warehouse_name,
        normalized.available_stock,
        normalized.reserved_stock,
        normalized.days_of_stock,
        normalized.turnover_rate,
        normalized.stockout_risk,
        normalized.loaded_at,
      ];
    });

    clearDataRows_(sheet);
    if (rows.length > 0) sheet.getRange(2, 1, rows.length, SHEETS_SCHEMA.Stocks.length).setValues(rows);

    writeDiagnostic('info', 'Stocks', 'Остатки обновлены: ' + rows.length + ' строк. Рабочий endpoint: ' + endpoint, endpoint, 'Stocks', 'Готово');
    ui.alert('Обновление завершено');
  } catch (err) {
    writeDiagnostic('error', 'Stocks', sanitizeError_(err), 'stocks fallback', 'Stocks', 'Проверьте endpoint, права API и ключи');
    ui.alert('Ошибка: смотри вкладку Diagnostics');
  }
}

function getStocksPage_(limit, logFallbackWarnings) {
  const candidates = [
    { path: '/v2/analytics/stock_on_warehouses', payload: { limit: limit || 1000, offset: 0, warehouse_type: 'ALL' } },
    { path: '/v1/analytics/stock_on_warehouses', payload: { limit: limit || 1000, offset: 0, warehouse_type: 'ALL' } },
    { path: '/v2/product/info/stocks', payload: { filter: {}, limit: limit || 1000 } },
  ];
  return ozonSellerPostWithFallback_(candidates, {
    component: 'Stocks',
    targetTab: 'Stocks',
    logFallbackWarnings: !!logFallbackWarnings,
  });
}
function runYesterday() {
  createSheets();
  try {
    const yesterday = shiftDate_(new Date(), -1);
    const y = formatDate_(yesterday, 'yyyy-MM-dd');
    updateProducts();
    updateStocks();
    updateSalesAndFinance(y, y);
    updateAdsDaily(y, y);
    buildForecasts();
    updateSummary();
    writeDiagnostic('info', 'Run', 'Обновление за вчера завершено.', '', '', 'Готово');
    SpreadsheetApp.getUi().alert('Обновление завершено');
  } catch (err) {
    writeDiagnostic('error', 'Run', sanitizeError_(err), '', '', 'Проверьте Diagnostics');
    SpreadsheetApp.getUi().alert('Ошибка: смотри вкладку Diagnostics');
  }
}
function runLast30Days() {
  createSheets();
  try {
    const today = new Date();
    const start = formatDate_(shiftDate_(today, -30), 'yyyy-MM-dd');
    const end = formatDate_(shiftDate_(today, -1), 'yyyy-MM-dd');
    updateProducts();
    updateStocks();
    updateSalesAndFinance(start, end);
    updateAdsDaily(start, end);
    buildForecasts();
    updateSummary();
    writeDiagnostic('info', 'Run', 'Обновление последних 30 дней завершено: ' + start + ' - ' + end, '', '', 'Готово');
    SpreadsheetApp.getUi().alert('Обновление завершено');
  } catch (err) {
    writeDiagnostic('error', 'Run', sanitizeError_(err), '', '', 'Проверьте Diagnostics');
    SpreadsheetApp.getUi().alert('Ошибка: смотри вкладку Diagnostics');
  }
}
function buildForecasts() {
  createSheets();
  try {
    const funnelRows = readSheetObjects_('Raw_Funnel_Daily');
    const financeRows = readSheetObjects_('Finance');
    const metricsByDate = aggregateDailyMetrics_(funnelRows, financeRows);
    const dates = Object.keys(metricsByDate).sort();
    const horizons = [7, 14, 30, 60];
    const metricNames = ['orders', 'ordered_units', 'gross_revenue', 'net_revenue'];
    const reliability = dates.length < 14 ? 'low' : (dates.length < 60 ? 'medium' : 'high');
    const baseDate = dates.length ? dates[dates.length - 1] : '';
    const generatedAt = nowIso_();
    const rows = [];

    metricNames.forEach(function (metric) {
      const series = dates.map(function (d) { return Number(metricsByDate[d][metric] || 0); }).filter(function (v) { return !isNaN(v); });
      if (!series.length || series.every(function (v) { return v === 0; })) return;
      horizons.forEach(function (horizon) {
        const avg = rollingAverage_(series, Math.min(horizon, series.length));
        if (avg === null) return;
        rows.push([
          generatedAt, baseDate, horizon, 'ALL', metric,
          round2_(avg * 0.8 * horizon),
          round2_(avg * horizon),
          round2_(avg * 1.2 * horizon),
          'Rolling average based forecast from Raw_Funnel_Daily/Finance',
          reliability,
        ]);
      });
    });

    const sheet = SpreadsheetApp.getActive().getSheetByName('Forecasts');
    clearDataRows_(sheet);
    if (rows.length) sheet.getRange(2, 1, rows.length, SHEETS_SCHEMA.Forecasts.length).setValues(rows);

    writeDiagnostic('info', 'Forecasts', 'Построены прогнозы: ' + rows.length + ' строк.', '', 'Forecasts', 'Готово');
    buildAlgorithmSignals();
    buildRecommendations();
    SpreadsheetApp.getUi().alert('Обновление завершено');
  } catch (err) {
    writeDiagnostic('error', 'Forecasts', sanitizeError_(err), '', 'Forecasts', 'Проверьте исходные вкладки и формат данных');
    SpreadsheetApp.getUi().alert('Ошибка: смотри вкладку Diagnostics');
  }
}
function buildAlgorithmSignals() {
  createSheets();
  try {
    const products = readSheetObjects_('Products');
    const stocks = readSheetObjects_('Stocks');
    const funnel = readSheetObjects_('Raw_Funnel_Daily');
    const finance = readSheetObjects_('Finance');
    const ads = readSheetObjects_('Raw_Ads_Daily');
    const signals = [];
    const bySku = buildSkuSummary_(products, stocks, funnel, finance);
    const today = formatDate_(new Date(), 'yyyy-MM-dd');

    Object.keys(bySku).forEach(function (sku) {
      const s = bySku[sku];
      if (s.stockoutRiskHigh || s.availableStockLow) {
        signals.push([today, sku, 'stock', 'Низкий остаток', 'Остатки низкие или stockout_risk высокий', 'Возможна просадка продаж/видимости из-за дефицита', 'high', 'stock=' + s.availableStock + ', risk=' + s.stockoutRisk, 'Пополнить склад и проверить поставку', 'high']);
      }
      if (s.salesDrop && s.availableStock > 0) {
        signals.push([today, sku, 'sales', 'Продажи падают при наличии остатков', 'Последние 7 дней хуже предыдущих 7 дней', 'Возможны проблемы цены/карточки/конкуренции/отзывов', 'medium', 'orders7=' + s.orders7 + ', ordersPrev7=' + s.ordersPrev7, 'Проверить цену, контент карточки и отзывы', 'medium']);
      }
      if (s.salesGrow && s.availableStockLow) {
        signals.push([today, sku, 'stock', 'Рост продаж при низком остатке', 'Продажи растут, остаток низкий', 'Высокий риск out-of-stock и потери выручки', 'high', 'orders7=' + s.orders7 + ', stock=' + s.availableStock, 'Срочно пополнить склад', 'high']);
      }
      if (s.grossGrow && s.netDrop) {
        signals.push([today, sku, 'finance', 'Рост gross_revenue при падении net_revenue', 'Выручка растет, чистая выручка падает', 'Могли вырасти комиссии, логистика или скидки', 'medium', 'gross7=' + s.gross7 + ', net7=' + s.net7, 'Проверить структуру затрат и условия продаж', 'medium']);
      }
      if (s.hasProduct && s.orders30 === 0) {
        signals.push([today, sku, 'content', 'Товар без продаж', 'Товар есть в Products, но нет продаж за 30 дней', 'Проблема в видимости/цене/контенте/наличии', 'medium', 'orders30=0', 'Проверить карточку, цену, рекламу и остатки', 'medium']);
      }
    });
    const adSummary = summarizeAds_(ads);
    adSummary.forEach(function (a) {
      if (a.lowCtr) signals.push([today, '', 'ads', 'Низкий CTR кампании', 'CTR ниже порога', 'Возможна слабая релевантность креатива или ставки', 'medium', 'campaign=' + a.campaignId + ', ctr=' + a.ctr, 'Проверить креативы, ставки и релевантность', 'medium']);
      if (a.cpcGrow) signals.push([today, '', 'ads', 'Рост CPC', 'CPC растет при том же/ниже CTR', 'Возможна конкуренция или завышенные ставки', 'medium', 'campaign=' + a.campaignId + ', cpc=' + a.cpc, 'Скорректировать ставки и сегменты', 'medium']);
      if (a.spendNoReturn) signals.push([today, '', 'ads', 'Расход без отдачи', 'spend растет, orders/revenue не растут', 'Неэффективное распределение бюджета', 'high', 'campaign=' + a.campaignId + ', spend=' + a.spend + ', orders=' + a.orders + ', revenue=' + a.revenue, 'Снизить/перераспределить бюджет', 'high']);
      if (a.lowRoasOrHighDrr) signals.push([today, '', 'ads', 'Низкий ROAS / высокий DRR', 'ROAS низкий или DRR высокий', 'Низкая окупаемость рекламы', 'high', 'campaign=' + a.campaignId + ', roas=' + a.roas + ', drr=' + a.drr, 'Оптимизировать кампанию и бюджет', 'high']);
    });

    const sheet = SpreadsheetApp.getActive().getSheetByName('Algorithm_Signals');
    clearDataRows_(sheet);
    if (signals.length) sheet.getRange(2, 1, signals.length, SHEETS_SCHEMA.Algorithm_Signals.length).setValues(signals);
    writeDiagnostic('info', 'Signals', 'Сформированы сигналы: ' + signals.length + ' строк.', '', 'Algorithm_Signals', 'Готово');
  } catch (err) {
    writeDiagnostic('error', 'Signals', sanitizeError_(err), '', 'Algorithm_Signals', 'Проверьте исходные вкладки и формат данных');
  }
}
function buildRecommendations() {
  createSheets();
  try {
    const signals = readSheetObjects_('Algorithm_Signals');
    const now = nowIso_();
    const rows = signals.map(function (s) {
      const priority = safe_(s.priority || 'medium');
      const area = safe_(s.signal_group || 'sales');
      const issueText = safe_(s.signal_name || s.observed_change || 'Наблюдаемое отклонение');
      const actionText = safe_(s.recommended_test || 'Провести проверку гипотезы');
      if (area === 'ads' && issueText.indexOf('Низкий CTR') >= 0) {
        return [now, priority, 'ads', safe_(s.sku || 'campaign'), issueText, 'Campaign cleanup и тест новых креативов', 'medium', 'low', 'Маркетинг', 'new'];
      }
      if (area === 'ads' && issueText.indexOf('Рост CPC') >= 0) {
        return [now, priority, 'ads', safe_(s.sku || 'campaign'), issueText, 'Bid adjustment и пересмотр таргетинга', 'medium', 'low', 'Маркетинг', 'new'];
      }
      if (area === 'ads' && (issueText.indexOf('ROAS') >= 0 || issueText.indexOf('Расход без отдачи') >= 0)) {
        return [now, priority, 'ads', safe_(s.sku || 'campaign'), issueText, 'Budget reallocation и search phrase tests', 'high', 'medium', 'Маркетинг', 'new'];
      }
      return [
        now,
        priority,
        area,
        safe_(s.sku || ''),
        issueText,
        actionText,
        priority === 'high' ? 'high' : 'medium',
        priority === 'high' ? 'medium' : 'low',
        'Команда',
        'new',
      ];
    });
    const sheet = SpreadsheetApp.getActive().getSheetByName('Recommendations');
    clearDataRows_(sheet);
    if (rows.length) sheet.getRange(2, 1, rows.length, SHEETS_SCHEMA.Recommendations.length).setValues(rows);
    writeDiagnostic('info', 'Recommendations', 'Сформированы рекомендации: ' + rows.length + ' строк.', '', 'Recommendations', 'Готово');
  } catch (err) {
    writeDiagnostic('error', 'Recommendations', sanitizeError_(err), '', 'Recommendations', 'Проверьте вкладку Algorithm_Signals');
  }
}
function ozonPerformanceToken_() {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('OZON_PERFORMANCE_CLIENT_ID');
  const clientSecret = props.getProperty('OZON_PERFORMANCE_CLIENT_SECRET');
  if (!clientId || !clientSecret) {
    writeDiagnostic('warning', 'Performance API', 'Ключи Performance API не заполнены. Пропускаем загрузку рекламы.', '', 'Raw_Ads_Daily', 'Заполните OZON_PERFORMANCE_CLIENT_ID и OZON_PERFORMANCE_CLIENT_SECRET');
    return null;
  }

  const url = OZON_PERFORMANCE_BASE_URL + '/api/client/token';
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ client_id: clientId, client_secret: clientSecret, grant_type: 'client_credentials' }),
    muteHttpExceptions: true,
  });
  const code = response.getResponseCode();
  const text = response.getContentText() || '';
  if (code >= 200 && code < 300) {
    try {
      const json = JSON.parse(text);
      const token = json.access_token || (json.result && json.result.access_token);
      if (token) return { token: token, baseUrl: OZON_PERFORMANCE_BASE_URL };
    } catch (e) {}
  }
  if (code === 401 || code === 403) throw new Error('Performance API: проверьте client_id/client_secret и права доступа.');
  if (code >= 400) throw new Error('Performance token HTTP ' + code + ': ' + shorten_(text, 160));
  writeDiagnostic('warning', 'Performance API', 'Не удалось получить token Performance API.', '', 'Raw_Ads_Daily', 'Проверьте endpoint и права доступа');
  return null;
}
function ozonPerformanceRequest_(path, payloadOrParams, method) {
  const auth = ozonPerformanceToken_();
  if (!auth || !auth.token) return null;
  const raw = ozonPerformanceRawRequest_(auth, path, payloadOrParams, method);
  if (raw.code < 200 || raw.code >= 300) throw new Error('Performance API HTTP ' + raw.code + ': ' + shorten_(raw.text, 180));
  return { data: raw.json || {}, endpoint: path };
}
function ozonPerformanceRawRequest_(auth, path, payloadOrParams, method) {
  const httpMethod = (method || 'post').toLowerCase();
  let url = auth.baseUrl + path;
  const options = {
    method: httpMethod,
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + auth.token },
    muteHttpExceptions: true,
  };
  if (httpMethod === 'get' && payloadOrParams) {
    const query = Object.keys(payloadOrParams).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(payloadOrParams[k]); }).join('&');
    if (query) url += '?' + query;
  } else if (payloadOrParams) {
    options.payload = JSON.stringify(payloadOrParams);
  }
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText() || '';
  let json = {};
  try { json = text ? JSON.parse(text) : {}; } catch (e) { json = {}; }
  return { code: code, text: text, json: json, endpoint: path };
}
function ozonPerformanceGet_(path, queryParams) {
  const auth = ozonPerformanceToken_();
  if (!auth || !auth.token) return { code: 0, text: '', json: {} };
  let url = OZON_PERFORMANCE_BASE_URL + path;
  if (queryParams && Object.keys(queryParams).length) {
    const query = Object.keys(queryParams)
      .map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(queryParams[k]); })
      .join('&');
    url += '?' + query;
  }
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + auth.token,
      Accept: 'application/json',
    },
    muteHttpExceptions: true,
  });
  const code = response.getResponseCode();
  const text = response.getContentText() || '';
  let json = {};
  try { json = text ? JSON.parse(text) : {}; } catch (e) { json = {}; }
  return { code: code, text: text, json: json };
}
function updateAdsDaily(startDate, endDate) {
  createSheets();
  const sheet = SpreadsheetApp.getActive().getSheetByName('Raw_Ads_Daily');
  const loadedAt = nowIso_();
  try {
    const props = PropertiesService.getScriptProperties();
    const keysPresent = !!(props.getProperty('OZON_PERFORMANCE_CLIENT_ID') && props.getProperty('OZON_PERFORMANCE_CLIENT_SECRET'));
    writeDiagnostic('info', 'Performance API', 'Запуск updateAdsDaily: period ' + startDate + ' - ' + endDate + ', performance keys present: ' + (keysPresent ? 'yes' : 'no'), '', 'Raw_Ads_Daily', 'Старт');

    const tokenInfo = ozonPerformanceToken_();
    writeDiagnostic('info', 'Performance API', 'Token received: ' + (tokenInfo && tokenInfo.token ? 'yes' : 'no'), '', 'Raw_Ads_Daily', 'Проверка токена');
    if (!tokenInfo || !tokenInfo.token) return;

    let campaignResult = getCampaignsFromSheet_();
    if (!campaignResult.items.length) {
      diagnosePerformanceCampaigns();
      campaignResult = getCampaignsFromSheet_();
    }
    const campaigns = campaignResult.items || [];
    writeDiagnostic('info', 'Performance API', 'Campaigns endpoint used: ' + campaignResult.path + ', campaigns count: ' + campaigns.length, campaignResult.path, 'Raw_Ads_Daily', 'Проверка кампаний');
    if (!campaigns.length) {
      writeDiagnostic('warning', 'Performance API', 'Campaigns count = 0. Обновление Raw_Ads_Daily пропущено.', campaignResult.path, 'Raw_Ads_Daily', 'Проверьте наличие кампаний и расходы за период');
      SpreadsheetApp.getUi().alert('Performance API доступен, но кампании не найдены. Проверьте, что ключ создан в том же Ozon Performance аккаунте.');
      return;
    }

    const statsResult = getAdsStatisticsWithFallback_(startDate, endDate, campaigns);
    const items = extractAdsItems_(statsResult.data || {});
    writeDiagnostic('info', 'Performance API', 'Statistics endpoint used: ' + statsResult.path + ', statistics rows received: ' + items.length, statsResult.path, 'Raw_Ads_Daily', 'Проверка статистики');
    const rows = items.map(function (item) {
      const ctr = toNum_(item.ctr);
      const spend = toNum_(item.spend || item.cost);
      const clicks = toNum_(item.clicks);
      const avgCpc = clicks > 0 ? spend / clicks : '';
      const revenue = item.revenue !== undefined ? item.revenue : '';
      const orders = item.orders !== undefined ? item.orders : '';
      const roas = (revenue !== '' && spend > 0) ? toNum_(revenue) / spend : '';
      const drr = (revenue !== '' && toNum_(revenue) > 0) ? spend / toNum_(revenue) : '';
      return [
        safe_(item.date || startDate),
        safe_(item.campaign_id || item.id),
        safe_(item.campaign_name || item.name),
        safe_(item.campaign_type || item.type),
        safe_(item.placement || ''),
        safe_(item.sku || ''),
        safe_(item.offer_id || ''),
        safe_(item.impressions || ''),
        safe_(clicks || ''),
        safe_(ctr || ''),
        safe_(spend || ''),
        safe_(avgCpc === '' ? '' : round2_(avgCpc)),
        safe_(orders),
        safe_(revenue),
        safe_(drr === '' ? '' : round2_(drr)),
        safe_(roas === '' ? '' : round2_(roas)),
        safe_(item.search_phrase || item.phrase || ''),
        statsResult.path,
        loadedAt,
      ];
    });
    if (!rows.length) {
      writeDiagnostic('warning', 'Performance API', 'Performance API вернул 0 строк рекламы за период', statsResult.path, 'Raw_Ads_Daily', 'Проверьте наличие активных кампаний, расходы за период, Performance API аккаунт и используемые endpoints');
      return;
    }
    clearDataRows_(sheet);
    sheet.getRange(2, 1, rows.length, SHEETS_SCHEMA.Raw_Ads_Daily.length).setValues(rows);
    writeDiagnostic('info', 'Performance API', 'Normalized rows written to Raw_Ads_Daily: ' + rows.length, statsResult.path, 'Raw_Ads_Daily', 'Готово');
    const missingAttribution = rows.some(function (r) { return r[12] === '' || r[13] === ''; });
    if (missingAttribution) writeDiagnostic('warning', 'Performance API', 'Часть строк без orders/revenue attribution. ROAS/DRR оставлены пустыми.', statsResult.path, 'Raw_Ads_Daily', 'Это возможно для некоторых отчетов API');
    writeDiagnostic('info', 'Performance API', 'Реклама обновлена: ' + rows.length + ' строк.', statsResult.path, 'Raw_Ads_Daily', 'Готово');
  } catch (err) {
    writeDiagnostic('error', 'Performance API', sanitizeError_(err), '', 'Raw_Ads_Daily', 'Проверьте keys, endpoint и права');
  }
}
function getCampaignsFromSheet_() {
  const rows = readSheetObjects_('Ad_Campaigns');
  return {
    path: 'Ad_Campaigns',
    items: rows.map(function (r) { return { id: r.campaign_id, campaign_id: r.campaign_id }; }).filter(function (x) { return !!x.id; }),
  };
}
function updateAdsLast30Days() {
  const today = new Date();
  const start = formatDate_(shiftDate_(today, -30), 'yyyy-MM-dd');
  const end = formatDate_(shiftDate_(today, -1), 'yyyy-MM-dd');
  updateAdsDaily(start, end);
  SpreadsheetApp.getUi().alert('Обновление завершено');
}
function updateAdsDailyMenu() {
  createSheets();
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('OZON_PERFORMANCE_CLIENT_ID');
  const clientSecret = props.getProperty('OZON_PERFORMANCE_CLIENT_SECRET');
  if (!clientId || !clientSecret) {
    writeDiagnostic('warning', 'Performance API', 'Ключи Performance API не заполнены. Обновление рекламы пропущено.', '', 'Raw_Ads_Daily', 'Заполните OZON_PERFORMANCE_CLIENT_ID и OZON_PERFORMANCE_CLIENT_SECRET');
    SpreadsheetApp.getUi().alert('Ошибка: не заполнены ключи Performance API. Смотри Diagnostics.');
    return;
  }
  try {
    const today = new Date();
    const start = formatDate_(shiftDate_(today, -7), 'yyyy-MM-dd');
    const end = formatDate_(shiftDate_(today, -1), 'yyyy-MM-dd');
    updateAdsDaily(start, end);
    writeDiagnostic('info', 'Performance API', 'Обновление рекламы через меню выполнено за период: ' + start + ' - ' + end, '', 'Raw_Ads_Daily', 'Готово');
    SpreadsheetApp.getUi().alert('Обновление завершено');
  } catch (err) {
    writeDiagnostic('error', 'Performance API', sanitizeError_(err), '', 'Raw_Ads_Daily', 'Проверьте настройки и endpoint');
    SpreadsheetApp.getUi().alert('Ошибка: смотри вкладку Diagnostics');
  }
}
function extractAdsItems_(responseJson) {
  if (responseJson.result && responseJson.result.items) return responseJson.result.items;
  if (responseJson.result && responseJson.result.rows) return responseJson.result.rows;
  if (Array.isArray(responseJson.result)) return responseJson.result;
  if (Array.isArray(responseJson.items)) return responseJson.items;
  return [];
}
function getCampaignsWithFallback_() {
  const paths = ['/api/client/campaign', '/api/client/campaigns'];
  for (let i = 0; i < paths.length; i++) {
    const path = paths[i];
    try {
      const res = ozonPerformanceRequest_(path, null, 'get');
      const items = extractCampaignItems_(res && res.data ? res.data : {});
      return { path: path, items: items };
    } catch (err) {
      const msg = sanitizeError_(err);
      writeDiagnostic('warning', 'Performance API', 'Campaign endpoint failed: ' + path + ' | ' + msg, path, 'Raw_Ads_Daily', 'Пробуем fallback endpoint');
    }
  }
  throw new Error('Не удалось получить кампании Performance API.');
}
function getAdsStatisticsWithFallback_(startDate, endDate, campaigns) {
  const campaignIds = campaigns.map(function (c) { return c.id || c.campaign_id; }).filter(function (x) { return !!x; });
  const paths = ['/api/client/statistics', '/api/client/campaign/statistics'];
  for (let i = 0; i < paths.length; i++) {
    const path = paths[i];
    const payload = { date_from: startDate, date_to: endDate, campaign_ids: campaignIds };
    try {
      const res = ozonPerformanceRequest_(path, payload, 'post');
      return { path: path, data: res && res.data ? res.data : {} };
    } catch (err) {
      const msg = sanitizeError_(err);
      writeDiagnostic('warning', 'Performance API', 'Statistics endpoint failed: ' + path + ' | ' + msg, path, 'Raw_Ads_Daily', 'Пробуем fallback endpoint');
    }
  }
  throw new Error('Не удалось получить статистику рекламы Performance API.');
}
function extractCampaignItems_(responseJson) {
  if (responseJson.result && responseJson.result.items) return responseJson.result.items;
  if (Array.isArray(responseJson.result)) return responseJson.result;
  if (Array.isArray(responseJson.items)) return responseJson.items;
  return [];
}
function diagnosePerformanceCampaigns() {
  createSheets();
  const sheet = SpreadsheetApp.getActive().getSheetByName('Ad_Campaigns');
  clearDataRows_(sheet);
  const result = fetchPerformanceCampaigns_();
  const rows = result.rows;
  const found = result.count;

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, SHEETS_SCHEMA.Ad_Campaigns.length).setValues(rows);
    SpreadsheetApp.getUi().alert('Найдено рекламных кампаний: ' + found);
  } else {
    writeDiagnostic('warning', 'Performance API', 'Проверьте, что Performance API key создан именно в том рекламном аккаунте, где видны кампании в интерфейсе Ozon', '/api/client/campaign*', 'Ad_Campaigns', 'Кампании через API не найдены');
    SpreadsheetApp.getUi().alert('Performance API доступен, но кампании не найдены. Проверьте, что ключ создан в том же Ozon Performance аккаунте.');
  }
}
function fetchPerformanceCampaigns_() {
  const endpoint = '/api/client/campaign';
  const advTypes = [null, 'SKU', 'BANNER', 'SEARCH_PROMO'];
  const states = [null, 'CAMPAIGN_STATE_RUNNING', 'CAMPAIGN_STATE_PLANNED', 'CAMPAIGN_STATE_STOPPED', 'CAMPAIGN_STATE_INACTIVE', 'CAMPAIGN_STATE_ARCHIVED', 'CAMPAIGN_STATE_MODERATION_DRAFT', 'CAMPAIGN_STATE_MODERATION_IN_PROGRESS', 'CAMPAIGN_STATE_MODERATION_FAILED', 'CAMPAIGN_STATE_FINISHED'];
  const maxPages = 20;
  const pageSize = 100;
  const loadedAt = nowIso_();
  const rows = [];
  let count = 0;

  function runAttempt(extraParams) {
    let page = 1;
    while (page <= maxPages) {
      const params = Object.assign({}, extraParams || {}, { page: page, pageSize: pageSize });
      const raw = ozonPerformanceGet_(endpoint, params);
      const list = (raw.json && raw.json.list && Array.isArray(raw.json.list)) ? raw.json.list : [];
      writeDiagnostic(
        raw.code >= 200 && raw.code < 300 ? 'info' : 'warning',
        'Performance Campaigns',
        'query=' + JSON.stringify(params) + ', http=' + raw.code + ', count=' + list.length + ', preview=' + shorten_(raw.text || '', 300),
        endpoint,
        'Ad_Campaigns',
        'Диагностика кампаний'
      );
      if (raw.code < 200 || raw.code >= 300) break;
      list.forEach(function (campaign) {
        rows.push(normalizeCampaignRow_(campaign, loadedAt, endpoint, params));
        count += 1;
      });
      if (list.length < pageSize) break;
      page += 1;
    }
  }

  runAttempt({});
  if (count === 0) advTypes.forEach(function (t) { if (t) runAttempt({ advObjectType: t }); });
  if (count === 0) states.forEach(function (s) { if (s) runAttempt({ state: s }); });
  return { rows: rows, count: count };
}
function normalizeCampaignRow_(campaign, loadedAt, endpoint, params) {
  return [
    loadedAt,
    endpoint,
    JSON.stringify(params || {}),
    safe_(campaign.id || campaign.campaignId || campaign.campaign_id || ''),
    safe_(campaign.title || campaign.name || campaign.campaignName || campaign.campaign_name || ''),
    safe_(campaign.paymentType || campaign.advObjectType || ''),
    safe_(campaign.advObjectType || campaign.adv_object_type || ''),
    safe_(campaign.state || ''),
    safe_(campaign.status || campaign.state || ''),
    safe_(campaign.paymentType || campaign.payment_type || ''),
    safe_(campaign.dailyBudget || campaign.daily_budget || ''),
    safe_(campaign.budget || ''),
    safe_(campaign.fromDate || campaign.startDate || campaign.start_date || ''),
    safe_(campaign.toDate || campaign.endDate || campaign.end_date || ''),
    shorten_(JSON.stringify(campaign || {}), 5000),
  ];
}
function normalizeOfferId_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim().toUpperCase().replace(/\s+/g, '');
}
function loadUnitCosts_() {
  const rows = readSheetObjects_('Unit_Costs');
  const map = {};
  rows.forEach(function (r) {
    const offer = normalizeOfferId_(r.offer_id);
    const cost = parseNumber_(r.unit_cost);
    if (offer && cost > 0) map[offer] = cost;
  });
  return map;
}
function seedUnitCostsIfEmpty_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Unit_Costs');
  if (!sheet) return;
  if (sheet.getLastRow() > 1) return;
  const seed = [
    ['SJ10',510],['SJ42',535],['SJK7',456],['SJK8',537],['SJK45',718],['SJK48',754],['SJK5',865],['SJ38',548],['SJ6',444],['SJ1',585],['SJ20',5],['SJK6',810],['SJ57',177],['SJK44',525],['SJK43',499],['SJK47',611],['SJ3',301],['SJ4',318],['SJ27',462],['SJ2',362],['SJ17',348],['SJ18',559],['SJ32',252],['SJ5',472],['SJ21',347],['SJK10',622],['SJ7',307],['SJ37',268],['SJK46',374],['SJ51',392],['SJ25',486],['SJ28',448],['SJ12',469],['SJK9',502],['SJ30',413],['SJ24',593],['SJ52',754],['SJK49',520],['SJ11',525],['SJ39',275],['SJK2',138],['SJ22',526],['SJ40',346],['SJ23',723],['SJK3',499],['SJ34',323],['SJ36',365],['SJK4',477],['SJ33',274],['SJ31',302],['SJ56',216],['SJ41',610],['SJ50',467],['SJ55',349],['SJN143',498],['SJKS1',312],['SJ16',329],['SJ53',379],['SJK13',567],['SJ54',578],['SJ58',499],['SJ59',571],
  ];
  const now = nowIso_();
  const values = seed.map(function (x) { return [x[0], x[1], 'RUB', now, 'seed']; });
  sheet.getRange(2, 1, values.length, 5).setValues(values);
}
function testPerformanceConnection() {
  createSheets();
  const props = PropertiesService.getScriptProperties();
  const hasKeys = !!(props.getProperty('OZON_PERFORMANCE_CLIENT_ID') && props.getProperty('OZON_PERFORMANCE_CLIENT_SECRET'));
  writeDiagnostic('info', 'Performance API', 'testPerformanceConnection: performance keys present: ' + (hasKeys ? 'yes' : 'no'), '', 'Raw_Ads_Daily', 'Проверка ключей');
  if (!hasKeys) {
    SpreadsheetApp.getUi().alert('Ошибка: не заполнены ключи Performance API. Смотри Diagnostics.');
    return;
  }
  try {
    const tokenInfo = ozonPerformanceToken_();
    const tokenReceived = !!(tokenInfo && tokenInfo.token);
    writeDiagnostic('info', 'Performance API', 'testPerformanceConnection: token received: ' + (tokenReceived ? 'yes' : 'no'), '', 'Raw_Ads_Daily', 'Проверка токена');
    if (!tokenReceived) throw new Error('Token не получен.');
    const campaigns = getCampaignsWithFallback_();
    writeDiagnostic('info', 'Performance API', 'testPerformanceConnection: campaigns count: ' + campaigns.items.length, campaigns.path, 'Raw_Ads_Daily', 'Готово');
    SpreadsheetApp.getUi().alert('Performance API работает, найдено кампаний: ' + campaigns.items.length);
  } catch (err) {
    writeDiagnostic('error', 'Performance API', sanitizeError_(err), '', 'Raw_Ads_Daily', 'Проверьте keys, права и endpoint');
    SpreadsheetApp.getUi().alert('Ошибка Performance API: смотри Diagnostics');
  }
}

function updateSalesAndFinance(startDate, endDate) {
  createSheets();
  const loadedAt = nowIso_();
  const funnelSheet = SpreadsheetApp.getActive().getSheetByName('Raw_Funnel_Daily');
  const financeSheet = SpreadsheetApp.getActive().getSheetByName('Finance');

  try {
    const postingsResult = getPostingsWithFallback_(startDate, endDate);
    const endpoint = postingsResult.path;
    const postings = extractPostingItems_(postingsResult.response || {});

    const aggregated = aggregatePostingsByDateSku_(postings);
    const funnelRows = [];
    const financeRows = [];
    Object.keys(aggregated).forEach(function (key) {
      const row = aggregated[key];
      funnelRows.push([
        row.date, row.sku, row.offer_id, row.product_id, '',
        '', '', '',
        row.orders,
        row.ordered_units,
        row.gross_revenue,
        '', '', '',
        '', '', '', '',
        endpoint,
        loadedAt,
      ]);
      financeRows.push([
        row.date, row.sku, row.offer_id,
        row.orders,
        row.gross_revenue,
        '',
        '',
        '',
        '',
        '',
        row.net_revenue,
        '',
        loadedAt,
      ]);
    });

    clearDataRows_(funnelSheet);
    clearDataRows_(financeSheet);
    if (funnelRows.length) funnelSheet.getRange(2, 1, funnelRows.length, SHEETS_SCHEMA.Raw_Funnel_Daily.length).setValues(funnelRows);
    if (financeRows.length) financeSheet.getRange(2, 1, financeRows.length, SHEETS_SCHEMA.Finance.length).setValues(financeRows);

    writeDiagnostic('info', 'Sales&Finance', 'Обновлены Raw_Funnel_Daily: ' + funnelRows.length + ', Finance: ' + financeRows.length + '. Endpoint: ' + endpoint, endpoint, 'Raw_Funnel_Daily,Finance', 'Готово');
  } catch (err) {
    writeDiagnostic('error', 'Sales&Finance', sanitizeError_(err), 'posting/finance fallback', 'Raw_Funnel_Daily,Finance', 'Проверьте endpoint, права API и период');
    throw err;
  }
}
function updateSummary() {
  createSheets();
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Summary');
    const funnel = readSheetObjects_('Raw_Funnel_Daily');
    const finance = readSheetObjects_('Finance');
    const ads = readSheetObjects_('Raw_Ads_Daily');
    const stocks = readSheetObjects_('Stocks');
    const products = readSheetObjects_('Products');
    const recs = readSheetObjects_('Recommendations');
    const unitCostMap = loadUnitCosts_();
    writeDiagnostic('info', 'Summary', 'Unit_Costs loaded count=' + Object.keys(unitCostMap).length, '', 'Summary', 'Unit costs loaded');

    writeDiagnostic('info', 'Summary', 'rows read: funnel=' + funnel.length + ', finance=' + finance.length + ', ads=' + ads.length + ', stocks=' + stocks.length, '', 'Summary', 'Data read');

    const dates = funnel.map(function (r) { return formatDateForSummary_(r.date); }).filter(function (d) { return d !== 'нет данных'; }).sort();
    const periodStart = dates.length ? dates[0] : 'нет данных';
    const periodEnd = dates.length ? dates[dates.length - 1] : 'нет данных';
    const orders = sumByParsed_(funnel, 'orders');
    const units = sumByParsed_(funnel, 'ordered_units');
    const grossRevenue = sumByParsed_(funnel, 'gross_revenue');
    const netRevenue = sumByParsed_(finance, 'net_revenue');
    const cogsResult = calcCogsFromFunnel_(funnel, products, unitCostMap);
    const cogs = cogsResult.total;
    const grossProfit = cogs > 0 ? (grossRevenue - cogs) : null;
    const ozonCommission = sumByParsed_(finance, 'commission');
    const ozonLogistics = sumByParsed_(finance, 'logistics');
    const returnsCost = sumByParsed_(finance, 'returns_cost');
    const adSpend = sumByParsed_(ads, 'spend');
    const adRevenue = sumByParsed_(ads, 'revenue');
    const adOrders = sumByParsed_(ads, 'orders');
    const impressions = sumByParsed_(ads, 'impressions');
    const clicks = sumByParsed_(ads, 'clicks');
    const ctr = impressions > 0 ? clicks / impressions : null;
    const avgCpc = clicks > 0 ? adSpend / clicks : null;
    const roas = adSpend > 0 && adRevenue > 0 ? adRevenue / adSpend : null;
    const drr = adRevenue > 0 ? adSpend / adRevenue : null;
    const avgCheck = orders > 0 ? grossRevenue / orders : null;
    const operatingProfit = grossProfit !== null ? (grossProfit - ozonCommission - ozonLogistics - returnsCost - adSpend) : null;
    const marginPct = (operatingProfit !== null && grossRevenue > 0) ? (operatingProfit / grossRevenue) : null;
    const lowStockRows = lowStockTop_(stocks, products, 10);
    const lowStockCount = lowStockRows.length;

    const productsMap = buildProductsMap_();
    const topRev = topByDimension_(funnel, 'gross_revenue', productsMap, 10);
    const topAds = topAdsByDimension_(ads, productsMap, 10);
    const weakCampaigns = highSpendLowReturnCampaigns_(ads, 5);
    const highPriorityRecs = groupedHighPriorityRecs_(recs, 10);
    const skuPnlRows = calcSkuPnl_(funnel, products, unitCostMap, 20);

    clearDataRows_(sheet);
    const rows = [];
    rows.push(['Период анализа', 'Период', periodStart + ' — ' + periodEnd, '']);
    rows.push(['P&L', 'Выручка gross', grossRevenue || 'нет данных', '']);
    rows.push(['P&L', 'Выручка net', netRevenue || 'нет данных', '']);
    rows.push(['P&L', 'Себестоимость товара', cogs > 0 ? cogs : 'нет данных', 'unit_cost * ordered_units']);
    rows.push(['P&L', 'Валовая прибыль', grossProfit !== null ? grossProfit : 'нет данных', 'gross revenue - cogs']);
    rows.push(['P&L', 'Валовая маржа', grossProfit !== null && grossRevenue > 0 ? (grossProfit / grossRevenue) : 'нет данных', 'gross profit / gross revenue']);
    rows.push(['P&L', 'Комиссия Ozon', ozonCommission || 'нет данных', '']);
    rows.push(['P&L', 'Логистика Ozon', ozonLogistics || 'нет данных', '']);
    rows.push(['P&L', 'Возвраты / удержания', returnsCost || 'нет данных', '']);
    rows.push(['P&L', 'Рекламные расходы', adSpend || 'нет данных', '']);
    rows.push(['P&L', 'Операционная прибыль', operatingProfit !== null ? operatingProfit : 'нет данных', 'gross profit - commission - logistics - returns - ads']);
    rows.push(['P&L', 'Операционная маржа', marginPct !== null ? marginPct : 'нет данных', 'operating profit / gross revenue']);
    rows.push(['P&L', 'ДРР / ACOS', drr !== null ? drr : (adSpend > 0 ? (grossRevenue > 0 ? adSpend / grossRevenue : 'нет данных по revenue attribution') : 'нет данных'), adRevenue > 0 ? 'ad spend / attributed ad revenue' : 'расчёт от общей выручки, не от attribution revenue']);
    rows.push(['P&L', 'ROAS', roas !== null ? roas : (adSpend > 0 ? 'нет данных по attribution revenue' : 'нет данных'), 'ad revenue / ad spend']);
    rows.push(['P&L', 'SKU без себестоимости', cogsResult.missing.length, cogsResult.missing.length ? 'см. Diagnostics' : '']);
    rows.push(['KPI', 'Выручка (gross)', grossRevenue || 'нет данных', '']);
    rows.push(['KPI', 'Выручка (net)', netRevenue || 'нет данных', '']);
    rows.push(['KPI', 'Заказы', orders || 'нет данных', '']);
    rows.push(['KPI', 'Штуки', units || 'нет данных', '']);
    rows.push(['KPI', 'Рекламные расходы', adSpend || 'нет данных', '']);
    rows.push(['KPI', 'Показы / Клики', (impressions || 0) + ' / ' + (clicks || 0), '']);
    rows.push(['KPI', 'CTR', ctr !== null ? ctr : 'нет данных', '%']);
    rows.push(['KPI', 'avg CPC', avgCpc !== null ? avgCpc : 'нет данных', '']);
    rows.push(['KPI', 'DRR / ACOS', drr !== null ? drr : (adSpend > 0 ? 'нет данных по revenue attribution' : 'нет данных'), '']);
    rows.push(['KPI', 'ROAS', roas !== null ? roas : (adSpend > 0 ? 'нет данных по revenue attribution' : 'нет данных'), '']);
    rows.push(['KPI', 'Операционная прибыль', operatingProfit !== null ? operatingProfit : 'нет данных', '']);
    rows.push(['KPI', 'Маржинальность %', marginPct !== null ? marginPct : 'нет данных', '']);
    rows.push(['KPI', 'Рекламные заказы', adOrders || 'нет данных', '']);
    rows.push(['KPI', 'Средний чек', avgCheck !== null ? avgCheck : 'нет данных', '']);
    rows.push(['Stocks', 'Товары с низким остатком', lowStockCount || 'нет данных', '']);

    if (topRev.length) topRev.forEach(function (x) { rows.push(['Топ SKU по выручке', x.key, x.value, '']); });
    else rows.push(['Топ SKU по выручке', 'нет данных', 'нет данных', '']);
    if (topAds.length) topAds.forEach(function (x) { rows.push(['Топ SKU по рекламе', x.key, x.value, '']); });
    else rows.push(['Топ SKU по рекламе', 'нет данных', 'нет данных', '']);
    if (weakCampaigns.length) weakCampaigns.forEach(function (x) { rows.push(['Кампании с высоким расходом и низкой отдачей', x.campaign, x.spend, x.roas === null ? 'есть расходы, нет revenue attribution' : ('ROAS=' + x.roas)]); });
    else if (adSpend > 0 && adRevenue <= 0) rows.push(['Кампании с высоким расходом и низкой отдачей', 'есть расходы, нет revenue attribution', adSpend, 'проверьте attribution']);
    else rows.push(['Кампании с высоким расходом и низкой отдачей', 'нет данных', 'нет данных', '']);
    if (lowStockRows.length) lowStockRows.forEach(function (r) { rows.push(['Низкие остатки', r.key, r.available + ' шт', 'days_of_stock=' + r.days]); });
    else rows.push(['Низкие остатки', 'нет данных', 'нет данных', '']);
    if (highPriorityRecs.length) highPriorityRecs.forEach(function (r) { rows.push(['Рекомендации high priority', r.metric, r.value, r.notes]); });
    else rows.push(['Рекомендации high priority', 'нет данных', 'нет данных', '']);
    if (skuPnlRows.top.length) skuPnlRows.top.forEach(function (r) { rows.push(['P&L по SKU', r.key, round2_(r.profit), 'revenue=' + round2_(r.revenue) + ', cogs=' + round2_(r.cogs) + ', units=' + round2_(r.units) + ', margin=' + round2_(r.margin * 100) + '%']); });
    if (skuPnlRows.loss.length) skuPnlRows.loss.forEach(function (r) { rows.push(['Убыточные SKU', r.key, round2_(r.profit), 'revenue=' + round2_(r.revenue) + ', cogs=' + round2_(r.cogs) + ', units=' + round2_(r.units)]); });

    if (rows.length) sheet.getRange(2, 1, rows.length, SHEETS_SCHEMA.Summary.length).setValues(rows);

    // formatting
    sheet.getRange(1, 1, 1, SHEETS_SCHEMA.Summary.length).setFontWeight('bold');
    sheet.getRange(2, 1, rows.length, 1).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, SHEETS_SCHEMA.Summary.length);
    if (rows.length) {
      sheet.getRange(2, 3, rows.length, 1).setNumberFormat('#,##0.00');
    }
    applySummaryFormats_(sheet, rows.length);
    writeDiagnostic('warning', 'Summary', 'missing unit cost count=' + cogsResult.missing.length, '', 'Summary', 'Missing costs');
    writeDiagnostic('info', 'Summary', 'P&L rows calculated=' + 13 + ', SKU P&L rows calculated=' + (skuPnlRows.top.length + skuPnlRows.loss.length) + ', ad spend used=' + (adSpend > 0 ? 'yes' : 'no') + ', finance rows used count=' + finance.length, '', 'Summary', 'Done');
    cogsResult.missing.slice(0, 100).forEach(function (m) { writeDiagnostic('warning', 'Summary', 'Не найдена себестоимость для SKU/OFFER_ID: ' + m, '', 'Unit_Costs', 'Добавьте unit_cost в Unit_Costs'); });
    writeDiagnostic('info', 'Summary', 'Summary обновлен: ' + rows.length + ' строк.', '', 'Summary', 'Готово');
    SpreadsheetApp.getUi().alert('Обновление завершено');
  } catch (err) {
    writeDiagnostic('error', 'Summary', sanitizeError_(err), '', 'Summary', 'Проверьте исходные вкладки');
    SpreadsheetApp.getUi().alert('Ошибка: смотри вкладку Diagnostics');
  }
}

function getPostingsWithFallback_(startDate, endDate) {
  const payloadV3 = { dir: 'ASC', filter: { since: startDate + 'T00:00:00Z', to: endDate + 'T23:59:59Z' }, limit: 1000, offset: 0, with: { analytics_data: true, financial_data: true } };
  const payloadV2 = { filter: { since: startDate + 'T00:00:00Z', to: endDate + 'T23:59:59Z' }, limit: 1000, offset: 0, with: { analytics_data: true, financial_data: true } };
  const candidates = [
    { path: '/v3/posting/fbo/list', payload: payloadV3 },
    { path: '/v2/posting/fbo/list', payload: payloadV2 },
    { path: '/v3/posting/fbs/list', payload: payloadV3 },
    { path: '/v2/posting/fbs/list', payload: payloadV2 },
  ];
  return ozonSellerPostWithFallback_(candidates, {
    component: 'Sales&Finance',
    targetTab: 'Raw_Funnel_Daily',
    logFallbackWarnings: true,
  });
}
function extractPostingItems_(responseJson) {
  if (responseJson.result && responseJson.result.postings) return responseJson.result.postings;
  if (responseJson.result && responseJson.result.items) return responseJson.result.items;
  if (Array.isArray(responseJson.result)) return responseJson.result;
  if (responseJson.postings) return responseJson.postings;
  if (responseJson.items) return responseJson.items;
  return [];
}
function aggregatePostingsByDateSku_(postings) {
  const map = {};
  postings.forEach(function (posting) {
    const postingDate = safe_(posting.in_process_at || posting.created_at || posting.order_date || '').substring(0, 10);
    const products = posting.products || posting.items || [];
    products.forEach(function (product) {
      const qty = Number(product.quantity || product.qty || 1) || 0;
      const price = Number(product.price || product.price_with_discount || product.payout || 0) || 0;
      const gross = qty * price;
      const sku = safe_(product.sku || product.product_id);
      const offerId = safe_(product.offer_id);
      const productId = safe_(product.product_id || product.id);
      const key = [postingDate || '', sku, offerId, productId].join('|');
      if (!map[key]) {
        map[key] = {
          date: postingDate || '',
          sku: sku,
          offer_id: offerId,
          product_id: productId,
          orders: 0,
          ordered_units: 0,
          gross_revenue: 0,
          net_revenue: 0,
        };
      }
      map[key].orders += 1;
      map[key].ordered_units += qty;
      map[key].gross_revenue += gross;
      map[key].net_revenue += gross;
    });
  });
  return map;
}

function createDailyTrigger() { deleteOzonTriggers(); ScriptApp.newTrigger('runYesterday').timeBased().everyDays(1).atHour(7).create(); writeDiagnostic('info', 'Trigger', 'Создан ежедневный триггер runYesterday.', '', '', 'Проверяйте выполнение в Diagnostics'); SpreadsheetApp.getUi().alert('Ежедневный автозапуск создан'); }
function deleteOzonTriggers() { ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'runYesterday') ScriptApp.deleteTrigger(t); }); writeDiagnostic('info', 'Trigger', 'Старые триггеры runYesterday удалены.', '', '', 'Готово'); }

function writeDiagnostic(severity, component, message, endpoint, tab, resolution) {
  ensureSheetWithHeader_('Diagnostics', SHEETS_SCHEMA.Diagnostics);
  SpreadsheetApp.getActive().getSheetByName('Diagnostics').appendRow([
    nowIso_(), severity, component, sanitizeError_(message), endpoint || '', tab || '', resolution || '',
  ]);
}

function ensureSheetWithHeader_(name, header) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const current = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  if (header.some(function (v, i) { return current[i] !== v; })) sheet.getRange(1, 1, 1, header.length).setValues([header]);
}

function clearDataRows_(sheet) { const r = sheet.getLastRow(); const c = sheet.getLastColumn(); if (r > 1 && c > 0) sheet.getRange(2, 1, r - 1, c).clearContent(); }
function extractItems_(responseJson) { return responseJson.result && responseJson.result.items ? responseJson.result.items : (Array.isArray(responseJson.result) ? responseJson.result : (responseJson.items || [])); }
function extractLastId_(responseJson) { return responseJson.result && responseJson.result.last_id ? responseJson.result.last_id : (responseJson.last_id || ''); }
function extractInfoItems_(responseJson) { return responseJson.result && responseJson.result.items ? responseJson.result.items : (Array.isArray(responseJson.result) ? responseJson.result : (responseJson.items || [])); }
function extractStockItems_(responseJson) {
  if (responseJson.result && responseJson.result.rows) return responseJson.result.rows;
  if (responseJson.result && responseJson.result.items) return responseJson.result.items;
  if (Array.isArray(responseJson.result)) return responseJson.result;
  if (responseJson.items) return responseJson.items;
  return [];
}
function normalizeProductListItem_(item) {
  const status = typeof item.status === 'object' ? (item.status.state || item.status.status || '') : safe_(item.status || item.visibility);
  return {
    sku: safe_(item.sku || item.product_id),
    offer_id: safe_(item.offer_id),
    product_id: safe_(item.product_id || item.id),
    name: safe_(item.name),
    category: safe_(item.category_name || item.category),
    brand: safe_(item.brand),
    status: safe_(status),
  };
}
function normalizeStockItem_(item, date, loadedAt) {
  const present = item.present && typeof item.present === 'object' ? (item.present.amount || item.present.total || '') : item.present;
  const reserved = item.reserved && typeof item.reserved === 'object' ? (item.reserved.amount || item.reserved.total || '') : item.reserved;
  const warehouse = item.warehouse || {};
  const availableStock = safe_(
    item.available_stock !== undefined ? item.available_stock :
      (item.free_to_sell_amount !== undefined ? item.free_to_sell_amount : present)
  );
  const reservedStock = safe_(
    item.reserved_stock !== undefined ? item.reserved_stock :
      (item.reserved_amount !== undefined ? item.reserved_amount : reserved)
  );
  const daysOfStock = safe_(
    item.days_of_stock !== undefined ? item.days_of_stock :
      (item.supply_days !== undefined ? item.supply_days : '')
  );
  const turnoverRate = safe_(
    item.turnover_rate !== undefined ? item.turnover_rate :
      (item.turnover_days !== undefined ? item.turnover_days : '')
  );
  const stockoutRisk = safe_(item.stockout_risk !== undefined ? item.stockout_risk : item.risk_level);

  return {
    date: date,
    sku: safe_(item.sku || item.product_id),
    offer_id: safe_(item.offer_id),
    warehouse_id: safe_(item.warehouse_id || warehouse.id),
    warehouse_name: safe_(item.warehouse_name || warehouse.name),
    available_stock: availableStock,
    reserved_stock: reservedStock,
    days_of_stock: daysOfStock,
    turnover_rate: turnoverRate,
    stockout_risk: stockoutRisk,
    loaded_at: loadedAt,
  };
}
function chunkArray_(arr, chunkSize) {
  const chunks = [];
  for (let i = 0; i < arr.length; i += chunkSize) chunks.push(arr.slice(i, i + chunkSize));
  return chunks;
}
function shiftDate_(date, deltaDays) {
  const d = new Date(date);
  d.setDate(d.getDate() + deltaDays);
  return d;
}
function readSheetObjects_(sheetName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return values.map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}
function aggregateDailyMetrics_(funnelRows, financeRows) {
  const map = {};
  funnelRows.forEach(function (r) {
    const d = safe_(r.date);
    if (!d) return;
    if (!map[d]) map[d] = { orders: 0, ordered_units: 0, gross_revenue: 0, net_revenue: 0 };
    map[d].orders += Number(r.orders || 0) || 0;
    map[d].ordered_units += Number(r.ordered_units || 0) || 0;
    map[d].gross_revenue += Number(r.gross_revenue || 0) || 0;
  });
  financeRows.forEach(function (r) {
    const d = safe_(r.date);
    if (!d) return;
    if (!map[d]) map[d] = { orders: 0, ordered_units: 0, gross_revenue: 0, net_revenue: 0 };
    map[d].net_revenue += Number(r.net_revenue || 0) || 0;
  });
  return map;
}
function rollingAverage_(values, windowSize) {
  if (!values.length || windowSize <= 0) return null;
  const slice = values.slice(-windowSize);
  const sum = slice.reduce(function (acc, v) { return acc + Number(v || 0); }, 0);
  return sum / slice.length;
}
function buildSkuSummary_(products, stocks, funnel, finance) {
  const allSkus = {};
  products.forEach(function (p) { if (p.sku) allSkus[p.sku] = { hasProduct: true }; });
  stocks.forEach(function (s) { if (s.sku) allSkus[s.sku] = allSkus[s.sku] || {}; });
  funnel.forEach(function (f) { if (f.sku) allSkus[f.sku] = allSkus[f.sku] || {}; });
  finance.forEach(function (f) { if (f.sku) allSkus[f.sku] = allSkus[f.sku] || {}; });

  Object.keys(allSkus).forEach(function (sku) {
    const stockRows = stocks.filter(function (s) { return safe_(s.sku) === sku; });
    const funnelRows = funnel.filter(function (f) { return safe_(f.sku) === sku; }).sort(function (a, b) { return safe_(a.date) > safe_(b.date) ? 1 : -1; });
    const financeRows = finance.filter(function (f) { return safe_(f.sku) === sku; }).sort(function (a, b) { return safe_(a.date) > safe_(b.date) ? 1 : -1; });
    const latestStock = stockRows.length ? stockRows[stockRows.length - 1] : {};
    const availableStock = Number(latestStock.available_stock || 0) || 0;
    const stockoutRisk = safe_(latestStock.stockout_risk || '').toLowerCase();
    const last14Funnel = funnelRows.slice(-14);
    const last7 = last14Funnel.slice(-7);
    const prev7 = last14Funnel.slice(0, Math.max(0, last14Funnel.length - 7));
    const orders7 = sumBy_(last7, 'orders');
    const ordersPrev7 = sumBy_(prev7, 'orders');
    const gross7 = sumBy_(funnelRows.slice(-7), 'gross_revenue');
    const grossPrev7 = sumBy_(funnelRows.slice(-14, -7), 'gross_revenue');
    const net7 = sumBy_(financeRows.slice(-7), 'net_revenue');
    const netPrev7 = sumBy_(financeRows.slice(-14, -7), 'net_revenue');

    allSkus[sku].hasProduct = !!products.find(function (p) { return safe_(p.sku) === sku; });
    allSkus[sku].availableStock = availableStock;
    allSkus[sku].stockoutRisk = stockoutRisk;
    allSkus[sku].stockoutRiskHigh = stockoutRisk === 'high';
    allSkus[sku].availableStockLow = availableStock <= 5;
    allSkus[sku].orders7 = orders7;
    allSkus[sku].ordersPrev7 = ordersPrev7;
    allSkus[sku].salesDrop = ordersPrev7 > 0 && orders7 < ordersPrev7 * 0.8;
    allSkus[sku].salesGrow = ordersPrev7 > 0 && orders7 > ordersPrev7 * 1.2;
    allSkus[sku].orders30 = sumBy_(funnelRows.slice(-30), 'orders');
    allSkus[sku].gross7 = gross7;
    allSkus[sku].net7 = net7;
    allSkus[sku].grossGrow = grossPrev7 > 0 && gross7 > grossPrev7 * 1.1;
    allSkus[sku].netDrop = netPrev7 > 0 && net7 < netPrev7 * 0.9;
  });
  return allSkus;
}
function sumBy_(rows, field) {
  return rows.reduce(function (acc, row) { return acc + (Number(row[field] || 0) || 0); }, 0);
}
function round2_(num) {
  return Math.round((Number(num || 0) + Number.EPSILON) * 100) / 100;
}
function toNum_(value) {
  const n = Number(value);
  return isNaN(n) ? 0 : n;
}
function summarizeAds_(adsRows) {
  const map = {};
  adsRows.forEach(function (row) {
    const id = safe_(row.campaign_id || row.id || 'unknown');
    if (!map[id]) map[id] = { campaignId: id, spend: 0, clicks: 0, impressions: 0, orders: 0, revenue: 0 };
    map[id].spend += toNum_(row.spend);
    map[id].clicks += toNum_(row.clicks);
    map[id].impressions += toNum_(row.impressions);
    map[id].orders += toNum_(row.orders);
    map[id].revenue += toNum_(row.revenue);
  });
  return Object.keys(map).map(function (id) {
    const x = map[id];
    const ctr = x.impressions > 0 ? x.clicks / x.impressions : 0;
    const cpc = x.clicks > 0 ? x.spend / x.clicks : 0;
    const roas = x.spend > 0 && x.revenue > 0 ? x.revenue / x.spend : 0;
    const drr = x.revenue > 0 ? x.spend / x.revenue : 0;
    return {
      campaignId: id,
      spend: round2_(x.spend),
      orders: round2_(x.orders),
      revenue: round2_(x.revenue),
      ctr: round2_(ctr),
      cpc: round2_(cpc),
      roas: roas ? round2_(roas) : 0,
      drr: drr ? round2_(drr) : 0,
      lowCtr: ctr > 0 && ctr < 0.01,
      cpcGrow: cpc > 0.3,
      spendNoReturn: x.spend > 0 && (x.orders <= 0 || x.revenue <= 0),
      lowRoasOrHighDrr: (roas > 0 && roas < 2) || (drr > 0.4),
    };
  });
}
function topBySku_(rows, field, limit) {
  const m = {};
  rows.forEach(function (r) {
    const sku = safe_(r.sku || '');
    if (!sku) return;
    m[sku] = (m[sku] || 0) + (Number(r[field] || 0) || 0);
  });
  return Object.keys(m)
    .map(function (sku) { return { sku: sku, value: round2_(m[sku]) }; })
    .sort(function (a, b) { return b.value - a.value; })
    .slice(0, limit || 5);
}
function highSpendLowReturnCampaigns_(adsRows, limit) {
  const byCampaign = {};
  adsRows.forEach(function (r) {
    const c = safe_(r.campaign_name || r.campaign_id || 'unknown');
    if (!byCampaign[c]) byCampaign[c] = { spend: 0, revenue: 0 };
    byCampaign[c].spend += parseNumber_(r.spend);
    byCampaign[c].revenue += parseNumber_(r.revenue);
  });
  return Object.keys(byCampaign).map(function (c) {
    const v = byCampaign[c];
    const roas = v.spend > 0 && v.revenue > 0 ? v.revenue / v.spend : null;
    return { campaign: c, spend: round2_(v.spend), roas: roas === null ? null : round2_(roas) };
  }).filter(function (x) { return x.spend > 0 && (x.roas === null || x.roas < 1 || (x.roas > 0 && 1 / x.roas > 0.5)); })
    .sort(function (a, b) { return b.spend - a.spend; })
    .slice(0, limit || 5);
}
function parseNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return isNaN(value) ? 0 : value;
  let s = String(value).trim();
  s = s.replace(/[₽\s]/g, '').replace(',', '.').replace(/[^0-9.\-]/g, '');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}
function formatDateForSummary_(dateValue) {
  if (!dateValue) return 'нет данных';
  const d = new Date(dateValue);
  if (isNaN(d.getTime())) return safe_(dateValue);
  return formatDate_(d, 'yyyy-MM-dd');
}
function sumByParsed_(rows, field) {
  return rows.reduce(function (acc, row) { return acc + parseNumber_(row[field]); }, 0);
}
function sumAnyField_(rows, fields) {
  return rows.reduce(function (acc, row) {
    for (let i = 0; i < fields.length; i++) {
      const v = parseNumber_(row[fields[i]]);
      if (v) return acc + v;
    }
    return acc;
  }, 0);
}
function buildProductsMap_() {
  const products = readSheetObjects_('Products');
  const map = {};
  products.forEach(function (p) {
    const pid = safe_(p.product_id);
    if (pid) map[pid] = { offer_id: safe_(p.offer_id), sku: safe_(p.sku), name: safe_(p.name) };
  });
  return map;
}
function chooseDimensionKey_(row, productsMap) {
  const pid = safe_(row.product_id);
  if (safe_(row.offer_id)) return normalizeOfferId_(row.offer_id);
  if (safe_(row.sku)) return normalizeOfferId_(row.sku);
  if (pid && productsMap[pid] && productsMap[pid].offer_id) return productsMap[pid].offer_id;
  if (pid && productsMap[pid] && productsMap[pid].sku) return productsMap[pid].sku;
  return pid || 'unknown';
}
function topByDimension_(rows, field, productsMap, limit) {
  const m = {};
  rows.forEach(function (r) {
    const key = chooseDimensionKey_(r, productsMap || {});
    m[key] = (m[key] || 0) + parseNumber_(r[field]);
  });
  return Object.keys(m).map(function (k) { return { key: k, value: round2_(m[k]) }; })
    .sort(function (a, b) { return b.value - a.value; })
    .slice(0, limit || 10);
}
function topAdsByDimension_(adsRows, productsMap, limit) {
  const m = {};
  adsRows.forEach(function (r) {
    let key = safe_(r.offer_id) || safe_(r.sku);
    if (!key) key = safe_(r.campaign_name) || safe_(r.campaign_id) || 'unknown_campaign';
    m[key] = (m[key] || 0) + parseNumber_(r.spend);
  });
  return Object.keys(m).map(function (k) { return { key: k, value: round2_(m[k]) }; })
    .sort(function (a, b) { return b.value - a.value; })
    .slice(0, limit || 10);
}
function lowStockTop_(stocksRows, productsRows, limit) {
  const pMap = {};
  (productsRows || []).forEach(function (p) { if (p.product_id) pMap[String(p.product_id)] = normalizeOfferId_(p.offer_id || p.sku || p.product_id); });
  const dedup = {};
  stocksRows.forEach(function (s) {
    const available = parseNumber_(s.available_stock);
    const days = parseNumber_(s.days_of_stock);
    const risk = String(s.stockout_risk || '').toLowerCase();
    const keyRaw = normalizeOfferId_(s.offer_id || s.sku || pMap[String(s.product_id)] || s.product_id || 'unknown');
    const wh = safe_(s.warehouse_id || '');
    const dedupKey = keyRaw + '|' + wh;
    dedup[dedupKey] = {
      key: keyRaw,
      available: available,
      days: days,
      riskHigh: risk === 'high',
    };
  });
  return Object.keys(dedup).map(function (k) { return dedup[k]; }).filter(function (x) { return x.riskHigh || x.days <= 7 || x.available <= 5; })
    .sort(function (a, b) {
      if (a.riskHigh !== b.riskHigh) return a.riskHigh ? -1 : 1;
      return a.available - b.available;
    }).slice(0, limit || 10);
}
function calcCogsFromFunnel_(funnelRows, productsRows, unitCostMap) {
  const pByPid = {};
  (productsRows || []).forEach(function (p) { if (p.product_id) pByPid[String(p.product_id)] = p; });
  let total = 0;
  const missing = {};
  funnelRows.forEach(function (row) {
    let offer = normalizeOfferId_(row.offer_id);
    if (!offer && row.product_id && pByPid[String(row.product_id)]) offer = normalizeOfferId_(pByPid[String(row.product_id)].offer_id || pByPid[String(row.product_id)].sku || row.product_id);
    if (!offer) offer = normalizeOfferId_(row.sku);
    const units = parseNumber_(row.ordered_units || row.units || row.orders || 0);
    const unitCost = unitCostMap[offer];
    if (!unitCost) {
      if (offer) missing[offer] = true;
      return;
    }
    total += units * unitCost;
  });
  return { total: total, missing: Object.keys(missing) };
}
function calcSkuPnl_(funnelRows, productsRows, unitCostMap, limit) {
  const pByPid = {};
  (productsRows || []).forEach(function (p) { if (p.product_id) pByPid[String(p.product_id)] = p; });
  const m = {};
  funnelRows.forEach(function (r) {
    let key = normalizeOfferId_(r.offer_id);
    if (!key && r.product_id && pByPid[String(r.product_id)]) key = normalizeOfferId_(pByPid[String(r.product_id)].offer_id || pByPid[String(r.product_id)].sku || r.product_id);
    if (!key) key = normalizeOfferId_(r.sku || r.product_id || 'unknown');
    if (!m[key]) m[key] = { key: key, revenue: 0, units: 0, cogs: 0 };
    const units = parseNumber_(r.ordered_units || r.units || r.orders || 0);
    const rev = parseNumber_(r.gross_revenue);
    m[key].units += units;
    m[key].revenue += rev;
    const uc = unitCostMap[key];
    if (uc) m[key].cogs += units * uc;
  });
  const arr = Object.keys(m).map(function (k) {
    const x = m[k];
    const profit = x.revenue - x.cogs;
    const margin = x.revenue > 0 ? profit / x.revenue : 0;
    return { key: k, revenue: x.revenue, units: x.units, cogs: x.cogs, profit: profit, margin: margin };
  });
  return {
    top: arr.sort(function (a, b) { return b.profit - a.profit; }).slice(0, limit || 20),
    loss: arr.filter(function (x) { return x.profit < 0; }).sort(function (a, b) { return a.profit - b.profit; }).slice(0, limit || 20),
  };
}
function groupedHighPriorityRecs_(recs, limit) {
  const high = recs.filter(function (r) { return String(r.priority || '').toLowerCase() === 'high'; });
  const groups = {};
  high.forEach(function (r) {
    const area = safe_(r.area || '');
    const issue = safe_(r.issue || '');
    const action = safe_(r.action || '');
    const key = [area, issue, action].join('|');
    if (!groups[key]) groups[key] = { area: area, issue: issue, action: action, count: 0, skuSet: {} };
    groups[key].count += 1;
    if (r.sku_or_campaign) groups[key].skuSet[String(r.sku_or_campaign)] = true;
  });
  return Object.keys(groups).map(function (k) {
    const g = groups[k];
    const skus = Object.keys(g.skuSet).slice(0, 10);
    const metric = g.area || 'unknown';
    const value = g.issue + ': ' + g.count + ' SKU';
    const notes = skus.length ? (skus.join(', ') + ' | ' + g.action) : g.action;
    return { metric: metric, value: value, notes: notes };
  }).slice(0, limit || 10);
}
function applySummaryFormats_(sheet, rowsCount) {
  if (!rowsCount) return;
  const values = sheet.getRange(2, 1, rowsCount, 4).getValues();
  for (let i = 0; i < rowsCount; i++) {
    const metric = String(values[i][1] || '');
    const cell = sheet.getRange(i + 2, 3);
    if (metric.indexOf('Выручка') >= 0 || metric.indexOf('Себестоимость') >= 0 || metric.indexOf('прибыль') >= 0 || metric.indexOf('Комиссия') >= 0 || metric.indexOf('Логистика') >= 0 || metric.indexOf('Возвраты') >= 0 || metric.indexOf('расходы') >= 0 || metric.indexOf('чек') >= 0 || metric.indexOf('profit:') >= 0) cell.setNumberFormat('#,##0.00 ₽');
    if (metric === 'CTR' || metric.indexOf('DRR') >= 0 || metric.indexOf('ROAS') >= 0 || metric.indexOf('маржа') >= 0 || metric.indexOf('Маржинальность') >= 0) cell.setNumberFormat('0.00%');
  }
}
function getSafeTimeZone_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss && typeof ss.getSpreadsheetTimeZone === 'function'
    ? ss.getSpreadsheetTimeZone()
    : '';
  if (tz && typeof tz === 'string') return tz;
  const scriptTz = Session.getScriptTimeZone();
  if (scriptTz && typeof scriptTz === 'string') return scriptTz;
  return 'Europe/Moscow';
}
function formatDate_(date, pattern) {
  return Utilities.formatDate(date, getSafeTimeZone_(), pattern || 'yyyy-MM-dd');
}
function nowIso_() {
  return formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
}
function sanitizeError_(err) { const text = typeof err === 'string' ? err : (err && err.message ? err.message : String(err)); return text.replace(/Api-Key\s*[:=]\s*[^,\s]+/gi, 'Api-Key=[redacted]').replace(/Client-Id\s*[:=]\s*[^,\s]+/gi, 'Client-Id=[redacted]').replace(/Bearer\s+[A-Za-z0-9._\-]+/gi, 'Bearer [redacted]'); }
function shorten_(text, maxLen) { return text.length > maxLen ? text.substring(0, maxLen) + '...' : text; }
function safe_(value) { return value === undefined || value === null ? '' : value; }
