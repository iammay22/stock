/**
 * 看懂台股 後端（Google Apps Script）
 * 作者：may22
 *
 * 做三件事：
 * 1. dailyFetch：每天收盤後抓證交所、櫃買中心的全市場資料，寫進試算表
 * 2. doGet：回答前端的查詢（home、list、stock），個股的歷史與財報向 FinMind 補
 * 3. doPost：收前端的使用紀錄，寫進 events 分頁
 *
 * 第一次使用：在編輯器上方選 setup 按執行，再部署成網頁應用程式。
 * 想確認各來源欄位有沒有改版：執行 testSources，看執行紀錄。
 */

var TZ = 'Asia/Taipei';
var TWSE = 'https://openapi.twse.com.tw/v1';
var TPEX = 'https://www.tpex.org.tw/openapi/v1';
var FINMIND = 'https://api.finmindtrade.com/api/v4/data';
var CODE_RE = /^(\d{4}|00\d{2,4}[A-Z]?)$/;

var SHEETS = {
  quotes: ['代號', '名稱', '市場', '收盤', '漲跌', '開盤', '最高', '最低', '成交股數', '成交金額', '資料日期'],
  valuation: ['代號', '本益比', '殖利率', '股價淨值比', '資料日期'],
  companies: ['代號', '簡稱', '全名', '市場', '產業', '董事長', '上市櫃日', '成立日', '實收資本額', '發行股數', '網址'],
  index: ['指數', '收盤', '漲跌點數', '漲跌百分比', '資料日期'],
  warnings: ['代號', '類型', '說明', '日期'],
  meta: ['來源', '資料日期', '抓取時間', '成功', '筆數', '訊息'],
  errors: ['時間', '來源', '訊息'],
  events: ['時間', 'session', '事件', '數值', '裝置', '來源']
};

var INDUSTRY = {
  '01': '水泥工業', '02': '食品工業', '03': '塑膠工業', '04': '紡織纖維', '05': '電機機械',
  '06': '電器電纜', '08': '玻璃陶瓷', '09': '造紙工業', '10': '鋼鐵工業', '11': '橡膠工業',
  '12': '汽車工業', '14': '建材營造', '15': '航運業', '16': '觀光餐旅', '17': '金融保險',
  '18': '貿易百貨', '19': '綜合', '20': '其他', '21': '化學工業', '22': '生技醫療',
  '23': '油電燃氣', '24': '半導體業', '25': '電腦及週邊設備', '26': '光電業', '27': '通信網路',
  '28': '電子零組件', '29': '電子通路', '30': '資訊服務', '31': '其他電子', '32': '文化創意',
  '33': '農業科技', '34': '電子商務', '35': '綠能環保', '36': '數位雲端', '37': '運動休閒',
  '38': '居家生活', '80': '管理股票', '91': '存託憑證'
};

/* ---------- 小工具 ---------- */

function pick_(o, keys) {
  for (var i = 0; i < keys.length; i++) {
    var v = o[keys[i]];
    if (v !== undefined && v !== null && String(v).trim() !== '') return v;
  }
  return '';
}

function num_(v) {
  if (v === null || v === undefined) return null;
  var s = String(v).replace(/,/g, '').replace(/\+/g, '').trim();
  if (s === '' || /^-+$/.test(s) || /^[^\d\-.]/.test(s)) return null;
  var n = Number(s);
  return isNaN(n) ? null : n;
}

/** 民國或西元日期字串轉成 YYYY-MM-DD，看不懂就回空字串 */
function isoDate_(v) {
  var s = String(v || '').trim();
  if (!s) return '';
  var y, m, d, t;
  if (/^\d{7}$/.test(s)) { y = +s.slice(0, 3); m = +s.slice(3, 5); d = +s.slice(5, 7); }
  else if (/^\d{8}$/.test(s)) { y = +s.slice(0, 4); m = +s.slice(4, 6); d = +s.slice(6, 8); }
  else {
    t = s.match(/(\d{2,4})\D+(\d{1,2})\D+(\d{1,2})/);
    if (!t) return '';
    y = +t[1]; m = +t[2]; d = +t[3];
  }
  if (y < 1911) y += 1911;
  if (!(m >= 1 && m <= 12 && d >= 1 && d <= 31)) return '';
  return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2);
}

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function yearsAgo_(n) {
  var d = new Date();
  d.setFullYear(d.getFullYear() - n);
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function writeSheet_(name, rows) {
  var sh = sheet_(name);
  var cols = SHEETS[name].length;
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, cols).clearContent();
  if (!rows.length) return;
  // 試算表不收 null，抓不到的欄位留空
  var clean = rows.map(function (r) {
    return r.map(function (x) { return x === null || x === undefined ? '' : x; });
  });
  // 全部存成純文字：不然 0050 會被試算表當成數字 50，日期也會被改格式
  var range = sh.getRange(2, 1, clean.length, cols);
  range.setNumberFormat('@');
  range.setValues(clean.map(function (r) { return r.map(function (x) { return String(x); }); }));
}

var memo_ = {};
function readSheet_(name) {
  if (memo_[name]) return memo_[name];
  var sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return (memo_[name] = []);
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, SHEETS[name].length).getDisplayValues();
  return (memo_[name] = rows);
}

function logError_(source, err) {
  try { sheet_('errors').appendRow([now_(), source, String(err && err.message || err).slice(0, 300)]); } catch (e) {}
}

function setMeta_(source, date, ok, rows, msg) {
  var sh = sheet_('meta');
  var data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
  var row = [source, date || '', now_(), ok ? '是' : '否', rows || 0, msg || ''];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === source) {
      // 抓失敗時保留上一次成功的資料日期，前端才能說「這是上次的資料」
      if (!ok) row[1] = sh.getRange(i + 2, 2).getDisplayValue();
      sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
      return;
    }
  }
  sh.appendRow(row);
}

function metaMap_() {
  var out = {};
  readSheet_('meta').forEach(function (r) {
    out[r[0]] = { date: r[1], at: r[2], ok: r[3] === '是', rows: +r[4] || 0, msg: r[5] };
  });
  return out;
}

function fetchJson_(url, headers) {
  var res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: headers || { Accept: 'application/json' }
  });
  var code = res.getResponseCode();
  if (code !== 200) throw new Error('HTTP ' + code);
  var data = JSON.parse(res.getContentText());
  return data;
}

/* CacheService 單筆上限 100KB，中文一個字佔 3 bytes，所以切小塊 */
function putBig_(key, str, ttl) {
  var cache = CacheService.getScriptCache();
  var size = 30000, n = Math.ceil(str.length / size), all = {};
  for (var i = 0; i < n; i++) all[key + '_' + i] = str.substr(i * size, size);
  all[key + '_n'] = String(n);
  try { cache.putAll(all, ttl); } catch (e) {}
}

function getBig_(key) {
  var cache = CacheService.getScriptCache();
  var n = +cache.get(key + '_n');
  if (!n) return null;
  var keys = [];
  for (var i = 0; i < n; i++) keys.push(key + '_' + i);
  var got = cache.getAll(keys), out = '';
  for (var j = 0; j < n; j++) {
    if (got[keys[j]] === undefined) return null;
    out += got[keys[j]];
  }
  return out;
}

function median_(arr) {
  var a = arr.slice().sort(function (x, y) { return x - y; });
  if (!a.length) return null;
  var m = Math.floor(a.length / 2);
  return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
}

function quantile_(arr, q) {
  var a = arr.slice().sort(function (x, y) { return x - y; });
  if (!a.length) return null;
  var pos = (a.length - 1) * q, lo = Math.floor(pos), hi = Math.ceil(pos);
  return a[lo] + (a[hi] - a[lo]) * (pos - lo);
}

function round_(n, d) {
  if (n === null || n === undefined || isNaN(n)) return null;
  var p = Math.pow(10, d === undefined ? 2 : d);
  return Math.round(n * p) / p;
}

/* ---------- 每日抓取 ---------- */

/** 每個來源各自 try，壞一個不影響其他 */
function source_(name, fn) {
  try {
    var r = fn();
    setMeta_(name, r.date, true, r.rows, '');
  } catch (e) {
    logError_(name, e);
    setMeta_(name, '', false, 0, String(e.message || e).slice(0, 120));
  }
}

function dailyFetch() {
  memo_ = {};
  var twseQ = [], tpexQ = [], twseV = [], tpexV = [], twseC = [], tpexC = [];
  var keepQ = readSheet_('quotes'), keepV = readSheet_('valuation'), keepC = readSheet_('companies');

  source_('證交所行情', function () {
    var date = '';
    fetchJson_(TWSE + '/exchangeReport/STOCK_DAY_ALL').forEach(function (r) {
      var id = String(pick_(r, ['Code'])).trim();
      if (!CODE_RE.test(id)) return;
      date = isoDate_(pick_(r, ['Date'])) || date;
      twseQ.push([id, String(pick_(r, ['Name'])).trim(), '上市', num_(r.ClosingPrice), num_(r.Change),
        num_(r.OpeningPrice), num_(r.HighestPrice), num_(r.LowestPrice), num_(r.TradeVolume), num_(r.TradeValue), date]);
    });
    if (!twseQ.length) throw new Error('沒有資料');
    return { date: date, rows: twseQ.length };
  });

  source_('櫃買行情', function () {
    var date = '';
    fetchJson_(TPEX + '/tpex_mainboard_daily_close_quotes').forEach(function (r) {
      var id = String(pick_(r, ['SecuritiesCompanyCode', 'Code'])).trim();
      if (!CODE_RE.test(id)) return;
      date = isoDate_(pick_(r, ['Date'])) || date;
      tpexQ.push([id, String(pick_(r, ['CompanyName', 'Name'])).trim(), '上櫃', num_(r.Close), num_(r.Change),
        num_(r.Open), num_(r.High), num_(r.Low),
        num_(pick_(r, ['TradingShares', 'TradeVolume'])), num_(pick_(r, ['TransactionAmount', 'TradeValue'])), date]);
    });
    if (!tpexQ.length) throw new Error('沒有資料');
    return { date: date, rows: tpexQ.length };
  });

  source_('證交所估值', function () {
    var date = '';
    fetchJson_(TWSE + '/exchangeReport/BWIBBU_ALL').forEach(function (r) {
      var id = String(r.Code || '').trim();
      if (!CODE_RE.test(id)) return;
      date = isoDate_(r.Date) || date;
      twseV.push([id, num_(r.PEratio), num_(r.DividendYield), num_(r.PBratio), date]);
    });
    if (!twseV.length) throw new Error('沒有資料');
    return { date: date, rows: twseV.length };
  });

  source_('櫃買估值', function () {
    var date = '';
    fetchJson_(TPEX + '/tpex_mainboard_peratio_analysis').forEach(function (r) {
      var id = String(pick_(r, ['SecuritiesCompanyCode', 'Code'])).trim();
      if (!CODE_RE.test(id)) return;
      date = isoDate_(r.Date) || date;
      tpexV.push([id, num_(pick_(r, ['PriceEarningRatio', 'PEratio'])), num_(pick_(r, ['YieldRatio', 'DividendYield'])),
        num_(pick_(r, ['PriceBookRatio', 'PBratio'])), date]);
    });
    if (!tpexV.length) throw new Error('沒有資料');
    return { date: date, rows: tpexV.length };
  });

  source_('上市公司基本資料', function () {
    fetchJson_(TWSE + '/opendata/t187ap03_L').forEach(function (r) {
      var id = String(r['公司代號'] || '').trim();
      if (!id) return;
      var code = String(r['產業別'] || '').trim();
      twseC.push([id, r['公司簡稱'] || '', r['公司名稱'] || '', '上市', INDUSTRY[code] || code || '',
        r['董事長'] || '', isoDate_(r['上市日期']), isoDate_(r['成立日期']), num_(r['實收資本額']),
        num_(r['已發行普通股數或TDR原股發行股數']), r['網址'] || '']);
    });
    if (!twseC.length) throw new Error('沒有資料');
    return { date: today_(), rows: twseC.length };
  });

  source_('上櫃公司基本資料', function () {
    fetchJson_(TPEX + '/mopsfin_t187ap03_O').forEach(function (r) {
      var id = String(pick_(r, ['SecuritiesCompanyCode', '公司代號'])).trim();
      if (!id) return;
      var code = String(pick_(r, ['SecuritiesIndustryCode', '產業別'])).trim();
      tpexC.push([id, pick_(r, ['CompanyAbbreviation', '公司簡稱']), pick_(r, ['CompanyName', '公司名稱']), '上櫃',
        INDUSTRY[code] || code || '', pick_(r, ['Chairman', '董事長']),
        isoDate_(pick_(r, ['DateOfListing', '上櫃日期'])), isoDate_(pick_(r, ['DateOfIncorporation', '成立日期'])),
        num_(pick_(r, ['Paidin.Capital.NTDollars', 'PaidinCapitalNTDollars', '實收資本額'])),
        num_(pick_(r, ['IssueShares', 'IssuedShares', '已發行普通股數或TDR原股發行股數'])),
        pick_(r, ['WebAddress', '網址'])]);
    });
    if (!tpexC.length) throw new Error('沒有資料');
    return { date: today_(), rows: tpexC.length };
  });

  // 某一邊抓失敗時，保留試算表裡那一邊的舊資料
  function merge(fresh1, fresh2, keep, marketCol) {
    var a = fresh1.length ? fresh1 : keep.filter(function (r) { return marketCol < 0 ? false : r[marketCol] === '上市'; });
    var b = fresh2.length ? fresh2 : keep.filter(function (r) { return marketCol < 0 ? false : r[marketCol] === '上櫃'; });
    return a.concat(b);
  }
  if (twseQ.length || tpexQ.length) writeSheet_('quotes', merge(twseQ, tpexQ, keepQ, 2));
  if (twseC.length || tpexC.length) writeSheet_('companies', merge(twseC, tpexC, keepC, 3));
  if (twseV.length || tpexV.length) {
    var mk = {};
    (twseQ.length || tpexQ.length ? twseQ.concat(tpexQ) : keepQ).forEach(function (r) { mk[r[0]] = r[2]; });
    var oldV = keepV.map(function (r) { return r; });
    var a = twseV.length ? twseV : oldV.filter(function (r) { return mk[r[0]] === '上市'; });
    var b = tpexV.length ? tpexV : oldV.filter(function (r) { return mk[r[0]] === '上櫃'; });
    writeSheet_('valuation', a.concat(b));
  }

  source_('大盤指數', function () {
    var rows = [], date = '';
    fetchJson_(TWSE + '/exchangeReport/MI_INDEX').forEach(function (r) {
      var name = String(r['指數'] || '').trim();
      if (!name) return;
      var sign = String(r['漲跌'] || '').indexOf('-') >= 0 ? -1 : 1;
      var pts = num_(r['漲跌點數']), pct = num_(r['漲跌百分比']);
      if (pts !== null) pts = Math.abs(pts) * sign;
      if (pct !== null) pct = Math.abs(pct) * sign;
      date = isoDate_(r['日期']) || date;
      rows.push([name, num_(r['收盤指數']), pts, pct, date]);
    });
    if (!rows.length) throw new Error('沒有資料');
    writeSheet_('index', rows);
    return { date: date, rows: rows.length };
  });

  fetchWarnings_();
  memo_ = {};
  // 資料換新了，把首頁與清單的快取版本往前推
  PropertiesService.getScriptProperties().setProperty('VER', String(Date.now()));
}

function fetchWarnings_() {
  var rows = [];
  var okTwse = true, okTpex = true, msgTwse = [], msgTpex = [];

  function grab(label, url, isTpex, mapper) {
    try {
      fetchJson_(url).forEach(function (r) {
        var out = mapper(r);
        if (out && out[0]) rows.push(out);
      });
    } catch (e) {
      logError_(label, e);
      if (isTpex) { okTpex = false; msgTpex.push(label); } else { okTwse = false; msgTwse.push(label); }
    }
  }

  grab('注意股', TWSE + '/announcement/notice', false, function (r) {
    return [String(r.Code || '').trim(), '注意股', String(r.TradingInfoForAttention || '').slice(0, 200), isoDate_(r.Date)];
  });
  grab('處置股', TWSE + '/announcement/punish', false, function (r) {
    return [String(r.Code || '').trim(), '處置股',
      ('處置期間 ' + (r.DispositionPeriod || '') + '。' + (r.DispositionMeasures || '')).slice(0, 200), isoDate_(r.Date)];
  });
  grab('變更交易', TWSE + '/exchangeReport/TWT85U', false, function (r) {
    return [String(r.Code || '').trim(), '變更交易', '', ''];
  });
  grab('暫停交易', TWSE + '/exchangeReport/TWTAWU', false, function (r) {
    return [String(r.Code || '').trim(), '暫停交易',
      '暫停日 ' + (r.TradingHaltDate || '') + (r.TradingResumptionDate ? '，恢復日 ' + r.TradingResumptionDate : ''), isoDate_(r.TradingHaltDate)];
  });
  grab('除權息預告', TWSE + '/exchangeReport/TWT48U_ALL', false, function (r) {
    var cash = num_(r.CashDividend), stock = num_(r.StockDividendRatio);
    var text = String(r.Exdividend || '除權息') + (cash ? '，每股現金 ' + cash + ' 元' : '') + (stock ? '，配股率 ' + stock : '');
    return [String(r.Code || '').trim(), '即將除權息', text, isoDate_(r.Date)];
  });

  grab('上櫃注意股', TPEX + '/tpex_trading_warning_information', true, function (r) {
    return [String(pick_(r, ['SecuritiesCompanyCode', 'Code'])).trim(), '注意股',
      String(pick_(r, ['TradingInformation', 'TradingInfoForAttention', 'Information'])).slice(0, 200), isoDate_(r.Date)];
  });
  grab('上櫃處置股', TPEX + '/tpex_disposal_information', true, function (r) {
    return [String(pick_(r, ['SecuritiesCompanyCode', 'Code'])).trim(), '處置股',
      ('處置期間 ' + pick_(r, ['DispositionPeriod', 'DisposalPeriod']) + '。' +
        pick_(r, ['DisposalCondition', 'DispositionMeasures', 'DispositionReasons'])).slice(0, 200), isoDate_(r.Date)];
  });
  grab('上櫃除權息預告', TPEX + '/tpex_exright_prepost', true, function (r) {
    return [String(pick_(r, ['SecuritiesCompanyCode', 'Code'])).trim(), '即將除權息',
      String(pick_(r, ['ExRrightsExDividend', 'ExRightsExDividend', 'Exdividend']) || '除權息') +
        (num_(r.CashDividend) ? '，每股現金 ' + num_(r.CashDividend) + ' 元' : ''),
      isoDate_(pick_(r, ['ExRrightsExDividendDate', 'ExRightsExDividendDate', 'ExDividendDate', 'Date']))];
  });

  writeSheet_('warnings', rows);
  setMeta_('上市警示', today_(), okTwse, rows.length, okTwse ? '' : '沒抓到：' + msgTwse.join('、'));
  setMeta_('上櫃警示', today_(), okTpex, rows.length, okTpex ? '' : '沒抓到：' + msgTpex.join('、'));
}

/* ---------- 回答前端 ---------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ver_() { return PropertiesService.getScriptProperties().getProperty('VER') || '0'; }

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'home') return json_(cached_('home_' + ver_(), 21600, home_));
    if (p.action === 'list') return json_(cached_('list_' + ver_(), 21600, list_));
    if (p.action === 'stock') return json_(stock_(String(p.id || '').toUpperCase().trim()));
    return ContentService.createTextOutput('alive');
  } catch (err) {
    logError_('doGet ' + (p.action || ''), err);
    return json_({ ok: false, error: 'server', msg: String(err.message || err).slice(0, 120) });
  }
}

function cached_(key, ttl, fn) {
  var hit = getBig_(key);
  if (hit) return JSON.parse(hit);
  var out = fn();
  if (out.ok) putBig_(key, JSON.stringify(out), ttl);
  return out;
}

function home_() {
  var idx = readSheet_('index');
  if (!readSheet_('quotes').length) return { ok: false, error: 'nosetup' };
  var main = null, sectors = [];
  idx.forEach(function (r) {
    if (r[0] === '發行量加權股價指數') main = { name: '加權指數', close: num_(r[1]), chg: num_(r[2]), pct: num_(r[3]), date: r[4] };
    else if (/類指數$/.test(r[0]) && num_(r[3]) !== null) sectors.push([r[0].replace(/類指數$/, ''), num_(r[3])]);
  });
  sectors.sort(function (a, b) { return b[1] - a[1]; });
  return { ok: true, idx: main, sectors: sectors, src: metaMap_() };
}

function list_() {
  var comp = {};
  readSheet_('companies').forEach(function (r) { comp[r[0]] = r; });
  var q = readSheet_('quotes');
  if (!q.length) return { ok: false, error: 'nosetup' };
  return {
    ok: true, ver: ver_(),
    list: q.map(function (r) {
      var c = comp[r[0]];
      return [r[0], (c && c[1]) || r[1], r[2], c ? c[4] : (/^00/.test(r[0]) ? 'ETF' : '')];
    })
  };
}

function stats_(values, trim) {
  var a = values.filter(function (v) { return v !== null && v > 0; });
  if (a.length < 3) return null;
  return {
    lo: round_(trim ? quantile_(a, 0.05) : Math.min.apply(null, a)),
    med: round_(median_(a)),
    hi: round_(trim ? quantile_(a, 0.95) : Math.max.apply(null, a)),
    n: a.length
  };
}

function finmindAll_(id) {
  var token = PropertiesService.getScriptProperties().getProperty('FINMIND_TOKEN');
  var headers = token ? { Authorization: 'Bearer ' + token } : {};
  var sets = [
    ['price', 'TaiwanStockPrice', yearsAgo_(5)],
    ['per', 'TaiwanStockPER', yearsAgo_(5)],
    ['rev', 'TaiwanStockMonthRevenue', yearsAgo_(4)],
    ['fin', 'TaiwanStockFinancialStatements', yearsAgo_(4)],
    ['bs', 'TaiwanStockBalanceSheet', yearsAgo_(2)],
    ['cf', 'TaiwanStockCashFlowsStatement', yearsAgo_(3)],
    ['div', 'TaiwanStockDividend', yearsAgo_(12)]
  ];
  var reqs = sets.map(function (s) {
    return {
      url: FINMIND + '?dataset=' + s[1] + '&data_id=' + encodeURIComponent(id) + '&start_date=' + s[2],
      headers: headers, muteHttpExceptions: true
    };
  });
  var out = {}, src = {};
  var res;
  // FinMind 會擋 Google 共用的連線位址（回 403）。被擋過就先停 6 小時不再打，
  // 這段時間由使用者的瀏覽器直接向 FinMind 要資料（見 index.html 的 fillFromBrowser）。
  var cache = CacheService.getScriptCache();
  if (cache.get('fm_block')) {
    sets.forEach(function (s) { out[s[0]] = []; src[s[0]] = { ok: false, msg: 'FinMind 擋掉了後端的連線' }; });
    return { data: out, src: src, blocked: true };
  }
  var blocked = false;
  try { res = UrlFetchApp.fetchAll(reqs); } catch (e) { res = []; logError_('FinMind', e); }
  sets.forEach(function (s, i) {
    var key = s[0];
    out[key] = [];
    try {
      if (!res[i]) throw new Error('連不上 FinMind');
      var code = res[i].getResponseCode();
      if (code === 402) throw new Error('今天這個小時的查詢額度用完了');
      if (code === 403) { blocked = true; throw new Error('FinMind 擋掉了後端的連線'); }
      if (code !== 200) throw new Error('HTTP ' + code);
      var body = JSON.parse(res[i].getContentText());
      out[key] = body.data || [];
      src[key] = { ok: true, msg: '' };
    } catch (e) {
      src[key] = { ok: false, msg: String(e.message || e).slice(0, 80) };
      if (!blocked) logError_('FinMind ' + s[1] + ' ' + id, e);
    }
  });
  if (blocked) { cache.put('fm_block', '1', 21600); logError_('FinMind', 'HTTP 403，改由瀏覽器抓'); }
  return { data: out, src: src, blocked: blocked };
}

function byType_(rows, wanted) {
  // 回傳 { date: { 類型: 值 } }，wanted 是 { 輸出鍵: [可能的 type 名稱] }
  var byDate = {};
  rows.forEach(function (r) {
    for (var k in wanted) {
      if (wanted[k].indexOf(r.type) >= 0) {
        byDate[r.date] = byDate[r.date] || {};
        if (byDate[r.date][k] === undefined) byDate[r.date][k] = Number(r.value);
      }
    }
  });
  return byDate;
}

function quarterLabel_(date) {
  return date.slice(0, 4) + 'Q' + Math.ceil(+date.slice(5, 7) / 3);
}

function stock_(id) {
  if (!/^[0-9A-Z]{4,7}$/.test(id)) return { ok: false, error: 'badid' };
  var key = 'stock_' + ver_() + '_' + id;
  var hit = getBig_(key);
  if (hit) return JSON.parse(hit);

  var quotes = readSheet_('quotes');
  if (!quotes.length) return { ok: false, error: 'nosetup' };
  var q = null;
  for (var i = 0; i < quotes.length; i++) if (quotes[i][0] === id) { q = quotes[i]; break; }
  if (!q) return { ok: false, error: 'notfound' };

  var comps = readSheet_('companies'), vals = readSheet_('valuation');
  var c = null, compMap = {}, valMap = {};
  comps.forEach(function (r) { compMap[r[0]] = r; if (r[0] === id) c = r; });
  vals.forEach(function (r) { valMap[r[0]] = r; });
  var isEtf = /^00/.test(id);
  var industry = c ? c[4] : (isEtf ? 'ETF' : '');
  var meta = metaMap_();

  // 市值排名：發行股數乘收盤價
  var mcap = null;
  if (c && num_(c[9]) && num_(q[3])) {
    var mine = num_(c[9]) * num_(q[3]), rank = 1, total = 0;
    quotes.forEach(function (r) {
      var cc = compMap[r[0]];
      if (!cc) return;
      var v = num_(cc[9]) * num_(r[3]);
      if (!v) return;
      total++;
      if (v > mine) rank++;
    });
    mcap = { value: mine, rank: rank, total: total };
  }

  // 同產業估值分布
  var ind = { pe: null, pb: null, dy: null };
  if (industry && !isEtf) {
    var pe = [], pb = [], dy = [];
    comps.forEach(function (r) {
      if (r[4] !== industry) return;
      var v = valMap[r[0]];
      if (!v) return;
      pe.push(num_(v[1])); dy.push(num_(v[2])); pb.push(num_(v[3]));
    });
    ind = { pe: stats_(pe, true), pb: stats_(pb, true), dy: stats_(dy, true) };
  }

  var fm = finmindAll_(id), d = fm.data;

  var hist = d.price.filter(function (r) { return Number(r.close) > 0; }).map(function (r) {
    return [r.date, +r.open, +r.max, +r.min, +r.close, +r.Trading_Volume];
  });

  var own = {
    pe: stats_(d.per.map(function (r) { return Number(r.PER); })),
    pb: stats_(d.per.map(function (r) { return Number(r.PBR); })),
    dy: stats_(d.per.map(function (r) { return Number(r.dividend_yield); }))
  };
  var v = valMap[id];
  var lastPer = d.per.length ? d.per[d.per.length - 1] : null;
  var val = {
    date: v ? v[4] : (lastPer ? lastPer.date : ''),
    pe: v ? num_(v[1]) : (lastPer ? Number(lastPer.PER) || null : null),
    dy: v ? num_(v[2]) : (lastPer ? Number(lastPer.dividend_yield) || null : null),
    pb: v ? num_(v[3]) : (lastPer ? Number(lastPer.PBR) || null : null),
    own: own, ind: ind
  };

  // 月營收與年增率
  var revMap = {};
  d.rev.forEach(function (r) {
    revMap[r.revenue_year + '-' + ('0' + r.revenue_month).slice(-2)] = Number(r.revenue);
  });
  var rev = Object.keys(revMap).sort().map(function (ym) {
    var prev = revMap[(+ym.slice(0, 4) - 1) + ym.slice(4)];
    return [ym, revMap[ym], prev ? round_((revMap[ym] / prev - 1) * 100, 1) : null];
  }).slice(-24);

  // 單季損益
  var finBy = byType_(d.fin, {
    rev: ['Revenue'], gp: ['GrossProfit'], oi: ['OperatingIncome'],
    ni: ['IncomeAfterTaxes'], eps: ['EPS']
  });
  var fin = Object.keys(finBy).sort().slice(-8).map(function (dt) {
    var f = finBy[dt];
    return {
      q: quarterLabel_(dt), date: dt,
      rev: f.rev === undefined ? null : f.rev, gp: f.gp === undefined ? null : f.gp,
      oi: f.oi === undefined ? null : f.oi, ni: f.ni === undefined ? null : f.ni,
      eps: f.eps === undefined ? null : f.eps
    };
  });

  // 資產負債（最新一季）
  var bsBy = byType_(d.bs, { assets: ['TotalAssets'], liab: ['Liabilities', 'TotalLiabilities'], equity: ['Equity', 'TotalEquity'] });
  var bsDates = Object.keys(bsBy).sort(), bs = null;
  if (bsDates.length) {
    var b = bsBy[bsDates[bsDates.length - 1]];
    bs = { date: bsDates[bsDates.length - 1], assets: b.assets || null, liab: b.liab || null, equity: b.equity || null };
  }

  // 現金流是年初至今累計，換算成最近四季
  var cfBy = byType_(d.cf, {
    ocf: ['CashFlowsFromOperatingActivities', 'NetCashInflowFromOperatingActivities'],
    capex: ['PropertyAndPlantAndEquipment']
  });
  var cfDates = Object.keys(cfBy).sort(), cf = null;
  if (cfDates.length) {
    var last = cfDates[cfDates.length - 1];
    var prevYearEnd = (+last.slice(0, 4) - 1) + '-12-31';
    var prevSame = (+last.slice(0, 4) - 1) + last.slice(4);
    var ttm = function (k) {
      var cur = cfBy[last][k];
      if (cur === undefined) return null;
      if (last.slice(5) === '12-31') return cur;
      if (!cfBy[prevYearEnd] || !cfBy[prevSame] || cfBy[prevYearEnd][k] === undefined || cfBy[prevSame][k] === undefined) return null;
      return cur + cfBy[prevYearEnd][k] - cfBy[prevSame][k];
    };
    cf = { date: last, ocf: ttm('ocf'), capex: ttm('capex') };
    if (cf.capex !== null) cf.capex = Math.abs(cf.capex);
  }

  // 股利：依除息日所在年度加總
  var divBy = {};
  d.div.forEach(function (r) {
    var cash = (Number(r.CashEarningsDistribution) || 0) + (Number(r.CashStatutorySurplus) || 0);
    var stock = (Number(r.StockEarningsDistribution) || 0) + (Number(r.StockStatutorySurplus) || 0);
    var ex = r.CashExDividendTradingDate || r.StockExDividendTradingDate || '';
    var y = (ex || r.date || '').slice(0, 4);
    if (!y || (!cash && !stock)) return;
    divBy[y] = divBy[y] || { y: y, cash: 0, stock: 0, ex: [] };
    divBy[y].cash = round_(divBy[y].cash + cash, 4);
    divBy[y].stock = round_(divBy[y].stock + stock, 4);
    if (r.CashExDividendTradingDate && cash) divBy[y].ex.push([r.CashExDividendTradingDate, round_(cash, 4)]);
  });
  var div = Object.keys(divBy).sort().map(function (y) { return divBy[y]; }).slice(-11);

  var warn = readSheet_('warnings').filter(function (r) { return r[0] === id; }).map(function (r) {
    return { type: r[1], text: r[2], date: r[3] };
  });
  var wm = meta[q[2] === '上市' ? '上市警示' : '上櫃警示'];

  var idx = null;
  readSheet_('index').forEach(function (r) {
    if (r[0] === '發行量加權股價指數') idx = { name: '加權指數', close: num_(r[1]), pct: num_(r[3]), date: r[4] };
  });

  var out = {
    ok: true, id: id, name: (c && c[1]) || q[1], market: q[2], industry: industry, isEtf: isEtf,
    co: c ? { full: c[2], chairman: c[5], listed: c[6], founded: c[7], capital: num_(c[8]), shares: num_(c[9]), web: c[10] } : null,
    mcap: mcap,
    day: { date: q[10], close: num_(q[3]), chg: num_(q[4]), open: num_(q[5]), high: num_(q[6]), low: num_(q[7]), vol: num_(q[8]), value: num_(q[9]) },
    hist: hist, val: val, rev: rev, fin: fin, bs: bs, cf: cf, div: div,
    warn: warn, warnOk: wm ? wm.ok : false, warnMsg: wm ? wm.msg : '警示資料還沒抓過',
    idx: idx, src: fm.src, fetched: now_()
  };

  var allOk = Object.keys(fm.src).every(function (k) { return fm.src[k].ok; });
  putBig_(key, JSON.stringify(out), allOk || fm.blocked ? 21600 : 600);
  return out;
}

/* ---------- 使用紀錄 ---------- */

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    sheet_('events').appendRow([new Date(), String(d.session || '').slice(0, 40), String(d.event || '').slice(0, 40),
      String(d.value || '').slice(0, 80), String(d.device || '').slice(0, 20), String(d.source || '').slice(0, 40)]);
  } catch (err) {}
  return ContentService.createTextOutput('ok');
}

/* ---------- 安裝與自我檢查 ---------- */

/** 第一次使用時執行一次：建分頁、抓第一次資料、設定每天自動抓 */
function setup() {
  Object.keys(SHEETS).forEach(function (name) { sheet_(name); });
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyFetch') ScriptApp.deleteTrigger(t);
  });
  // 兩個時間點：17 點先抓一次，20 點再補一次，避免來源還沒更新
  ScriptApp.newTrigger('dailyFetch').timeBased().atHour(17).everyDays(1).inTimezone(TZ).create();
  ScriptApp.newTrigger('dailyFetch').timeBased().atHour(20).everyDays(1).inTimezone(TZ).create();
  dailyFetch();
}

/** 把每個來源第一筆資料的欄位名稱印出來，來源改版時用這個對照 */
function testSources() {
  var urls = [
    TWSE + '/exchangeReport/STOCK_DAY_ALL', TWSE + '/exchangeReport/BWIBBU_ALL', TWSE + '/exchangeReport/MI_INDEX',
    TWSE + '/opendata/t187ap03_L', TWSE + '/announcement/notice', TWSE + '/announcement/punish',
    TWSE + '/exchangeReport/TWT85U', TWSE + '/exchangeReport/TWTAWU', TWSE + '/exchangeReport/TWT48U_ALL',
    TPEX + '/tpex_mainboard_daily_close_quotes', TPEX + '/tpex_mainboard_peratio_analysis', TPEX + '/mopsfin_t187ap03_O',
    TPEX + '/tpex_trading_warning_information', TPEX + '/tpex_disposal_information', TPEX + '/tpex_exright_prepost',
    FINMIND + '?dataset=TaiwanStockPrice&data_id=2330&start_date=' + yearsAgo_(0).slice(0, 8) + '01',
    FINMIND + '?dataset=TaiwanStockBalanceSheet&data_id=2330&start_date=' + yearsAgo_(1)
  ];
  urls.forEach(function (u) {
    try {
      var res = UrlFetchApp.fetch(u, { muteHttpExceptions: true });
      var code = res.getResponseCode();
      if (code !== 200) { Logger.log('✗ ' + u + ' → HTTP ' + code); return; }
      var body = JSON.parse(res.getContentText());
      var rows = Array.isArray(body) ? body : (body.data || []);
      var types = '';
      if (rows.length && rows[0].type !== undefined) {
        var seen = {};
        rows.forEach(function (r) { seen[r.type] = 1; });
        types = '\n   type 有：' + Object.keys(seen).slice(0, 60).join(', ');
      }
      Logger.log('✓ ' + u + ' → ' + rows.length + ' 筆\n   欄位：' + (rows[0] ? Object.keys(rows[0]).join(', ') : '沒有資料') + types);
    } catch (e) {
      Logger.log('✗ ' + u + ' → ' + e.message);
    }
  });
}
