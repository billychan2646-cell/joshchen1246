/**
 * 金煎所 GOLDEN CRISP — 訂單系統（Google Apps Script）
 *
 * 使用方式請看 docs/google-sheets-setup.md
 *  1. 在 Google 試算表開啟「擴充功能 → Apps Script」，貼上本檔
 *  2. 執行一次 setup()（建立分頁、商品、設定）
 *  3. 執行一次 installTriggers()（狀態通知、每日摘要）
 *  4. 部署 → 新增部署作業 → 網頁應用程式（執行身分：我；存取權：所有人）
 */

const CONFIG = {
  TZ: 'Asia/Taipei',
  SHEET: { ORDERS: '訂單', ITEMS: '訂單明細', MENU: '商品', SETTINGS: '設定' },
  STATUS: ['新訂單', '製作中', '可取餐', '已完成', '已取消'],
  MAX_LINES: 30,       // 單筆訂單最多幾種商品
  MAX_QTY: 50,         // 單一商品最大數量
  RATE_LIMIT_SEC: 20,  // 同一支電話的最短下單間隔
  DEDUPE_SEC: 21600    // 重送同一份訂單的辨識時間（6 小時）
};

const ORDER_HEADERS = [
  '訂單編號', '建立時間', '狀態', '取餐方式', '姓名', '電話', 'Email',
  '宅配地址', '備註', '品項摘要', '件數', '商品小計', '運費', '總金額', '通知紀錄', '來源'
];
const ITEM_HEADERS = ['訂單編號', '商品ID', '商品名稱', '單價', '數量', '小計'];
const MENU_HEADERS = ['商品ID', '分類', '商品名稱', '價格', '上架'];

// 欄位位置（1 起算）
const COL = { ORDER_NO: 1, CREATED: 2, STATUS: 3, METHOD: 4, NAME: 5, PHONE: 6, EMAIL: 7, NOTIFY: 15 };

// 與網站 index.html 的 products 對應（ID 必須一致）
const MENU_SEED = [
  ['p1', '招牌鍋貼', '經典鮮肉鍋貼', 85, true],
  ['p2', '招牌鍋貼', '韭香鮮肉鍋貼', 90, true],
  ['p3', '招牌鍋貼', '香辣椒麻鍋貼', 95, true],
  ['p4', '招牌鍋貼', '玉米鮮肉鍋貼', 90, true],
  ['d1', '精選水餃', '手工高麗菜水餃', 80, true],
  ['d2', '精選水餃', '韭菜鮮肉水餃', 85, true],
  ['d3', '精選水餃', '鮮蝦水餃', 110, true],
  ['s1', '湯品／飲料', '酸辣湯', 45, true],
  ['s2', '湯品／飲料', '玉米濃湯', 45, true],
  ['s3', '湯品／飲料', '古早味紅茶', 30, true],
  ['a1', '經典小菜', '涼拌小黃瓜', 40, true],
  ['a2', '經典小菜', '香滷豆干', 40, true],
  ['a3', '經典小菜', '椒香海帶絲', 40, true],
  ['gift', '冷凍禮盒', '招牌招財鍋貼禮盒 50 入', 399, true]
];

const SETTINGS_SEED = [
  ['通知信箱', ''],            // 留空＝使用部署者的 Google 帳號信箱
  ['宅配運費', 0],
  ['寄送顧客確認信', true],
  ['店名', '金煎所 GOLDEN CRISP']
];

/* ───────────── 初始化 ───────────── */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const orders = ensureSheet_(ss, CONFIG.SHEET.ORDERS, ORDER_HEADERS);
  const items = ensureSheet_(ss, CONFIG.SHEET.ITEMS, ITEM_HEADERS);
  const menu = ensureSheet_(ss, CONFIG.SHEET.MENU, MENU_HEADERS);
  const settings = ensureSheet_(ss, CONFIG.SHEET.SETTINGS, ['項目', '值']);

  if (menu.getLastRow() < 2) {
    menu.getRange(2, 1, MENU_SEED.length, MENU_HEADERS.length).setValues(MENU_SEED);
  }
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, SETTINGS_SEED.length, 2).setValues(SETTINGS_SEED);
  }

  // 訂單分頁：格式、狀態下拉選單、狀態顏色
  orders.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  orders.getRange('D2:J').setNumberFormat('@');           // 純文字，避免被當成公式
  orders.getRange('L2:N').setNumberFormat('#,##0');
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.STATUS, true).setAllowInvalid(false).build();
  orders.getRange('C2:C').setDataValidation(statusRule);

  const colorRules = [
    ['新訂單', '#fde7c8'], ['製作中', '#d9e8fb'], ['可取餐', '#d6f0d9'],
    ['已完成', '#e6e6e6'], ['已取消', '#f6d3d3']
  ].map(function (pair) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(pair[0]).setBackground(pair[1])
      .setRanges([orders.getRange('C2:C')]).build();
  });
  orders.setConditionalFormatRules(colorRules);

  items.getRange('B2:C').setNumberFormat('@');
  menu.getRange('A2:C').setNumberFormat('@');

  orders.setColumnWidths(1, 1, 130);
  orders.setColumnWidths(2, 1, 150);
  orders.setColumnWidths(5, 3, 130);
  orders.setColumnWidths(8, 3, 220);
  orders.setColumnWidths(15, 1, 160);
  menu.setColumnWidths(3, 1, 220);
  settings.setColumnWidths(1, 1, 140);
  settings.setColumnWidths(2, 1, 260);

  safeAlert_('初始化完成！\n\n接著請：\n1. 到「設定」分頁填入通知信箱（可留空）\n2. 執行 installTriggers()\n3. 部署成網頁應用程式');
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const head = sheet.getRange(1, 1, 1, headers.length);
  head.setValues([headers]).setFontWeight('bold').setBackground('#38281e').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  return sheet;
}

function safeAlert_(message) {
  try { SpreadsheetApp.getUi().alert(message); } catch (e) { Logger.log(message); }
}

/** 建立觸發條件：改成「可取餐」自動通知顧客、每天 21:30 寄摘要 */
function installTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'onStatusEdit' || fn === 'dailySummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onStatusEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('dailySummary').timeBased().atHour(21).nearMinute(30).everyDays(1)
    .inTimezone(CONFIG.TZ).create();
  safeAlert_('已建立 2 個觸發條件：狀態通知、每日摘要。');
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('金煎所')
      .addItem('初始化工作表', 'setup')
      .addItem('建立觸發條件', 'installTriggers')
      .addItem('立即寄出今日摘要', 'dailySummary')
      .addToUi();
  } catch (e) { /* 非試算表介面時略過 */ }
}

/* ───────────── 網頁應用程式 ───────────── */

function doGet(e) {
  // 健康檢查：直接用瀏覽器開網址即可看到初始化狀態（不會洩漏訂單內容）
  const out = { ok: true, service: 'golden-crisp-orders', time: nowText_() };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      out.ready = false;
      out.problem = '此指令碼沒有綁定試算表，請從試算表的「擴充功能 → Apps Script」開啟';
    } else {
      const missing = Object.keys(CONFIG.SHEET)
        .map(function (k) { return CONFIG.SHEET[k]; })
        .filter(function (name) { return !ss.getSheetByName(name); });
      const menuCount = Object.keys(readMenu_(ss)).length;
      out.ready = missing.length === 0 && menuCount > 0;
      out.menuCount = menuCount;
      if (missing.length) out.missingSheets = missing;
      if (!out.ready) out.problem = '尚未完成初始化，請在 Apps Script 執行 setup()';
    }
  } catch (err) {
    out.ready = false;
    out.problem = String(err && err.message || err);
  }
  return json_(out);
}

function doPost(e) {
  let body;
  try {
    const raw = (e && e.postData && e.postData.contents) || '';
    if (raw.length > 20000) return json_(fail_('TOO_LARGE', '資料過大'));
    body = JSON.parse(raw);
  } catch (err) {
    return json_(fail_('BAD_REQUEST', '資料格式錯誤'));
  }

  // 機器人陷阱：隱藏欄位被填寫時，假裝成功但不處理
  if (body && body.website) {
    return json_({ ok: true, orderNumber: 'GC-000000-000', total: 0 });
  }

  const cache = CacheService.getScriptCache();
  const requestId = /^[A-Za-z0-9\-]{8,64}$/.test(String(body && body.requestId || '')) ? body.requestId : '';

  // 同一份訂單重送：回傳原結果，不重複建立
  if (requestId) {
    const done = cache.get('req_' + requestId);
    if (done) return json_(JSON.parse(done));
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_(fail_('BUSY', '系統忙碌，請稍後再試'));
  }

  try {
    if (requestId) {
      const done2 = cache.get('req_' + requestId);
      if (done2) return json_(JSON.parse(done2));
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const menu = readMenu_(ss);
    const settings = readSettings_(ss);
    if (!Object.keys(menu).length) {
      return json_(fail_('NOT_SETUP', '訂單系統尚未完成初始化，請聯絡店家'));
    }

    const built = buildOrder_(body, menu, settings);
    if (!built.ok) return json_(built);

    const order = built.order;
    const rateKey = 'rl_' + order.phone.replace(/\D/g, '');
    if (cache.get(rateKey)) return json_(fail_('TOO_FAST', '您剛剛才送出訂單，請稍候再試'));

    order.orderNumber = nextOrderNumber_();
    writeOrder_(ss, order, String(body.source || 'web').slice(0, 10));

    const result = {
      ok: true,
      orderNumber: order.orderNumber,
      subtotal: order.subtotal,
      shipping: order.shipping,
      total: order.total
    };
    cache.put(rateKey, '1', CONFIG.RATE_LIMIT_SEC);
    if (requestId) cache.put('req_' + requestId, JSON.stringify(result), CONFIG.DEDUPE_SEC);

    // 通知信失敗不影響訂單
    try { notifyNewOrder_(order, settings); } catch (err) { console.error('notify failed', err); }

    return json_(result);
  } catch (err) {
    console.error(err && err.stack || err);
    return json_(fail_('SERVER_ERROR', '系統忙碌，請稍後再試'));
  } finally {
    lock.releaseLock();
  }
}

/* ───────────── 驗證與計價（以試算表價格為準） ───────────── */

function buildOrder_(body, menu, settings) {
  if (!body || typeof body !== 'object') return fail_('BAD_REQUEST', '資料格式錯誤');

  const method = String(body.method || '');
  if (['內用', '外帶', '宅配'].indexOf(method) < 0) return fail_('INVALID', '取餐方式不正確');

  const name = String(body.name || '').trim();
  const phone = String(body.phone || '').trim();
  const email = String(body.email || '').trim();
  const address = String(body.address || '').trim();
  const note = String(body.note || '').trim();

  if (!name || name.length > 60) return fail_('INVALID', '請填寫姓名（60 字內）');
  if (!/^[0-9+\-\s()]{8,20}$/.test(phone)) return fail_('INVALID', '電話格式不正確');
  if (email && (email.length > 120 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))) {
    return fail_('INVALID', 'Email 格式不正確');
  }
  if (method === '宅配' && (!address || address.length > 200)) {
    return fail_('INVALID', '宅配請填寫地址（200 字內）');
  }
  if (note.length > 300) return fail_('INVALID', '備註請在 300 字內');

  const rawItems = Array.isArray(body.items) ? body.items : [];
  if (!rawItems.length || rawItems.length > CONFIG.MAX_LINES) {
    return fail_('INVALID', '購物車內容不正確');
  }

  const merged = {};
  for (let i = 0; i < rawItems.length; i++) {
    const id = String(rawItems[i] && rawItems[i].id || '');
    const qty = Number(rawItems[i] && rawItems[i].qty);
    if (!menu[id]) return fail_('UNKNOWN_ITEM', '有商品已下架或不存在，請重新整理後再試');
    if (!menu[id].active) return fail_('SOLD_OUT', menu[id].name + ' 目前暫停供應');
    if (!Number.isInteger(qty) || qty < 1 || qty > CONFIG.MAX_QTY) {
      return fail_('INVALID', '商品數量不正確（1–' + CONFIG.MAX_QTY + '）');
    }
    merged[id] = (merged[id] || 0) + qty;
  }

  const lines = [];
  let subtotal = 0;
  let count = 0;
  Object.keys(merged).forEach(function (id) {
    const qty = Math.min(merged[id], CONFIG.MAX_QTY);
    const lineTotal = menu[id].price * qty;
    subtotal += lineTotal;
    count += qty;
    lines.push({ id: id, name: menu[id].name, price: menu[id].price, qty: qty, total: lineTotal });
  });

  const shipping = method === '宅配' ? Number(settings.shipping) || 0 : 0;

  return {
    ok: true,
    order: {
      method: method, name: name, phone: phone, email: email, address: address, note: note,
      lines: lines, count: count, subtotal: subtotal, shipping: shipping, total: subtotal + shipping
    }
  };
}

/* ───────────── 讀寫試算表 ───────────── */

function readMenu_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET.MENU);
  const rows = sheet ? sheet.getDataRange().getValues() : [];
  const menu = {};
  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][0] || '').trim();
    if (!id) continue;
    const active = rows[i][4];
    menu[id] = {
      name: String(rows[i][2] || ''),
      price: Number(rows[i][3]) || 0,
      active: active === true || String(active).toUpperCase() === 'TRUE'
    };
  }
  return menu;
}

function readSettings_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET.SETTINGS);
  const rows = sheet ? sheet.getDataRange().getValues() : [];
  const map = {};
  for (let i = 1; i < rows.length; i++) map[String(rows[i][0])] = rows[i][1];

  let owner = String(map['通知信箱'] || '').trim();
  if (!owner) {
    try { owner = Session.getEffectiveUser().getEmail(); } catch (e) { owner = ''; }
  }
  const confirm = map['寄送顧客確認信'];
  return {
    ownerEmail: owner,
    shipping: Number(map['宅配運費']) || 0,
    sendConfirm: !(confirm === false || String(confirm).toUpperCase() === 'FALSE'),
    shopName: String(map['店名'] || '金煎所')
  };
}

function nextOrderNumber_() {
  const props = PropertiesService.getScriptProperties();
  const day = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyMMdd');
  const key = 'seq_' + day;
  const seq = (Number(props.getProperty(key)) || 0) + 1;
  props.setProperty(key, String(seq));
  return 'GC-' + day + '-' + ('000' + seq).slice(-3);
}

function writeOrder_(ss, order, source) {
  const orders = ss.getSheetByName(CONFIG.SHEET.ORDERS);
  const items = ss.getSheetByName(CONFIG.SHEET.ITEMS);
  if (!orders || !items) throw new Error('找不到工作表，請先執行 setup()');

  const summary = order.lines.map(function (l) { return l.name + ' ×' + l.qty; }).join('、');

  orders.appendRow([
    order.orderNumber,
    new Date(),
    '新訂單',
    order.method,
    safe_(order.name),
    order.phone,
    safe_(order.email),
    safe_(order.address),
    safe_(order.note),
    summary,
    order.count,
    order.subtotal,
    order.shipping,
    order.total,
    '',
    source
  ]);

  const rows = order.lines.map(function (l) {
    return [order.orderNumber, l.id, l.name, l.price, l.qty, l.total];
  });
  items.getRange(items.getLastRow() + 1, 1, rows.length, ITEM_HEADERS.length).setValues(rows);
}

/** 避免使用者輸入被當成試算表公式 */
function safe_(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@\t\r]/.test(text) ? "'" + text : text;
}

/* ───────────── 通知信 ───────────── */

function notifyNewOrder_(order, settings) {
  const table = orderTableHtml_(order);
  const info =
    '<p><b>取餐方式：</b>' + esc_(order.method) + '<br>' +
    '<b>姓名：</b>' + esc_(order.name) + '<br>' +
    '<b>電話：</b>' + esc_(order.phone) + '<br>' +
    (order.email ? '<b>Email：</b>' + esc_(order.email) + '<br>' : '') +
    (order.address ? '<b>地址：</b>' + esc_(order.address) + '<br>' : '') +
    (order.note ? '<b>備註：</b>' + esc_(order.note) + '<br>' : '') + '</p>';

  if (settings.ownerEmail) {
    MailApp.sendEmail({
      to: settings.ownerEmail,
      subject: '【新訂單】' + order.orderNumber + '｜NT$ ' + order.total + '｜' + order.method,
      htmlBody: '<h3>收到新訂單 ' + esc_(order.orderNumber) + '</h3>' + info + table,
      name: settings.shopName
    });
  }

  if (settings.sendConfirm && order.email) {
    MailApp.sendEmail({
      to: order.email,
      subject: '【' + settings.shopName + '】訂單已收到 ' + order.orderNumber,
      htmlBody:
        '<p>' + esc_(order.name) + ' 您好，我們已收到您的訂單，將盡快與您確認。</p>' +
        '<p><b>訂單編號：</b>' + esc_(order.orderNumber) + '</p>' + table +
        '<p style="color:#786b5f;font-size:12px">本信件由系統自動寄出，請勿直接回覆。如需更改訂單，請以電話聯絡店家。</p>',
      name: settings.shopName
    });
  }
}

function orderTableHtml_(order) {
  const rows = order.lines.map(function (l) {
    return '<tr><td style="padding:4px 12px 4px 0">' + esc_(l.name) + '</td>' +
      '<td style="padding:4px 12px">×' + l.qty + '</td>' +
      '<td style="padding:4px 0;text-align:right">NT$ ' + l.total + '</td></tr>';
  }).join('');
  const ship = order.shipping ? '<tr><td colspan="2">運費</td><td style="text-align:right">NT$ ' + order.shipping + '</td></tr>' : '';
  return '<table style="border-collapse:collapse">' + rows + ship +
    '<tr><td colspan="2" style="padding-top:8px"><b>總金額</b></td>' +
    '<td style="padding-top:8px;text-align:right"><b>NT$ ' + order.total + '</b></td></tr></table>';
}

/** 店員把「狀態」改成「可取餐」時，自動通知顧客（需先執行 installTriggers） */
function onStatusEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET.ORDERS) return;
  if (e.range.getColumn() !== COL.STATUS || e.range.getRow() < 2) return;
  if (String(e.value) !== '可取餐') return;

  const row = e.range.getRow();
  if (String(sheet.getRange(row, COL.NOTIFY).getValue() || '').indexOf('可取餐') >= 0) return;

  const email = String(sheet.getRange(row, COL.EMAIL).getValue() || '').trim();
  if (!email) return;

  const settings = readSettings_(SpreadsheetApp.getActiveSpreadsheet());
  const orderNo = sheet.getRange(row, COL.ORDER_NO).getValue();
  const name = sheet.getRange(row, COL.NAME).getValue();
  const method = sheet.getRange(row, COL.METHOD).getValue();

  MailApp.sendEmail({
    to: email,
    subject: '【' + settings.shopName + '】您的餐點已完成 ' + orderNo,
    htmlBody: '<p>' + esc_(name) + ' 您好，您的訂單 <b>' + esc_(orderNo) + '</b> 已完成' +
      (method === '宅配' ? '，即將為您配送。' : '，歡迎前來取餐。') + '</p>',
    name: settings.shopName
  });
  sheet.getRange(row, COL.NOTIFY).setValue('可取餐通知 ' + nowText_());
}

/** 每日摘要：今天的訂單數與金額 */
function dailySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET.ORDERS);
  const settings = readSettings_(ss);
  if (!sheet || !settings.ownerEmail) return;

  const today = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd');
  const rows = sheet.getDataRange().getValues();
  let count = 0, total = 0;
  const byStatus = {};
  const lines = [];

  for (let i = 1; i < rows.length; i++) {
    const created = rows[i][COL.CREATED - 1];
    if (!(created instanceof Date)) continue;
    if (Utilities.formatDate(created, CONFIG.TZ, 'yyyy-MM-dd') !== today) continue;
    const status = String(rows[i][COL.STATUS - 1]);
    byStatus[status] = (byStatus[status] || 0) + 1;
    if (status !== '已取消') total += Number(rows[i][13]) || 0;
    count++;
    lines.push(rows[i][0] + '｜' + rows[i][3] + '｜' + rows[i][4] + '｜NT$ ' + rows[i][13] + '｜' + status);
  }

  const statusText = Object.keys(byStatus).map(function (k) { return k + ' ' + byStatus[k]; }).join('、') || '無';
  MailApp.sendEmail({
    to: settings.ownerEmail,
    subject: '【每日摘要】' + today + '｜' + count + ' 筆｜NT$ ' + total,
    htmlBody: '<p>今日訂單 <b>' + count + '</b> 筆，營業額（不含已取消）<b>NT$ ' + total + '</b></p>' +
      '<p>狀態分布：' + esc_(statusText) + '</p>' +
      '<pre style="font-size:13px">' + esc_(lines.join('\n')) + '</pre>',
    name: settings.shopName
  });
}

/* ───────────── 工具 ───────────── */

function fail_(code, message) { return { ok: false, code: code, message: message }; }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function nowText_() { return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss'); }

function esc_(text) {
  return String(text == null ? '' : text)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
