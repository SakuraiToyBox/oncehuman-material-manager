/**
 * Once Human クラフト管理アプリ (GAS + スプレッドシート)
 *
 * 初回のみ setup() を1回実行するとシートが作られます。
 * 合言葉(任意): スクリプトプロパティ PASSPHRASE
 * Discord通知(任意): スクリプトプロパティ DISCORD_WEBHOOK
 */

const SH = {
  recipe: 'レシピ',
  stock: '在庫',
  log: '納品ログ',
  target: '目標',
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Once Human クラフト管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 初期セットアップ(エディタから1回だけ実行) */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートに紐づいたスクリプトとして作成してください(拡張機能 > Apps Script)');
  const defs = [
    [SH.recipe, ['アイテム', '素材', '必要数', '1回の生産数'], [
      ['サンプル武器', 'サンプル部品', 2, 1],
      ['サンプル武器', 'サンプル鉄', 10, 1],
      ['サンプル部品', 'サンプル鉄', 3, 1],
      ['サンプル部品', 'サンプル木材', 5, 1],
    ]],
    [SH.stock, ['素材', '数量'], [['サンプル鉄', 20], ['サンプル木材', 5]]],
    [SH.log, ['日時', 'メンバー', '素材', '数量'], []],
    [SH.target, ['アイテム', '個数'], [['サンプル武器', 3]]],
  ];
  defs.forEach(([name, header, rows]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
      if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
      sh.setFrozenRows(1);
    }
  });
}

/* ---------- 共通 ---------- */

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。setup() を実行してください');
  return sh;
}

function rows_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues()
    .filter(r => String(r[0]).trim() !== '');
}

function props_() {
  return PropertiesService.getScriptProperties();
}

function checkAuth_(pass) {
  const p = props_().getProperty('PASSPHRASE');
  if (p && String(pass || '') !== p) throw new Error('合言葉が違います');
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function num_(v) {
  const n = Number(v);
  if (!isFinite(n)) throw new Error('数値が不正です: ' + v);
  return n;
}

/* ---------- 読み取りAPI ---------- */

function getData() {
  const recipes = rows_(SH.recipe).map(r => ({
    item: String(r[0]).trim(),
    material: String(r[1]).trim(),
    qty: Number(r[2]) || 0,
    yield: Number(r[3]) || 1,
  }));
  const stock = {};
  rows_(SH.stock).forEach(r => { stock[String(r[0]).trim()] = Number(r[1]) || 0; });
  const logRows = rows_(SH.log);
  const log = logRows.slice(-100).reverse().map(r => ({
    time: r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Tokyo', 'MM/dd HH:mm') : String(r[0]),
    member: String(r[1]),
    material: String(r[2]),
    qty: Number(r[3]) || 0,
  }));
  const targets = rows_(SH.target).map(r => ({ item: String(r[0]).trim(), count: Number(r[1]) || 0 }));
  return {
    recipes, stock, log, targets,
    needPass: !!props_().getProperty('PASSPHRASE'),
    hasWebhook: !!props_().getProperty('DISCORD_WEBHOOK'),
  };
}

/* ---------- 書き込みAPI ---------- */

/** 在庫を指定値に設定 */
function setStock(pass, material, qty) {
  checkAuth_(pass);
  material = String(material).trim();
  qty = num_(qty);
  if (!material) throw new Error('素材名が空です');
  return withLock_(() => {
    const sh = sheet_(SH.stock);
    const last = sh.getLastRow();
    const names = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
    const i = names.indexOf(material);
    if (i >= 0) sh.getRange(i + 2, 2).setValue(qty);
    else sh.appendRow([material, qty]);
    return true;
  });
}

/** 納品: 在庫に加算してログに記録 */
function deliver(pass, member, material, qty) {
  checkAuth_(pass);
  member = String(member).trim();
  material = String(material).trim();
  qty = num_(qty);
  if (!member || !material) throw new Error('メンバー名と素材名は必須です');
  return withLock_(() => {
    const sh = sheet_(SH.stock);
    const last = sh.getLastRow();
    const names = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
    const i = names.indexOf(material);
    if (i >= 0) {
      const cell = sh.getRange(i + 2, 2);
      cell.setValue((Number(cell.getValue()) || 0) + qty);
    } else {
      sh.appendRow([material, qty]);
    }
    sheet_(SH.log).appendRow([new Date(), member, material, qty]);
    return true;
  });
}

/** レシピを全置換。text: 1行1素材「アイテム,素材,必要数[,1回の生産数]」(カンマ/タブ区切り) */
function saveRecipes(pass, text) {
  checkAuth_(pass);
  const out = [];
  String(text || '').split(/\r?\n/).forEach((line, idx) => {
    line = line.trim();
    if (!line || line.startsWith('#')) return;
    const c = line.split(/[,\t]/).map(s => s.trim());
    if (c.length < 3) throw new Error((idx + 1) + '行目: 「アイテム,素材,必要数」の形式で入力してください');
    const qty = num_(c[2]);
    const y = c[3] ? num_(c[3]) : 1;
    if (y <= 0) throw new Error((idx + 1) + '行目: 生産数は1以上にしてください');
    out.push([c[0], c[1], qty, y]);
  });
  return withLock_(() => {
    const sh = sheet_(SH.recipe);
    const last = sh.getLastRow();
    if (last >= 2) sh.getRange(2, 1, last - 1, 4).clearContent();
    if (out.length) sh.getRange(2, 1, out.length, 4).setValues(out);
    return out.length;
  });
}

/** 目標(作りたいアイテムと個数)を全置換 */
function saveTargets(pass, targets) {
  checkAuth_(pass);
  const out = (targets || [])
    .filter(t => t && String(t.item).trim() && Number(t.count) > 0)
    .map(t => [String(t.item).trim(), Number(t.count)]);
  return withLock_(() => {
    const sh = sheet_(SH.target);
    const last = sh.getLastRow();
    if (last >= 2) sh.getRange(2, 1, last - 1, 2).clearContent();
    if (out.length) sh.getRange(2, 1, out.length, 2).setValues(out);
    return out.length;
  });
}

/** Discordへ通知 */
function notifyDiscord(pass, text) {
  checkAuth_(pass);
  const url = props_().getProperty('DISCORD_WEBHOOK');
  if (!url) throw new Error('スクリプトプロパティ DISCORD_WEBHOOK が未設定です');
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ content: String(text).slice(0, 1900) }),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('Discord通知に失敗しました (HTTP ' + code + ')');
  return true;
}
