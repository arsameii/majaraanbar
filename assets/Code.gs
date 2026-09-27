// ===== انبار ماجرا بوک — بک‌اند گوگل شیت =====
// 1) رمز مدیر را همین‌جا عوض کنید (فقط شما باید بدانید):
const ADMIN_PASSWORD = '1234';

const CATS_SEED = [{"id": "k", "name": "کتاب کودک"}, {"id": "t", "name": "تحریر"}, {"id": "d", "name": "دفتر"}, {"id": "b", "name": "برد گیم"}, {"id": "o2jdp9", "name": "اسباب بازی"}];
const ITEMS_SEED = [{"id": "e5ft73", "cat": "b", "name": "Celever dice", "qty": 2}, {"id": "lp1nj4", "cat": "d", "name": "سری استرالیا - دفتر ۱۰۰ برگ جلد سخت پیل", "qty": 30}, {"id": "iy7gbf", "cat": "t", "name": "مداد رنگی ۲۴ رنگ فابر کاستل", "qty": 5}, {"id": "pg13j2", "cat": "t", "name": "پاستل ۱۲ رنگ آریا", "qty": 10}, {"id": "p0j3j4", "cat": "t", "name": "خمیر بازی سطلی بزرگ آریا", "qty": 5}, {"id": "z3l7ps", "cat": "o2jdp9", "name": "آجره ۲۱ قطعه", "qty": 3}, {"id": "yb787v", "cat": "k", "name": "هر کسی چیزی می‌رونه", "qty": 5}, {"id": "gcklip", "cat": "k", "name": "هر کسی کاری میکنه", "qty": 3}, {"id": "y9orbh", "cat": "k", "name": "مهمانی در دریا", "qty": 6}];

// یک بار این تابع را اجرا کنید تا برگه‌ها ساخته شوند و اطلاعات فعلی وارد شود
function setup() {
  const ss = SpreadsheetApp.getActive();
  const mk = (name, header) => {
    let s = ss.getSheetByName(name);
    if (!s) { s = ss.insertSheet(name); s.appendRow(header); }
    return s;
  };
  const cats = mk('cats', ['id', 'name']);
  const items = mk('items', ['id', 'cat', 'name', 'qty', 'created']);
  const meta = mk('meta', ['updated', '']);
  if (cats.getLastRow() < 2) CATS_SEED.forEach(c => cats.appendRow([c.id, c.name]));
  if (items.getLastRow() < 2) ITEMS_SEED.forEach((i, n) => items.appendRow([i.id, i.cat, i.name, i.qty, n]));
  meta.getRange('B1').setValue(new Date());
}

function sh_(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function rows_(n) {
  const v = sh_(n).getDataRange().getValues(); v.shift();
  return v.filter(r => r[0] !== '');
}
function read_() {
  const cats = rows_('cats').map(r => ({ id: String(r[0]), name: String(r[1]) }));
  const items = rows_('items').map(r => ({ id: String(r[0]), cat: String(r[1]), name: String(r[2]), qty: Number(r[3]) || 0, created: Number(r[4]) || 0 }));
  const u = sh_('meta').getRange('B1').getValue();
  return { cats: cats, items: items, updated: u ? new Date(u).toISOString() : null };
}
function touch_() { sh_('meta').getRange('B1').setValue(new Date()); }
function find_(sheet, id) {
  const last = sheet.getLastRow(); if (last < 2) return 0;
  const ids = sheet.getRange(1, 1, last, 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return 0;
}

// نمایش سایت (فایل Index.html)
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// اگر فایل HTML جدا باز شود از این مسیرها با گوگل شیت حرف می‌زند
function doPost(e) { return out_(api_post(e.postData.contents)); }

function doGet(e) {
  if (e && e.parameter && e.parameter.api) return out_(read_());
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('موجودی انبار — ماجرا بوک')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

// این دو تابع را صفحه صدا می‌زند
function api_get() { return read_(); }
function api_post(s) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return handle_(JSON.parse(s)); }
  catch (err) { return { error: String(err) }; }
  finally { lock.releaseLock(); }
}

function handle_(b) {
  if (b.action === 'take') return take_(b);
  if (b.password !== ADMIN_PASSWORD) return { error: 'bad_password' };
  if (b.op === 'check') return { ok: true };
  const items = sh_('items'), cats = sh_('cats');
  const r = find_(items, b.id);
  const extra = {};
  if (b.op === 'add') items.appendRow([Utilities.getUuid().slice(0, 6), b.cat, b.name, Number(b.qty) || 0, Date.now()]);
  else if (b.op === 'update' && r) {
    if (b.name !== undefined) items.getRange(r, 3).setValue(b.name);
    if (b.qty !== undefined) items.getRange(r, 4).setValue(Math.max(0, Number(b.qty) || 0));
  }
  else if (b.op === 'adjust' && r) items.getRange(r, 4).setValue(Math.max(0, (Number(items.getRange(r, 4).getValue()) || 0) + Number(b.delta)));
  else if (b.op === 'delete' && r) items.deleteRow(r);
  else if (b.op === 'addcat') { const id = Utilities.getUuid().slice(0, 6); cats.appendRow([id, b.name]); extra.newCat = id; }
  else if (b.op === 'delcat') {
    const cr = find_(cats, b.id); if (cr) cats.deleteRow(cr);
    for (let i = items.getLastRow(); i >= 2; i--) if (String(items.getRange(i, 2).getValue()) === String(b.id)) items.deleteRow(i);
  }
  touch_();
  return Object.assign(read_(), extra);
}

// کم کردن موجودی توسط بچه‌ها (بدون رمز). rid از ثبت دوباره‌ی تکراری جلوگیری می‌کند.
function take_(b) {
  const cache = CacheService.getScriptCache();
  if (b.rid) { const prev = cache.get('r' + b.rid); if (prev) return JSON.parse(prev); }
  const items = sh_('items'), results = [];
  (b.lines || []).forEach(l => {
    const r = find_(items, l.id); let got = 0;
    if (r) {
      const cur = Number(items.getRange(r, 4).getValue()) || 0;
      got = Math.max(0, Math.min(cur, Number(l.n) || 0));
      if (got > 0) items.getRange(r, 4).setValue(cur - got);
    }
    results.push({ id: l.id, got: got });
  });
  touch_();
  const res = { results: results, data: read_() };
  if (b.rid) cache.put('r' + b.rid, JSON.stringify(res), 21600);
  return res;
}
