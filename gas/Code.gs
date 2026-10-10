/**
 * “英語は友達” 利用者サイト — 保存先(Google スプレッドシート)の窓口
 *
 * 使い方(くわしくは gas/README.md):
 *   1. スプレッドシートを作り、「拡張機能 → Apps Script」にこのファイルの中身を貼る
 *   2. 関数「setup」を一度だけ実行する(シートが3枚できます)
 *   3. 「デプロイ → 新しいデプロイ → ウェブアプリ」で、
 *      次のユーザーとして実行: 自分 / アクセスできるユーザー: 全員
 *   4. できた URL(…/exec)を、サイトの CONFIG.gasUrl に貼る
 *
 * 書き込みを消したいときは、その行の「非表示」にチェックを入れます(行は消さないでください)。
 */

var SHEETS = {
  sub:    { name: "編集版",   head: ["ID", "日時", "ニックネーム", "改善ポイント", "URL", "非表示"] },
  thread: { name: "スレッド", head: ["スレッドID", "作成日時", "更新日時", "タイトル", "レス数", "非表示"] },
  post:   { name: "書き込み", head: ["スレッドID", "番号", "日時", "名前", "ID", "本文", "非表示"] }
};
var LIMIT = { nick: 30, point: 600, title: 60, body: 1500, name: 30 };
var URL_OK = /^https:\/\/claude\.ai\/(?:code\/)?artifact\/[A-Za-z0-9_-]+\/?$/;
var MAX_RES = 1000;          // 1スレッドの上限
var WAIT_SEC = 30;           // 同じ人の連続書き込みの間隔
var MAX_PER_MIN = 30;        // サイト全体の1分あたりの上限
var NONAME = "名無しさん";

/* ---------- 準備 ---------- */
function setup() {
  Object.keys(SHEETS).forEach(function (k) { sheet_(k); });
  salt_();
}

function sheet_(k) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), d = SHEETS[k];
  var sh = ss.getSheetByName(d.name);
  if (!sh) {
    sh = ss.insertSheet(d.name);
    sh.getRange(1, 1, 1, d.head.length).setValues([d.head]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function salt_() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty("SALT");
  if (!s) { s = Utilities.getUuid(); p.setProperty("SALT", s); }
  return s;
}

/* ---------- 読み取り ---------- */
function doGet(e) {
  var a = (e && e.parameter && e.parameter.a) || "";
  try {
    if (a === "subs") return out_({ ok: true, items: listSubs_() });
    if (a === "threads") return out_({ ok: true, threads: listThreads_() });
    if (a === "thread") return out_(getThread_(Number(e.parameter.id)));
    return out_({ ok: false, error: "bad_request" });
  } catch (err) {
    return out_({ ok: false, error: "server" });
  }
}

function rows_(k) {
  var sh = sheet_(k), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, SHEETS[k].head.length).getValues();
}
function hidden_(v) { return v === true || String(v).toUpperCase() === "TRUE"; }
function iso_(v) { return v instanceof Date ? v.toISOString() : String(v || ""); }

function listSubs_() {
  return rows_("sub").filter(function (r) { return r[0] !== "" && !hidden_(r[5]); })
    .map(function (r) { return { id: Number(r[0]), at: iso_(r[1]), nick: String(r[2]), point: String(r[3]), url: String(r[4]) }; })
    .filter(function (x) { return URL_OK.test(x.url); })
    .sort(function (a, b) { return b.id - a.id; })
    .slice(0, 100);
}

function listThreads_() {
  return rows_("thread").filter(function (r) { return r[0] !== "" && !hidden_(r[5]); })
    .map(function (r) { return { id: Number(r[0]), created: iso_(r[1]), updated: iso_(r[2]), title: String(r[3]), count: Number(r[4]) || 0 }; })
    .sort(function (a, b) { return a.updated < b.updated ? 1 : a.updated > b.updated ? -1 : 0; })
    .slice(0, 100);
}

function getThread_(id) {
  var t = findThread_(id);
  if (!t) return { ok: false, error: "not_found" };
  var posts = rows_("post").filter(function (r) { return Number(r[0]) === id; })
    .map(function (r) {
      var p = { no: Number(r[1]), at: iso_(r[2]) };
      if (hidden_(r[6])) p.deleted = true;
      else { p.name = String(r[3]) || NONAME; p.uid = String(r[4]); p.body = String(r[5]); }
      return p;
    })
    .sort(function (a, b) { return a.no - b.no; });
  return { ok: true, thread: { id: id, title: t.title, count: t.count, created: t.created, updated: t.updated }, posts: posts };
}

function findThread_(id) {
  if (!(id > 0)) return null;
  var rs = rows_("thread");
  for (var i = 0; i < rs.length; i++) {
    if (Number(rs[i][0]) === id && !hidden_(rs[i][5])) {
      return { row: i + 2, title: String(rs[i][3]), count: Number(rs[i][4]) || 0, created: iso_(rs[i][1]), updated: iso_(rs[i][2]) };
    }
  }
  return null;
}

/* ---------- 書き込み ---------- */
function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: "bad_request" }); }
  if (!d || typeof d !== "object") return out_({ ok: false, error: "bad_request" });
  if (d.hp) return out_({ ok: true });                       // 迷惑投稿よけ(人には見えない欄)
  var uid = String(d.uid || "");
  if (!/^[A-Za-z0-9]{8,64}$/.test(uid)) return out_({ ok: false, error: "bad_request" });

  var v = validate_(d);
  if (v) return out_({ ok: false, error: v });

  var cache = CacheService.getScriptCache();
  if (cache.get("rl:" + uid)) return out_({ ok: false, error: "rate_limited" });
  var minKey = "rl:min:" + Math.floor(Date.now() / 60000);
  var cnt = Number(cache.get(minKey) || 0);
  if (cnt >= MAX_PER_MIN) return out_({ ok: false, error: "busy" });

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return out_({ ok: false, error: "busy" });
  try {
    var res;
    if (d.action === "sub") res = addSub_(d);
    else if (d.action === "thread") res = addThread_(d, uid);
    else if (d.action === "reply") res = addReply_(d, uid);
    else res = { ok: false, error: "bad_request" };
    if (res.ok) {
      cache.put("rl:" + uid, "1", WAIT_SEC);
      cache.put(minKey, String(cnt + 1), 120);
    }
    return out_(res);
  } catch (err) {
    return out_({ ok: false, error: "server" });
  } finally {
    lock.releaseLock();
  }
}

function str_(v) { return String(v == null ? "" : v).replace(/\r\n?/g, "\n").replace(/[\u0000-\u0009\u000B-\u001F\u007F]/g, "").trim(); }

function validate_(d) {
  if (d.action === "sub") {
    var nick = str_(d.nick), point = str_(d.point), url = str_(d.url);
    if (!nick || nick.length > LIMIT.nick || !point || point.length > LIMIT.point) return "invalid";
    if (!URL_OK.test(url)) return "invalid_url";
    return "";
  }
  if (d.action === "thread") {
    var title = str_(d.title), body = str_(d.body);
    if (!title || title.length > LIMIT.title || !body || body.length > LIMIT.body || str_(d.name).length > LIMIT.name) return "invalid";
    return "";
  }
  if (d.action === "reply") {
    if (!(Number(d.id) > 0)) return "invalid";
    var b = str_(d.body);
    if (!b || b.length > LIMIT.body || str_(d.name).length > LIMIT.name) return "invalid";
    return "";
  }
  return "bad_request";
}

// 先頭に ' を付けて、数式や日付として解釈されないようにする(読み出すと ' は付きません)
function txt_(v) { return "'" + v; }

function nextId_(k) {
  var rs = rows_(k), m = 0;
  rs.forEach(function (r) { var n = Number(r[0]); if (n > m) m = n; });
  return m + 1;
}

function dailyId_(uid) {
  var day = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, uid + "|" + day + "|" + salt_(), Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes).replace(/[^A-Za-z0-9]/g, "").slice(0, 8);
}

// 行を足して、最後の列(非表示)をチェックボックスにする
function append_(k, row) {
  var sh = sheet_(k);
  sh.appendRow(row);
  sh.getRange(sh.getLastRow(), row.length).insertCheckboxes();
}

function addSub_(d) {
  var id = nextId_("sub");
  append_("sub", [id, new Date(), txt_(str_(d.nick)), txt_(str_(d.point)), str_(d.url), false]);
  return { ok: true, id: id };
}

function addThread_(d, uid) {
  var id = nextId_("thread"), now = new Date();
  append_("thread", [id, now, now, txt_(str_(d.title)), 1, false]);
  append_("post", [id, 1, now, txt_(str_(d.name) || NONAME), dailyId_(uid), txt_(str_(d.body)), false]);
  return { ok: true, id: id, no: 1 };
}

function addReply_(d, uid) {
  var id = Number(d.id), t = findThread_(id);
  if (!t) return { ok: false, error: "not_found" };
  if (t.count >= MAX_RES) return { ok: false, error: "full" };
  var no = t.count + 1, now = new Date();
  var sh = sheet_("thread");
  sh.getRange(t.row, 3).setValue(now);
  sh.getRange(t.row, 5).setValue(no);
  append_("post", [id, no, now, txt_(str_(d.name) || NONAME), dailyId_(uid), txt_(str_(d.body)), false]);
  return { ok: true, id: id, no: no };
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
