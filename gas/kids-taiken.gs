/**
 * ワンバランス ジュニア体験予約 GAS（APIキー不要）
 * ----------------------------------------------------
 * 役割
 *  1) getKidsDays : 直近の土曜16:30を返す。カレンダーに【キッズ休み】があるその土曜は「お休み(off)」
 *  2) saveKids    : 予約を受付 → カレンダー登録 → スプレッドシート記録 → 確認メール送信
 *
 * 使い方：gas/kids-taiken-SETUP.md を見てください。
 * ----------------------------------------------------
 */

/* ===== 設定エリア（ここだけ変えればOK） ===== */
var CONFIG = {
  TZ: 'Asia/Tokyo',
  START_TIME: '16:30',            // ジュニアクラスの開始時刻
  DURATION_MIN: 90,               // カレンダー上の長さ（分）
  WEEKS: 8,                       // 何週先まで受け付けるか
  SAME_DAY_CUTOFF_HOUR: 12,       // 当日の土曜は、この時刻（時）まで受付

  // この言葉がタイトルに含まれる予定がある土曜は「お休み」。あとから増やせます
  OFF_KEYWORDS: ['【ジュニア休み】', '【キッズ休み】'],  // 旧表記も残してある（すでに入れた予定を無効にしないため）

  // 休み予定を探すカレンダー（'' = このスクリプトを実行するアカウントのメインカレンダー）
  // 別のカレンダーに【キッズ休み】を入れる場合は、そのカレンダーのIDをここに足す
  CALENDAR_IDS: [''],

  ADMIN_EMAILS: ['one.balance.yume@gmail.com'],      // 予約が入ったら通知するアドレス
  GUEST_EMAILS: ['feel.good.yoshi3@gmail.com'],      // 予定にゲストとして追加（個人カレンダーにも出る）

  BRING_ITEMS: ['軍手', 'タオル', '飲み物', '上履き'],   // 当日の持ち物（メールに出る）
  GYM_NAME: 'ONE-BALANCE BOXING GYM（ワンバランスボクシングジム）',
  GYM_ADDRESS: '鳥取県米子市車尾6丁目8-59',
  GYM_TEL: '0859-49-5275',
  LINE_URL: 'https://lin.ee/2dtDdkF',

  SHEET_TITLE: '【ワンバランス】キッズ体験予約一覧',
  SHEET_NAME: 'キッズ体験予約',
  SHEET_ID_PROP: 'KIDS_SHEET_ID'
};
/* ===== 設定エリアここまで ===== */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = String(p.callback || '');
  var out;
  try {
    if (p.action === 'getKidsDays') out = getKidsDays_(p);
    else if (p.action === 'saveKids') out = saveKids_(p);
    else out = { error: 'unknown action' };
  } catch (err) {
    out = { error: 'サーバーエラーが発生しました。お電話でご連絡ください。' };
    console.error(err && err.stack || err);
  }
  return respond_(out, cb);
}

function respond_(obj, cb) {
  var json = JSON.stringify(obj);
  if (/^[A-Za-z0-9_$.]+$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 日付まわり ---------- */
function todayStr_() { return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd'); }
function dateAt_(ds, hm) { return Utilities.parseDate(ds + ' ' + hm, CONFIG.TZ, 'yyyy-MM-dd HH:mm'); }
function isSaturday_(ds) { return Utilities.formatDate(dateAt_(ds, '12:00'), CONFIG.TZ, 'u') === '6'; }

/** 受付できる土曜の一覧（日付文字列） */
function upcomingSaturdays_(weeks) {
  var res = [];
  var now = new Date();
  var base = dateAt_(todayStr_(), '12:00').getTime();
  for (var i = 0; i < 400 && res.length < weeks; i++) {
    var ds = Utilities.formatDate(new Date(base + i * 86400000), CONFIG.TZ, 'yyyy-MM-dd');
    if (!isSaturday_(ds)) continue;
    if (i === 0) {
      var h = Number(Utilities.formatDate(now, CONFIG.TZ, 'H'));
      if (h >= CONFIG.SAME_DAY_CUTOFF_HOUR) continue;   // 当日の受付は締切時刻まで
    }
    res.push(ds);
  }
  return res;
}

/** その日に【キッズ休み】があるか（繰り返し予定・終日予定にも対応） */
function isOffDay_(ds) {
  var day = dateAt_(ds, '12:00');
  for (var i = 0; i < CONFIG.CALENDAR_IDS.length; i++) {
    var id = CONFIG.CALENDAR_IDS[i];
    var cal = id ? CalendarApp.getCalendarById(id) : CalendarApp.getDefaultCalendar();
    if (!cal) continue;
    var evs = cal.getEventsForDay(day);
    for (var j = 0; j < evs.length; j++) {
      var title = evs[j].getTitle() || '';
      for (var k = 0; k < CONFIG.OFF_KEYWORDS.length; k++) {
        if (title.indexOf(CONFIG.OFF_KEYWORDS[k]) !== -1) return true;
      }
    }
  }
  return false;
}

/* ---------- 予約できる日の一覧 ---------- */
function getKidsDays_(p) {
  var weeks = Math.min(Math.max(parseInt(p.weeks, 10) || CONFIG.WEEKS, 1), 12);
  var days = upcomingSaturdays_(weeks).map(function (ds) {
    return { date: ds, status: isOffDay_(ds) ? 'off' : 'ok' };
  });
  return { ok: true, time: CONFIG.START_TIME, days: days };
}

/* ---------- 予約の保存 ---------- */
function clean_(s, max) {
  return String(s == null ? '' : s).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, max || 200);
}

function saveKids_(p) {
  var ds = clean_(p.date, 10);
  var name = clean_(p.name, 60);
  var phone = clean_(p.phone, 30), email = clean_(p.email, 100), address = clean_(p.address, 150);
  var note = clean_(p.note, 300); var source = clean_(p.source, 40);

  // お子さま（きょうだい対応：children=[{name,grade}]。古い形式 child/grade も受け付ける）
  var kids = [];
  try { kids = JSON.parse(p.children || '[]'); } catch (e) { kids = []; }
  if (!kids.length && p.child) kids = [{ name: p.child, grade: p.grade }];
  kids = kids.slice(0, 5).map(function (k) { return { name: clean_(k.name, 60), grade: clean_(k.grade, 10) }; });
  var pax = kids.length;

  if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) return { error: '日にちが正しくありません。' };
  if (!pax || !name || !phone || !address) return { error: '入力されていない項目があります。' };
  for (var q = 0; q < kids.length; q++) { if (!kids[q].name || !kids[q].grade) return { error: '入力されていない項目があります。' }; }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { error: 'メールアドレスが正しくありません。' };
  if (String(p.minorConsent) !== '1') return { error: '保護者の同意が必要です。' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // その日がいま予約できる土曜か、サーバー側でも確認（改ざん対策）
    var okDays = upcomingSaturdays_(CONFIG.WEEKS);
    if (okDays.indexOf(ds) === -1) return { error: 'この日は受付できません。', reload: true };
    if (isOffDay_(ds)) return { error: 'この日はお休みです。別の日をお選びください。', reload: true };

    // 同じメール・同じ日・同じお子さまの二重予約を防ぐ
    var sheet = getSheet_(); if (sheet.getRange(1, 15).getValue() === '') { sheet.getRange(1, 15).setValue('流入元'); }
    var last = sheet.getLastRow();
    if (last > 1) {
      var rows = sheet.getRange(2, 1, last - 1, 13).getValues();
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (String(r[2]) !== ds || String(r[8]).toLowerCase() !== email.toLowerCase() || r[12] === 'キャンセル') continue;
        for (var m = 0; m < kids.length; m++) {
          if (String(r[6]) === kids[m].name) return { error: kids[m].name + 'さんはこの日にすでに予約されています。' };
        }
      }
    }

    var id = makeId_();
    var start = dateAt_(ds, CONFIG.START_TIME);
    var end = new Date(start.getTime() + CONFIG.DURATION_MIN * 60000);
    var cal = CalendarApp.getDefaultCalendar();
    var stamp = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');

    // お子さま1人につき、予定1件＋シート1行（予約IDは同じ）
    kids.forEach(function (k, idx) {
      var title = '【ジュニア体験】' + k.name + '様（' + k.grade + '）' + (pax > 1 ? '※きょうだい' + pax + '名' : '');
      var desc = '予約ID: ' + id + '\n保護者: ' + name + '\n電話: ' + phone + '\nメール: ' + email +
        '\n住所: ' + address + '\n同時申込: ' + pax + '名' + (pax > 1 ? '（' + kids.map(function (x) { return x.name; }).join('・') + '）' : '') +
        '\n備考: ' + (note || 'なし') + '\n流入元: ' + (source || '未回答') +
        '\n体験代1,100円（当日お支払い）\n保護者の同意: あり';
      var opts = { description: desc, location: CONFIG.GYM_ADDRESS };
      if (CONFIG.GUEST_EMAILS.length) { opts.guests = CONFIG.GUEST_EMAILS.join(','); opts.sendInvites = false; }
      var ev = cal.createEvent(title, start, end, opts);
      sheet.appendRow([stamp, id, ds, CONFIG.START_TIME, name, phone, k.name, k.grade, email, address, pax, note, '予約中', ev.getId(), source]);
    });

    sendMails_(id, ds, name, kids, email, phone, address, note, source);
    return { ok: true, reservationId: id };
  } finally {
    lock.releaseLock();
  }
}

function makeId_() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', s = '';
  for (var i = 0; i < 8; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

/* ---------- スプレッドシート ---------- */
function getSheet_() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty(CONFIG.SHEET_ID_PROP);
  var ss;
  if (ssId) { try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.SHEET_TITLE);
    props.setProperty(CONFIG.SHEET_ID_PROP, ss.getId());
  }
  var sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(CONFIG.SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    // 列：A受付日時 B予約ID C体験日 D開始 E保護者 F電話 G お子さま H学年 I メール J住所 K人数 L備考 M状態 N予定ID
    sh.appendRow(['受付日時', '予約ID', '体験日', '開始', '保護者名', '電話', 'お子さま名', '学年', 'メール', '住所', '人数', '備考', '状態', 'カレンダー予定ID', '流入元']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------- メール ---------- */
function jpDate_(ds) {
  var d = dateAt_(ds, '12:00');
  var w = ['日', '月', '火', '水', '木', '金', '土'][Number(Utilities.formatDate(d, CONFIG.TZ, 'u')) % 7];
  return Utilities.formatDate(d, CONFIG.TZ, 'M月d日') + '（' + w + '）';
}

function esc_(t) {
  return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function sendMails_(id, ds, name, kids, email, phone, address, note, source) {
  var pax = kids.length;
  var child = kids.map(function (k) { return k.name; }).join('・');
  var grade = kids.map(function (k) { return k.grade; }).join('・');
  var when = jpDate_(ds) + ' ' + CONFIG.START_TIME + '〜（60〜90分ほど）';
  var kidList = kids.map(function (k) { return k.name + '（' + k.grade + '）'; });
  var total = (1100 * pax).toLocaleString('en-US');
  var pay = pax > 1 ? '1,100円 × ' + pax + '名 = ' + total + '円（当日、現金・PayPay・クレジットカード）' : '1,100円（当日、現金・PayPay・クレジットカード）';
  var items = CONFIG.BRING_ITEMS;

  // 文字だけのメール（HTMLが表示できない場合用）。1行を短くして改行位置が崩れないようにする
  var lines = [
    name + ' 様', '',
    'ジュニアボクシング体験に', 'お申し込みいただき、ありがとうございます。', '下記の内容でご予約を承りました。', '',
    '【ご予約内容】',
    '・予約ID：' + id,
    '・日時：' + when,
    '・お子さま（' + pax + '名）：' + kidList.join('、'),
    '・体験代：' + pay, '',
    '【当日の持ち物】'
  ];
  items.forEach(function (x) { lines.push('・' + x); });
  lines.push('※動きやすい服装でお越しください', '',
    '【保護者の方へ】',
    '・体験には、保護者の方の同意が必要です', '',
    '【変更・キャンセル】',
    '・LINE：' + CONFIG.LINE_URL,
    '・お電話：' + CONFIG.GYM_TEL, '',
    CONFIG.GYM_NAME, CONFIG.GYM_ADDRESS, 'TEL ' + CONFIG.GYM_TEL);
  var body = lines.join('\n');

  // 見やすいHTMLメール（箇条書き）
  var li = function (t) { return '<li style="margin:4px 0;">' + t + '</li>'; };
  var h = function (t) { return '<p style="margin:20px 0 6px;font-weight:bold;color:#DD0B1B;">' + t + '</p>'; };
  var html =
    '<div style="font-family:Helvetica,Arial,sans-serif;font-size:15px;line-height:1.7;color:#222;max-width:520px;">' +
    '<p>' + esc_(name) + ' 様</p>' +
    '<p>ジュニアボクシング体験に<br>お申し込みいただき、ありがとうございます。<br>下記の内容でご予約を承りました。</p>' +
    h('■ ご予約内容') +
    '<ul style="margin:0;padding-left:20px;">' +
      li('予約ID：<b>' + esc_(id) + '</b>') +
      li('日時：<b>' + esc_(when) + '</b>') +
      li('お子さま（' + pax + '名）：' + esc_(kidList.join('、'))) +
      li('体験代：' + esc_(pay)) +
    '</ul>' +
    h('■ 当日の持ち物') +
    '<ul style="margin:0;padding-left:20px;">' + items.map(function (x) { return li(esc_(x)); }).join('') + '</ul>' +
    '<p style="margin:6px 0 0;">※動きやすい服装でお越しください</p>' +
    h('■ 保護者の方へ') +
    '<ul style="margin:0;padding-left:20px;">' + li('体験には、保護者の方の同意が必要です') + '</ul>' +
    h('■ 変更・キャンセル') +
    '<ul style="margin:0;padding-left:20px;">' +
      li('LINE：<a href="' + esc_(CONFIG.LINE_URL) + '">' + esc_(CONFIG.LINE_URL) + '</a>') +
      li('お電話：' + esc_(CONFIG.GYM_TEL)) +
    '</ul>' +
    '<p style="margin-top:24px;border-top:1px solid #ddd;padding-top:12px;">' + esc_(CONFIG.GYM_NAME) + '<br>' + esc_(CONFIG.GYM_ADDRESS) + '<br>TEL ' + esc_(CONFIG.GYM_TEL) + '</p>' +
    '</div>';

  MailApp.sendEmail({
    to: email,
    subject: '【ご予約完了】ジュニアボクシング体験 ' + jpDate_(ds) + ' ' + CONFIG.START_TIME + '〜（' + CONFIG.GYM_NAME + '）',
    body: body, htmlBody: html, name: 'ワンバランスボクシングジム'
  });

  if (CONFIG.ADMIN_EMAILS.length) {
    var adminBody = '【ジュニア体験の予約が入りました】\n\n予約ID：' + id + '\n日時：' + when + '\nお子さま（' + pax + '名）：' + kidList.join('、') + '\n体験代：' + total + '円\n' +
      '保護者：' + name + '\n電話：' + phone + '\nメール：' + email + '\n住所：' + address + '\n備考：' + (note || 'なし') + '\n流入元：' + (source || '未回答') + '\n';
    MailApp.sendEmail({ to: CONFIG.ADMIN_EMAILS.join(','), subject: '【ジュニア体験予約】' + child + '様（' + grade + '）' + pax + '名 ' + jpDate_(ds), body: adminBody, name: 'キッズ体験予約' });
  }
}

/* ---------- 最初に1回だけ実行（権限の許可＋動作確認用） ---------- */
function setupOnce() {
  var sh = getSheet_();
  var days = getKidsDays_({ weeks: 4 });
  Logger.log('スプレッドシート: ' + sh.getParent().getUrl());
  Logger.log(JSON.stringify(days));
}
