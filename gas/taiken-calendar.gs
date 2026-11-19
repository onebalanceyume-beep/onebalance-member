/**
 * 体験受付 → Googleカレンダー連携（APIキー不要のGAS方式）
 * 体験受付ページ(taiken-reception.html)の送信時に、受付した時刻から30分の予定を作り、
 * アンケートの回答を「説明」欄にそのまま入れる。
 * ※既存の受付GAS（submitTaiken）には触らない、カレンダー専用の別GAS。
 */

// ===== 設定エリア =====
var CAL_CONFIG = {
  CALENDAR_ID: '',        // 空ならこのGASを作ったアカウントのメインカレンダー。別カレンダーに入れたい時だけIDを入れる
  DURATION_MIN: 30,       // 体験の長さ（分）
  TITLE_PREFIX: '【体験受付】', // 予定タイトルの頭（満員判定キーワード【/長尾】【COACHING】とは別にしてある）
  TIMEZONE: 'Asia/Tokyo'
};
// ===== ここまで =====

function doGet(e) {
  var cb = e.parameter.callback;
  var out;
  try {
    if (e.parameter.action !== 'addTaikenEvent') throw new Error('actionが違います');
    var d = JSON.parse(e.parameter.data);
    out = addTaikenEvent_(d);
  } catch (err) {
    out = { success: false, message: String(err) };
  }
  var json = JSON.stringify(out);
  return ContentService
    .createTextOutput(cb ? cb + '(' + json + ')' : json)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function addTaikenEvent_(d) {
  var cal = CAL_CONFIG.CALENDAR_ID
    ? CalendarApp.getCalendarById(CAL_CONFIG.CALENDAR_ID)
    : CalendarApp.getDefaultCalendar();
  if (!cal) throw new Error('カレンダーが見つかりません');

  var start = new Date();
  var end = new Date(start.getTime() + CAL_CONFIG.DURATION_MIN * 60000);

  var kind = String(d.taikenType || '').replace('体験', '');   // ボクシング / キックボクササイズ
  var title = CAL_CONFIG.TITLE_PREFIX + (d.name || '') + '様（' + kind + '）';

  var age = calcAge_(d.birthday);
  var desc = [
    '■ 体験メニュー：' + (d.taikenType || ''),
    '■ お名前：' + (d.name || ''),
    '■ 性別：' + (d.gender || ''),
    '■ 年齢：' + (age !== '' ? age + '歳' : '') + '（生年月日 ' + (d.birthday || '') + '）',
    '■ 電話番号：' + (d.tel || ''),
    '■ 住所：' + (d.address || ''),
    '',
    '【アンケート】',
    '■ 知ったきっかけ：' + (d.kikkake || ''),
    '■ 運動経験：' + (d.undou || ''),
    '■ 今のイメージ：' + String(d.images || '').split(',').join('、'),
    '',
    '受付時刻：' + Utilities.formatDate(start, CAL_CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm')
  ].join('\n');

  var ev = cal.createEvent(title, start, end, { description: desc });
  return { success: true, eventId: ev.getId() };
}

function calcAge_(birthday) {
  if (!birthday) return '';
  var b = new Date(birthday);
  if (isNaN(b)) return '';
  var t = new Date();
  var age = t.getFullYear() - b.getFullYear();
  if (t.getMonth() < b.getMonth() || (t.getMonth() === b.getMonth() && t.getDate() < b.getDate())) age--;
  return age;
}

// 権限の許可用（初回に1回だけ実行）
function setupOnce() {
  CalendarApp.getDefaultCalendar().getName();
}

// 動作テスト用（実行するとカレンダーにテスト予定が1件入る）
function testAdd() {
  Logger.log(JSON.stringify(addTaikenEvent_({
    taikenType: 'ボクシング体験', name: '長尾 儀弘', gender: '男性', tel: '09012345678',
    address: '米子市車尾6丁目8-59', birthday: '1985-04-01', kikkake: 'YouTube',
    undou: '少しある', images: 'カッコいい,楽しそう'
  })));
}
