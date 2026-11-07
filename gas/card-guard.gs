/**
 * ワンバランスボクシングジム 入会カード登録（Stripe）API
 * 目的：入会申込のときにカードを登録してもらい（課金なし）、サブスクを「お試し期間つき」で作っておく。
 *       来なかった人だけ、スタッフが「サブスク開始」を押すと課金が始まる。
 *       来た人は「来店済み」を押してサブスクを取り消す（口座引落のまま）。
 *
 * 設定（Apps Script → プロジェクトの設定 → スクリプト プロパティ）
 *   STRIPE_KEY   Stripeの制限付きキー（rk_live_...）  ※チャットやメールに貼らない
 *   STAFF_PASS   スタッフ用の暗証番号（4桁）
 * 検索用ワード：OWNER_EMAIL, REMIND_DAYS, TRIAL_DAYS, RETURN_URL
 */
var OWNER_EMAIL = 'feel.good.yoshi3@gmail.com'; // 3か月たっても取り消していない人のお知らせ先
var REMIND_DAYS = 90;                            // 何日たっても「登録済み」のままなら知らせるか
var TRIAL_DAYS = 700;                            // お試し期間（Stripeの上限は730日）
var RETURN_URL = 'https://onebalanceyume-beep.github.io/onebalance-member/card-done.html';
var SHEET_NAME = 'カード登録台帳';
var HEADERS = ['会員ID','名前','メール','コース','月会費','登録日時','状態','Stripe顧客','Stripe決済画面','Stripeサブスク','更新日時'];

function prop_(k){ return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function now_(){ return new Date(); }
function ts_(d){ return Utilities.formatDate(d || now_(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'); }

/* ---------- Stripe呼び出し ---------- */
function flatten_(obj, prefix, out){
  out = out || [];
  Object.keys(obj).forEach(function(k){
    var v = obj[k], key = prefix ? prefix + '[' + k + ']' : k;
    if(v === null || v === undefined) return;
    if(Array.isArray(v)){ v.forEach(function(item, i){
      if(typeof item === 'object') flatten_(item, key + '[' + i + ']', out); else out.push([key + '[' + i + ']', item]);
    }); }
    else if(typeof v === 'object'){ flatten_(v, key, out); }
    else out.push([key, v]);
  });
  return out;
}
function stripe_(method, path, params){
  var key = prop_('STRIPE_KEY');
  if(!key) throw new Error('STRIPE_KEY が未設定です');
  var opt = { method: method, headers: { Authorization: 'Bearer ' + key }, muteHttpExceptions: true };
  var url = 'https://api.stripe.com' + path;
  var body = flatten_(params || {}).map(function(p){ return encodeURIComponent(p[0]) + '=' + encodeURIComponent(p[1]); }).join('&');
  if(method === 'get'){ if(body) url += '?' + body; }
  else { opt.payload = body; opt.contentType = 'application/x-www-form-urlencoded'; }
  var res = UrlFetchApp.fetch(url, opt);
  var json = JSON.parse(res.getContentText());
  if(res.getResponseCode() >= 300) throw new Error('Stripe: ' + ((json.error && json.error.message) || res.getResponseCode()));
  return json;
}

/* ---------- 台帳（シート） ---------- */
function sheet_(){
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('LEDGER_ID'), ss = null;
  if(id){ try{ ss = SpreadsheetApp.openById(id); }catch(e){ ss = null; } }
  if(!ss){ ss = SpreadsheetApp.create('ワンバランス入会カード登録台帳（Stripe）'); props.setProperty('LEDGER_ID', ss.getId()); }
  var sh = ss.getSheetByName(SHEET_NAME);
  if(!sh){ sh = ss.getSheets()[0]; sh.setName(SHEET_NAME); sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  return sh;
}
function rows_(){
  var sh = sheet_(), v = sh.getDataRange().getValues(), out = [];
  for(var i = 1; i < v.length; i++){
    var r = {}; HEADERS.forEach(function(h, j){ r[h] = v[i][j]; }); r._row = i + 1; out.push(r);
  }
  return out;
}
function findRow_(memberId){
  var rs = rows_(); for(var i = rs.length - 1; i >= 0; i--){ if(String(rs[i]['会員ID']) === String(memberId)) return rs[i]; } return null;
}
function setCell_(row, header, val){ sheet_().getRange(row._row, HEADERS.indexOf(header) + 1).setValue(val); }

/* ---------- 受付（お客さん側：入会アプリの完了画面から） ---------- */
function createRegistration_(p){
  var memberId = String(p.memberId || ''), fee = parseInt(p.fee, 10);
  if(!memberId || !fee || fee < 1000 || fee > 30000) throw new Error('会員IDまたは月会費が正しくありません');
  var today = ts_().slice(0, 10), cnt = 0;
  rows_().forEach(function(r){ if(String(r['登録日時']).slice(0, 10) === today) cnt++; });
  if(cnt >= 60) throw new Error('本日の受付上限です。スタッフにお声がけください');
  var ex = findRow_(memberId);
  if(ex && ex['Stripe決済画面'] && ex['状態'] === '未登録'){ return { url: ex['Stripe決済画面'] }; }
  var cust = stripe_('post', '/v1/customers', { name: p.name || '', email: p.email || '', metadata: { memberId: memberId } });
  var courseName = String(p.courseName || '月会費');
  var msg = '入会日にご来店いただけない場合は、月会費' + fee.toLocaleString() + '円（毎月）の請求が始まります。' +
            'ご来店された方は、このカード登録を取り消します。今日は課金されません。' +
            'また、口座引落ができなかった場合は、登録したクレジットカードで決済される場合があります。';
  var sess = stripe_('post', '/v1/checkout/sessions', {
    mode: 'subscription',
    customer: cust.id,
    payment_method_collection: 'always',
    line_items: [{ quantity: 1, price_data: { currency: 'jpy', unit_amount: fee, recurring: { interval: 'month' },
      product_data: { name: 'ワンバランスボクシングジム 月会費（' + courseName + '）' } } }],
    subscription_data: { trial_period_days: TRIAL_DAYS, metadata: { memberId: memberId } },
    custom_text: { submit: { message: msg } },
    metadata: { memberId: memberId },
    success_url: RETURN_URL + '?ok=1',
    cancel_url: RETURN_URL + '?ok=0'
  });
  var sh = sheet_();
  sh.appendRow([memberId, p.name || '', p.email || '', courseName, fee, ts_(), '未登録', cust.id, sess.url, '', ts_()]);
  // appendRow した行は Stripe セッションIDを別に保存
  PropertiesService.getScriptProperties().setProperty('SESS_' + memberId, sess.id);
  return { url: sess.url };
}

/* ---------- 状態の更新（Stripeに問い合わせる） ---------- */
function refresh_(r){
  if(r['状態'] === '取消済み' || r['状態'] === '課金開始' || r['状態'] === '停止中') return r;
  var sid = prop_('SESS_' + r['会員ID']);
  if(!sid) return r;
  var s = stripe_('get', '/v1/checkout/sessions/' + sid, {});
  if(s.status === 'complete' && s.subscription){
    var sub = typeof s.subscription === 'string' ? s.subscription : s.subscription.id;
    if(r['状態'] !== '登録済み'){ setCell_(r, '状態', '登録済み'); setCell_(r, 'Stripeサブスク', sub); setCell_(r, '更新日時', ts_()); r['状態'] = '登録済み'; r['Stripeサブスク'] = sub; }
  } else if(s.status === 'expired' && r['状態'] === '未登録'){
    setCell_(r, '状態', '期限切れ'); r['状態'] = '期限切れ';
  }
  return r;
}

/* ---------- スタッフ操作 ---------- */
function staffOk_(pass){ var p = prop_('STAFF_PASS'); return !!p && String(pass) === p; }
function list_(){
  return rows_().reverse().slice(0, 60).map(function(r){
    try{ r = refresh_(r); }catch(e){}
    return { memberId: r['会員ID'], name: r['名前'], course: r['コース'], fee: r['月会費'], at: r['登録日時'], status: r['状態'] };
  });
}
function startSubscription_(memberId){
  var r = findRow_(memberId); if(!r) throw new Error('会員が見つかりません');
  r = refresh_(r);
  if(r['状態'] !== '登録済み') throw new Error('カードが登録されていません（状態：' + r['状態'] + '）');
  stripe_('post', '/v1/subscriptions/' + r['Stripeサブスク'], { trial_end: 'now' });
  setCell_(r, '状態', '課金開始'); setCell_(r, '更新日時', ts_());
  return { ok: true };
}
function markVisited_(memberId){
  var r = findRow_(memberId); if(!r) throw new Error('会員が見つかりません');
  r = refresh_(r);
  if((r['状態'] === '登録済み' || r['状態'] === '課金開始' || r['状態'] === '停止中') && r['Stripeサブスク']){ stripe_('delete', '/v1/subscriptions/' + r['Stripeサブスク'], {}); }
  setCell_(r, '状態', '取消済み'); setCell_(r, '更新日時', ts_());
  return { ok: true };
}

function pauseSubscription_(memberId){
  var r = findRow_(memberId); if(!r) throw new Error('会員が見つかりません');
  if(r['状態'] !== '課金開始') throw new Error('課金中ではありません（状態：' + r['状態'] + '）');
  stripe_('post', '/v1/subscriptions/' + r['Stripeサブスク'], { pause_collection: { behavior: 'void' } });
  setCell_(r, '状態', '停止中'); setCell_(r, '更新日時', ts_());
  return { ok: true };
}
function resumeSubscription_(memberId){
  var r = findRow_(memberId); if(!r) throw new Error('会員が見つかりません');
  if(r['状態'] !== '停止中') throw new Error('停止中ではありません（状態：' + r['状態'] + '）');
  stripe_('post', '/v1/subscriptions/' + r['Stripeサブスク'], { pause_collection: '' });
  setCell_(r, '状態', '課金開始'); setCell_(r, '更新日時', ts_());
  return { ok: true };
}

/* ---------- 3か月たっても取り消していない人をお知らせ（毎日1回） ---------- */
function remindOld(){
  var limit = now_().getTime() - REMIND_DAYS * 86400000, list = [];
  rows_().forEach(function(r){
    if(r['状態'] !== '登録済み') return;
    var d = new Date(String(r['登録日時']).replace(/-/g, '/'));
    if(isNaN(d.getTime()) || d.getTime() > limit) return;
    var last = prop_('REMIND_' + r['会員ID']);
    if(last && now_().getTime() - Number(last) < 30 * 86400000) return; // 30日に1回まで
    list.push(r);
  });
  if(!list.length) return 0;
  var body = 'カード登録のまま、3か月以上たっている人です。\n来店済みなら「来店済み」を、来ていないなら「サブスク開始」を押してください。\n\n';
  list.forEach(function(r){
    body += '・' + r['名前'] + '様（会員ID ' + r['会員ID'] + '／' + r['コース'] + '／登録 ' + r['登録日時'] + '）\n';
    PropertiesService.getScriptProperties().setProperty('REMIND_' + r['会員ID'], String(now_().getTime()));
  });
  body += '\n管理画面：https://onebalanceyume-beep.github.io/onebalance-member/card-guard.html';
  MailApp.sendEmail(OWNER_EMAIL, '【ワンバランス】カード登録のまま3か月たっている人がいます（' + list.length + '名）', body);
  return list.length;
}
function setupDailyReminder(){ // 最初に1回だけ実行
  ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction() === 'remindOld') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('remindOld').timeBased().everyDays(1).atHour(9).create();
}

/* ---------- 入口（JSONP） ---------- */
function out_(cb, obj){
  var s = JSON.stringify(obj);
  return ContentService.createTextOutput(cb ? cb + '(' + s + ')' : s)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
function doGet(e){
  var p = e.parameter || {}, cb = p.callback || '';
  try{
    if(p.action === 'register'){ return out_(cb, { ok: true, url: createRegistration_(p).url }); }
    if(!staffOk_(p.pass)) return out_(cb, { ok: false, error: '暗証番号が違います' });
    if(p.action === 'list'){ return out_(cb, { ok: true, items: list_() }); }
    if(p.action === 'start'){ return out_(cb, startSubscription_(p.memberId)); }
    if(p.action === 'pause'){ return out_(cb, pauseSubscription_(p.memberId)); }
    if(p.action === 'resume'){ return out_(cb, resumeSubscription_(p.memberId)); }
    if(p.action === 'visited'){ return out_(cb, markVisited_(p.memberId)); }
    return out_(cb, { ok: false, error: '不明なactionです' });
  }catch(err){ return out_(cb, { ok: false, error: String(err.message || err) }); }
}
