// ════════════════════════════════════════════════
// AuthGuard.gs ― Googleログインの本人確認（GAS側）  v3
//
// ■ 許可リスト
//   ・共通の「アクセス許可」スプレッドシート（シート名「許可アカウント」）を、管理画面から編集します。
//     スクリプト プロパティ ACL_SHEET_ID にそのスプレッドシートのIDを登録してください。
//   ・予備として、スクリプト プロパティ ALLOWED_EMAILS（カンマ区切り）も有効です。
//   ・管理者：シートの「権限」列が「管理者」の人、またはスクリプト プロパティ ADMIN_EMAILS の人。
//   ・許可リストが空のときは全員を拒否します。
//   ・許可のない人は、画面の「利用を申請する」から申請できます。申請はシート「申請中」に入り、
//     管理画面で承認すると、許可アカウントに移ります（管理者にはメールで通知されます）。
//
// ■ 初回だけ：エディタで関数 authorizeOnce を選んで「実行」し、権限を承認してください。
//   （名前の末尾が _ の関数は、実行メニューに表示されません）
//
// ■ 導入のしかた（既存のコードは置き換えません）
//   1. このファイルを、新しいスクリプトファイル「AuthGuard」として追加する（既存の Code.gs などはそのまま）。
//   2. 既存の doGet(e) と doPost(e) の、{ の次の行（先頭）に、次の1行を足す。
//        const __g = authGuard_(e); if (__g) return __g;
//      （引数の名前が e でないときは、その名前に合わせる）
//   3. 関数 authorizeOnce を実行して権限を承認し、新バージョンで再デプロイする。
// ════════════════════════════════════════════════

const AG_CLIENT_ID = '782190614995-5d4fur401smd1gpnnvs43h3m2q2q2ar3.apps.googleusercontent.com';
const AG_ACL_SHEET = '許可アカウント';
const AG_REQ_SHEET = '申請中';
const AG_ADMIN_URL = 'https://uchiyamazion.github.io/sion-portal/access/';

// 初回だけ実行する関数：外部通信とスプレッドシートの権限の承認画面を出す
function authorizeOnce() {
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=test', { muteHttpExceptions: true });
  CacheService.getScriptCache().put('authorize_test', '1', 5);
  MailApp.getRemainingDailyQuota();   // 管理者へのメール通知の権限の承認用
  const id = PropertiesService.getScriptProperties().getProperty('ACL_SHEET_ID');
  if (id) { const sh = agAclSheet_(); Logger.log('許可シート: ' + sh.getParent().getName() + ' / 登録 ' + Math.max(0, sh.getLastRow() - 1) + ' 件'); }
  Logger.log('承認OK。ACL_SHEET_ID=' + (id ? '設定あり' : '未設定（ALLOWED_EMAILSだけで動作）'));
}

// ---------- 許可リスト ----------
function agListProp_(name) {
  const v = PropertiesService.getScriptProperties().getProperty(name) || '';
  return v.split(/[\s,;]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
}
function agAclSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('ACL_SHEET_ID');
  if (!id) return null;
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(AG_ACL_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AG_ACL_SHEET);
    sh.getRange(1, 1, 1, 4).setValues([['メールアドレス', '名前', '権限', '登録日時']]);
  }
  return sh;
}
function agReqSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('ACL_SHEET_ID');
  if (!id) return null;
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(AG_REQ_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AG_REQ_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['メールアドレス', '名前', '申請日時']]);
  }
  return sh;
}
function agReqRows_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .map(r => ({ email: String(r[0]).trim().toLowerCase(), name: String(r[1] || ''), at: r[2] instanceof Date ? r[2].getTime() : 0 }))
    .filter(r => r.email);
}

function agAclRows_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('acl_rows');
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* 再取得 */ } }
  let rows = [];
  try {
    const sh = agAclSheet_();
    if (sh && sh.getLastRow() > 1) {
      rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
        .map(r => ({ email: String(r[0]).trim().toLowerCase(), name: String(r[1] || ''), role: String(r[2] || '') }))
        .filter(r => r.email);
    }
  } catch (err) { rows = []; }
  cache.put('acl_rows', JSON.stringify(rows), 60);
  return rows;
}
function agAllowed_() { return agListProp_('ALLOWED_EMAILS').concat(agAclRows_().map(r => r.email)); }
function agAdmins_() { return agListProp_('ADMIN_EMAILS').concat(agAclRows_().filter(r => r.role === '管理者').map(r => r.email)); }

// ---------- IDトークンの確認 ----------
// 戻り値：{ ok:true, email } | { ok:false, reason:'unauthorized' } | { ok:false, reason:'forbidden', email }
function agVerify_(idToken) {
  if (!idToken) return { ok: false, reason: 'unauthorized' };
  const cache = CacheService.getScriptCache();
  const key = 'idt_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken)).slice(0, 40);
  let info = null;
  const hit = cache.get(key);
  if (hit) { try { info = JSON.parse(hit); } catch (e) { info = null; } }
  if (!info) {
    const res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { ok: false, reason: 'unauthorized' };
    const r = JSON.parse(res.getContentText());
    info = { aud: r.aud, iss: r.iss, email: r.email, email_verified: r.email_verified, exp: r.exp };
    const ttl = Math.min(300, Number(info.exp) - Math.floor(Date.now() / 1000));
    if (ttl > 0) cache.put(key, JSON.stringify(info), ttl);
  }
  const now = Math.floor(Date.now() / 1000);
  const iss = String(info.iss || '');
  if (info.aud !== AG_CLIENT_ID) return { ok: false, reason: 'unauthorized' };
  if (iss !== 'accounts.google.com' && iss !== 'https://accounts.google.com') return { ok: false, reason: 'unauthorized' };
  if (Number(info.exp) <= now) return { ok: false, reason: 'unauthorized' };
  if (String(info.email_verified) !== 'true') return { ok: false, reason: 'unauthorized' };
  const email = String(info.email || '').toLowerCase();
  if (agAllowed_().indexOf(email) < 0) return { ok: false, reason: 'forbidden', email: email };
  return { ok: true, email: email };
}

function agJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- 利用申請（許可のない人も、Googleにログインしていれば申請できる） ----------
function agRequestHandle_(p, a) {
  if (!a.email) return agJson_({ status: 'error', success: false, error: a.reason, data: a.reason, message: a.reason });
  if (a.ok) return agJson_({ status: 'success', success: true, data: { state: 'already_allowed' } });
  const sh = agReqSheet_();
  if (!sh) return agJson_({ status: 'error', success: false, error: 'no_sheet', data: 'no_sheet', message: '申請を受け付けられません。管理者に直接連絡してください' });
  const cache = CacheService.getScriptCache();
  const rk = 'req_' + a.email;
  if (cache.get(rk)) return agJson_({ status: 'success', success: true, data: { state: 'already_requested' } });
  cache.put(rk, '1', 60);
  const name = String(p.name || '').slice(0, 60);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let isNew = false;
  try {
    const rows = agReqRows_(sh);
    if (rows.some(r => r.email === a.email)) return agJson_({ status: 'success', success: true, data: { state: 'already_requested' } });
    if (rows.length >= 200) return agJson_({ status: 'error', success: false, error: 'full', data: 'full', message: '申請が多すぎます。管理者に直接連絡してください' });
    sh.appendRow([a.email, name, new Date()]);
    isNew = true;
  } finally { lock.releaseLock(); }
  if (isNew) agNotify_(a.email, name);
  return agJson_({ status: 'success', success: true, data: { state: 'requested' } });
}
function agNotify_(email, name) {
  try {
    const to = agAdmins_().filter((v, i, arr) => arr.indexOf(v) === i);
    if (!to.length) return;
    MailApp.sendEmail({
      to: to.join(','),
      subject: '【アクセス申請】' + email,
      body: '利用の申請がありました。\n\nメールアドレス：' + email + '\n名前：' + (name || '（未入力）') +
            '\n\n承認・却下は、次の管理画面で行ってください。\n' + AG_ADMIN_URL
    });
  } catch (err) { /* メール送信に失敗しても申請自体は有効 */ }
}

// ---------- 管理画面用の操作（acl_list / acl_add / acl_remove / acl_approve / acl_reject）。管理者だけが使える ----------
function agAclHandle_(action, p, a) {
  const bad = (code, msg) => agJson_({ status: 'error', success: false, error: code, data: code, message: msg });
  if (!a.email) return bad(a.reason, a.reason);
  if (agAdmins_().indexOf(a.email) < 0) return bad('not_admin', '管理者のみ操作できます');
  const props = { me: a.email };
  const sheetId = PropertiesService.getScriptProperties().getProperty('ACL_SHEET_ID');
  const result = () => {
    CacheService.getScriptCache().remove('acl_rows');
    props.rows = agAclRows_();
    props.pending = agReqRows_(agReqSheet_());
    props.propEmails = agListProp_('ALLOWED_EMAILS');
    props.propAdmins = agListProp_('ADMIN_EMAILS');
    props.sheetReady = !!sheetId;
    return agJson_({ status: 'success', success: true, data: props });
  };
  if (action === 'acl_list') {
    if (!sheetId) { props.rows = []; props.pending = []; props.propEmails = agListProp_('ALLOWED_EMAILS'); props.propAdmins = agListProp_('ADMIN_EMAILS'); props.sheetReady = false; return agJson_({ status: 'success', success: true, data: props }); }
    return result();
  }
  const sh = agAclSheet_();
  if (!sh) return bad('no_sheet', 'スクリプト プロパティ ACL_SHEET_ID が未設定です');
  const email = String(p.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return bad('bad_email', 'メールアドレスの形式が正しくありません');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const last = sh.getLastRow();
    const emails = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase()) : [];
    const idx = emails.indexOf(email);
    const psh = agReqSheet_();
    const pend = agReqRows_(psh);
    const pidx = pend.findIndex(r => r.email === email);
    const dropPending = () => { if (pidx >= 0) psh.deleteRow(pidx + 2); };
    if (action === 'acl_add' || action === 'acl_approve') {
      const name = String(p.name || (pidx >= 0 ? pend[pidx].name : '') || '').slice(0, 60);
      const role = p.role === '管理者' ? '管理者' : '';
      if (idx >= 0) sh.getRange(idx + 2, 1, 1, 3).setValues([[email, name, role]]);
      else sh.appendRow([email, name, role, new Date()]);
      dropPending();
    } else if (action === 'acl_reject') {
      dropPending();
    } else if (action === 'acl_remove') {
      if (email === a.email) return bad('self', '自分自身は削除できません');
      if (idx >= 0) sh.deleteRow(idx + 2);
    } else {
      return bad('bad_action', '不明な操作です');
    }
  } finally { lock.releaseLock(); }
  return result();
}

// 許可されていれば null、そうでなければエラー応答（JSON）を返す
function authGuard_(e) {
  const p = (e && e.parameter) || {};
  let token = p.idToken || '';
  if (!token && e && e.postData && e.postData.contents) {
    try { token = (JSON.parse(e.postData.contents) || {}).idToken || ''; } catch (err) { /* JSON以外は無視 */ }
  }
  const a = agVerify_(token);
  if (p.action === 'acl_request') return agRequestHandle_(p, a);
  if (/^acl_/.test(p.action || '')) return agAclHandle_(p.action, p, a);
  if (a.ok) return null;
  return agJson_({ status: 'error', success: false, ok: false, error: a.reason, data: a.reason, message: a.reason });
}
