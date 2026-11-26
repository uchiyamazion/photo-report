/* gas-auth.js ― Googleログイン共通モジュール
 * 画面を開くとGoogleログインを求め、GAS（script.google.com/macros/...）への通信すべてに
 * ログイン証明書（idToken）を自動で付ける。許可アカウントかどうかの判定はGAS側（AuthGuard.gs）で行う。
 * 使い方：各ページの <head> の先頭で <script src="gas-auth.js"></script> を読み込むだけ。 */
(function(){
  if (window.__gasAuth) return; window.__gasAuth = true;

  var CLIENT_ID = '782190614995-5d4fur401smd1gpnnvs43h3m2q2q2ar3.apps.googleusercontent.com';
  var KEY = 'gasauth_idt';
  var GAS_RE = /^https:\/\/script\.google\.com\/macros\/s\//;
  var origFetch = window.fetch ? window.fetch.bind(window) : null;
  if (!origFetch) return;

  /* ---------- トークン ---------- */
  function exp(t){
    try { var p = t.split('.')[1].replace(/-/g,'+').replace(/_/g,'/'); return JSON.parse(atob(p)).exp * 1000; }
    catch(e){ return 0; }
  }
  function load(){ try { return localStorage.getItem(KEY) || ''; } catch(e){ return ''; } }
  var TOKEN = load();
  function getToken(){
    var t = TOKEN || load();
    if (t && exp(t) > Date.now() + 30000) { TOKEN = t; return t; }
    return '';
  }
  function clearToken(){ TOKEN = ''; try { localStorage.removeItem(KEY); } catch(e){} }

  /* ---------- ログイン画面 ---------- */
  var waiters = [], gsiReady = false, overlay = null, msgEl = null;
  function buildOverlay(){
    if (overlay || !document.body) return;
    overlay = document.createElement('div');
    overlay.id = 'gasauth-overlay';
    overlay.style.cssText = 'position:fixed;inset:0;z-index:2147483000;display:none;align-items:center;justify-content:center;background:#0f1b2d;font-family:"Hiragino Sans","Noto Sans JP",sans-serif;';
    overlay.innerHTML =
      '<div style="background:#fff;color:#1b2430;border-radius:16px;padding:28px 24px;margin:0 16px;width:100%;max-width:340px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.45);">'
      + '<div style="font-size:17px;font-weight:700;margin-bottom:6px;">ログインしてください</div>'
      + '<div style="font-size:12.5px;color:#5d6b7c;line-height:1.7;margin-bottom:18px;">許可されたGoogleアカウントでログインします</div>'
      + '<div id="gasauth-btn" style="display:flex;justify-content:center;min-height:44px;"></div>'
      + '<div id="gasauth-msg" style="display:none;color:#c0392b;font-size:12.5px;line-height:1.7;margin-top:14px;"></div>'
      + '</div>';
    document.body.appendChild(overlay);
    msgEl = overlay.querySelector('#gasauth-msg');
  }
  function showLogin(msg){
    return new Promise(function(resolve){
      waiters.push(resolve);
      var go = function(){
        buildOverlay();
        if (msg) { msgEl.textContent = msg; msgEl.style.display = 'block'; } else { msgEl.style.display = 'none'; }
        overlay.style.display = 'flex';
        initGsi();
      };
      if (document.body) go(); else document.addEventListener('DOMContentLoaded', go);
    });
  }
  function onCredential(resp){
    TOKEN = resp.credential;
    try { localStorage.setItem(KEY, TOKEN); } catch(e){}
    if (overlay) overlay.style.display = 'none';
    var w = waiters; waiters = []; w.forEach(function(f){ f(); });
  }
  function loadGsi(){
    if (document.getElementById('gasauth-gsi')) return;
    var s = document.createElement('script');
    s.id = 'gasauth-gsi'; s.src = 'https://accounts.google.com/gsi/client'; s.async = true; s.defer = true;
    document.head.appendChild(s);
  }
  function initGsi(){
    loadGsi();
    if (!(window.google && google.accounts && google.accounts.id)) { setTimeout(initGsi, 200); return; }
    if (!gsiReady) {
      google.accounts.id.initialize({ client_id: CLIENT_ID, callback: onCredential, auto_select: true, cancel_on_tap_outside: false });
      google.accounts.id.renderButton(overlay.querySelector('#gasauth-btn'), { theme:'outline', size:'large', text:'signin_with', locale:'ja', width:260 });
      gsiReady = true;
    }
    google.accounts.id.prompt();
  }
  async function authToken(){
    var t = getToken(); if (t) return t;
    await showLogin(TOKEN ? 'ログインの有効期限が切れました。もう一度ログインしてください' : '');
    return getToken();
  }

  /* ---------- fetch をラップ：GASへの通信に idToken を付け、認証エラーなら再ログインして再送 ---------- */
  async function authError(res){
    try {
      var cl = +(res.headers.get('content-length') || 0);
      if (cl > 2000) return '';
      var t = await res.clone().text();
      if (t.length > 2000) return '';
      var j = JSON.parse(t);
      var v = [j && j.error, j && j.data, j && j.message].map(String);
      if (v.indexOf('forbidden') >= 0) return 'forbidden';
      if (v.indexOf('unauthorized') >= 0) return 'unauthorized';
    } catch(e){}
    return '';
  }
  async function gasFetch(input, init){
    var token = await authToken();
    var isReq = (typeof input !== 'string');
    var url = isReq ? input.url : input;
    var u = url + (url.indexOf('?') >= 0 ? '&' : '?') + 'idToken=' + encodeURIComponent(token);
    var target = isReq ? new Request(u, input) : u;
    var res = await origFetch(target, init);
    var bad = await authError(res);
    if (bad) {
      var forbidden = (bad === 'forbidden');
      clearToken();
      if (forbidden && window.google && google.accounts && google.accounts.id) google.accounts.id.disableAutoSelect();
      await showLogin(forbidden
        ? 'このアカウントには利用の許可がありません。管理者に許可を依頼するか、別のアカウントでログインしてください'
        : 'ログインの有効期限が切れました。もう一度ログインしてください');
      return gasFetch(input, init);
    }
    return res;
  }
  window.fetch = function(input, init){
    var url = (typeof input === 'string') ? input : (input && input.url) || '';
    if (GAS_RE.test(url)) return gasFetch(input, init);
    return origFetch(input, init);
  };

  /* ---------- 開いた直後にログイン画面を出す（ログイン済みなら何もしない） ---------- */
  function gate(){ if (!getToken()) showLogin(''); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', gate); else gate();

  /* 手動ログアウト用（コンソールや画面のボタンから window.gasAuthSignOut() で呼べる） */
  window.gasAuthSignOut = function(){
    clearToken();
    if (window.google && google.accounts && google.accounts.id) google.accounts.id.disableAutoSelect();
    location.reload();
  };
})();
