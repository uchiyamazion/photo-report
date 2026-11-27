/**
 * 写真報告書アプリ バックエンド (Google Apps Script)
 * ------------------------------------------------------
 * 元の「写真報告書送付用原紙」xlsm/マクロと同じレイアウトを
 * Googleスプレッドシート上に自動生成し、xlsxとして書き出す。
 *
 * デプロイ方法:
 *  1. script.google.com で新規プロジェクトを作成し、このファイルを貼り付け
 *  2. 初回のみ initSetup() を手動実行 → マスタ用スプレッドシートと
 *     写真保存用/出力用フォルダを自動作成し、Script Propertiesに保存
 *  3. 「デプロイ」→「ウェブアプリ」
 *      - 実行ユーザー: 自分
 *      - アクセスできるユーザー: 全員（社内共有ならリンクを知る全員でも可）
 *  4. 発行されたURLを index.html の GAS_URL に設定
 */

// ===================== 初期セットアップ =====================

function initSetup() {
  const props = PropertiesService.getScriptProperties();

  if (!props.getProperty('MASTER_SS_ID')) {
    const ss = SpreadsheetApp.create('写真報告書アプリ_マスタデータ');
    const sh1 = ss.getSheets()[0];
    sh1.setName('型式_製造番号');
    sh1.appendRow(['型式', '製造番号']);
    ss.insertSheet('工事内容').appendRow(['工事内容']);
    ss.insertSheet('宛先会社').appendRow(['会社名1', '会社名2']);
    props.setProperty('MASTER_SS_ID', ss.getId());
    Logger.log('マスタスプレッドシート作成: ' + ss.getUrl());
  }

  if (!props.getProperty('PHOTO_FOLDER_ID')) {
    const folder = DriveApp.createFolder('写真報告書アプリ_写真');
    props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  }

  if (!props.getProperty('OUTPUT_FOLDER_ID')) {
    const folder = DriveApp.createFolder('写真報告書アプリ_出力');
    props.setProperty('OUTPUT_FOLDER_ID', folder.getId());
  }

  ensureReportsSheet_();

  Logger.log('セットアップ完了');
}

/**
 * 「報告書一覧」シートを確保する（既存ユーザーが後からinitSetup()を再実行しても追加されるように）。
 */
function ensureReportsSheet_() {
  const ss = getMasterSheet_();
  let sh = ss.getSheetByName('報告書一覧');
  if (!sh) {
    sh = ss.insertSheet('報告書一覧');
    sh.appendRow(['fileId', 'spreadsheetId', '物件名', '作業日', '作成日時', 'ファイル名', 'dataFileId']);
  }
  // D列(作業日)・E列(作成日時)は、日付らしい文字列を書き込むと自動で日付型に変換されて
  // 後で読み出す際にUTCのタイムスタンプに化けてしまうため、常に文字列として扱うよう固定する。
  sh.getRange('D:E').setNumberFormat('@');
  return sh;
}

function getProps_() {
  return PropertiesService.getScriptProperties();
}

/**
 * ★1回だけ手動実行してください★
 * 以前のバージョンで作られたまま残っている中間生成物のGoogleスプレッドシート
 * （名前が「写真報告書_」で始まるもの）をまとめてゴミ箱に移動する。
 * 現在のバージョンでは、生成のたびに中間スプレッドシートは自動削除されるため、
 * 今後はこのファイルは増えません（マスタデータのスプレッドシートは対象外です）。
 *
 * 実行方法: GASエディタ上部の関数選択で cleanupOrphanedSpreadsheets_ を選び、実行ボタンを押す。
 * 実行後、実行ログに削除件数が出ます。
 */
function cleanupOrphanedSpreadsheets_() {
  const query = "mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false and name contains '写真報告書_'";
  const files = DriveApp.searchFiles(query);
  let count = 0;
  while (files.hasNext()) {
    const f = files.next();
    // 「写真報告書_」で始まるものだけを対象にする（部分一致ではなく前方一致で安全確認）
    if (f.getName().indexOf('写真報告書_') === 0) {
      f.setTrashed(true);
      count++;
    }
  }
  Logger.log('削除した中間スプレッドシート数: ' + count);
}

// ===================== Web API =====================

function doGet(e) {
  const __g = authGuard_(e); if (__g) return __g;
  const action = (e.parameter && e.parameter.action) || 'master';
  if (action === 'master') {
    return jsonOut_(getMasterData_());
  } else if (action === 'reports') {
    return jsonOut_(getReportsList_());
  } else if (action === 'reportData') {
    return jsonOut_(getReportData_(e.parameter.fileId));
  }
  return jsonOut_({ error: 'unknown action' });
}

function doPost(e) {
  const __g = authGuard_(e); if (__g) return __g;
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'invalid JSON' });
  }

  const action = body.action;
  try {
    if (action === 'saveMaster') {
      return jsonOut_(saveMaster_(body));
    } else if (action === 'updateMaster') {
      return jsonOut_(updateMaster_(body));
    } else if (action === 'deleteMaster') {
      return jsonOut_(deleteMaster_(body));
    } else if (action === 'generateReport') {
      return jsonOut_(generateReport_(body));
    } else if (action === 'deleteReport') {
      return jsonOut_(deleteReport_(body));
    }
    return jsonOut_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===================== マスタ管理 =====================

function getMasterSheet_() {
  const id = getProps_().getProperty('MASTER_SS_ID');
  if (!id) throw new Error('未セットアップです。initSetup() を実行してください。');
  return SpreadsheetApp.openById(id);
}

function getMasterData_() {
  const ss = getMasterSheet_();

  const kataSheet = ss.getSheetByName('型式_製造番号');
  const kataValues = kataSheet.getDataRange().getValues().slice(1)
    .filter(r => r[0] !== '');
  const byKata = {};
  kataValues.forEach(r => {
    const kata = String(r[0]);
    const seizo = String(r[1] || '');
    if (!byKata[kata]) byKata[kata] = [];
    if (seizo) byKata[kata].push(seizo);
  });

  const kojiSheet = ss.getSheetByName('工事内容');
  const kojiValues = kojiSheet.getDataRange().getValues().slice(1)
    .map(r => r[0]).filter(v => v !== '');

  const companySheet = ss.getSheetByName('宛先会社');
  const companyValues = companySheet.getDataRange().getValues().slice(1)
    .filter(r => r[0] !== '')
    .map(r => ({ company1: String(r[0]), company2: String(r[1] || '') }));

  return {
    ok: true,
    kataList: Object.keys(byKata),
    seizoByKata: byKata,
    kojiList: kojiValues,
    companyList: companyValues
  };
}

/**
 * body = { type: 'kata'|'koji', kata, seizo, koji }
 */
function saveMaster_(body) {
  const ss = getMasterSheet_();
  if (body.type === 'kata') {
    const sh = ss.getSheetByName('型式_製造番号');
    if (!body.kata) throw new Error('型式が空です');
    if (body.seizo) {
      // 型式+製造番号のペアを追加（重複はスキップ）
      const data = sh.getDataRange().getValues();
      const exists = data.some(r => r[0] === body.kata && r[1] === body.seizo);
      if (!exists) sh.appendRow([body.kata, body.seizo]);
    } else {
      // 製造番号なしで型式だけ登録したい場合
      const data = sh.getDataRange().getValues();
      const exists = data.some(r => r[0] === body.kata);
      if (!exists) sh.appendRow([body.kata, '']);
    }
  } else if (body.type === 'koji') {
    const sh = ss.getSheetByName('工事内容');
    if (!body.koji) throw new Error('工事内容が空です');
    const data = sh.getDataRange().getValues();
    const exists = data.some(r => r[0] === body.koji);
    if (!exists) sh.appendRow([body.koji]);
  } else if (body.type === 'company') {
    const sh = ss.getSheetByName('宛先会社');
    if (!body.company1) throw new Error('会社名が空です');
    const data = sh.getDataRange().getValues();
    const exists = data.some(r => r[0] === body.company1 && r[1] === (body.company2 || ''));
    if (!exists) sh.appendRow([body.company1, body.company2 || '']);
  }
  return { ok: true, data: getMasterData_() };
}

/**
 * 既存の1件を新しい内容に書き換える。
 * body = {
 *   type:'kata',    oldKata, oldSeizo, kata, seizo
 *   type:'koji',    oldKoji, koji
 *   type:'company', oldCompany1, oldCompany2, company1, company2
 * }
 */
function updateMaster_(body) {
  const ss = getMasterSheet_();

  if (body.type === 'kata') {
    if (!body.kata) throw new Error('型式が空です');
    const sh = ss.getSheetByName('型式_製造番号');
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === body.oldKata && data[i][1] === (body.oldSeizo || '')) {
        sh.getRange(i + 1, 1, 1, 2).setValues([[body.kata, body.seizo || '']]);
        break;
      }
    }
  } else if (body.type === 'koji') {
    if (!body.koji) throw new Error('工事内容が空です');
    const sh = ss.getSheetByName('工事内容');
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === body.oldKoji) {
        sh.getRange(i + 1, 1).setValue(body.koji);
        break;
      }
    }
  } else if (body.type === 'company') {
    if (!body.company1) throw new Error('会社名が空です');
    const sh = ss.getSheetByName('宛先会社');
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === body.oldCompany1 && data[i][1] === (body.oldCompany2 || '')) {
        sh.getRange(i + 1, 1, 1, 2).setValues([[body.company1, body.company2 || '']]);
        break;
      }
    }
  }
  return { ok: true, data: getMasterData_() };
}

/**
 * body = { type: 'kata'|'koji', kata, seizo, koji }
 */
function deleteMaster_(body) {
  const ss = getMasterSheet_();
  if (body.type === 'kata') {
    const sh = ss.getSheetByName('型式_製造番号');
    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === body.kata && data[i][1] === (body.seizo || '')) {
        sh.deleteRow(i + 1);
      }
    }
  } else if (body.type === 'koji') {
    const sh = ss.getSheetByName('工事内容');
    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === body.koji) sh.deleteRow(i + 1);
    }
  } else if (body.type === 'company') {
    const sh = ss.getSheetByName('宛先会社');
    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === body.company1 && data[i][1] === (body.company2 || '')) sh.deleteRow(i + 1);
    }
  }
  return { ok: true, data: getMasterData_() };
}

// ===================== 報告書一覧 =====================

function getReportsList_() {
  const sh = ensureReportsSheet_();
  const data = sh.getDataRange().getValues().slice(1).filter(r => r[0] !== '');

  // 既に日付型として保存されてしまった過去データを安全に文字列へ戻す
  const toDateStr = v => (v instanceof Date) ? Utilities.formatDate(v, 'JST', 'yyyy-MM-dd') : v;
  const toDateTimeStr = v => (v instanceof Date) ? Utilities.formatDate(v, 'JST', 'yyyy-MM-dd HH:mm') : v;

  const reports = data.map(r => ({
    fileId: r[0],
    spreadsheetId: r[1] || '',   // 旧バージョン互換用（新規作成分は空）
    siteName: r[2],
    workDate: toDateStr(r[3]),
    createdAt: toDateTimeStr(r[4]),
    fileName: r[5],
    dataFileId: r[6] || '',      // 旧バージョン互換用（新規作成分は空。編集データはxlsx自体に埋め込み）
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + r[0],
    spreadsheetUrl: r[1] ? ('https://docs.google.com/spreadsheets/d/' + r[1] + '/edit') : ''
  }));
  reports.reverse(); // 新しい順

  return { ok: true, reports: reports };
}

/**
 * 編集用に、保存済み報告書の元データ(現場情報＋写真)を取得する。
 * 新方式: xlsxファイル自体に埋め込まれたJSONを取り出す（Driveのファイル数を増やさないため）。
 * 旧方式: 別ファイル(dataFileId)に保存されていた場合はそちらから取得（後方互換）。
 */
function getReportData_(fileIdOrDataFileId) {
  if (!fileIdOrDataFileId) throw new Error('fileIdが指定されていません');

  // まずxlsx自体に埋め込まれたデータを試す
  try {
    const data = extractEmbeddedReportData_(fileIdOrDataFileId);
    if (data) return { ok: true, data: data };
  } catch (e) { /* 埋め込みなし、または旧形式 */ }

  // 旧形式（独立したJSONファイル）を試す
  try {
    const text = DriveApp.getFileById(fileIdOrDataFileId).getBlob().getDataAsString('UTF-8');
    return { ok: true, data: JSON.parse(text) };
  } catch (e) {
    throw new Error('編集データが見つかりません(この報告書は編集に対応していない可能性があります)');
  }
}

/**
 * body = { fileId }
 */
function deleteReport_(body) {
  if (!body.fileId) throw new Error('fileIdが指定されていません');

  const sh = ensureReportsSheet_();
  const data = sh.getDataRange().getValues();
  let spreadsheetId = null;
  let dataFileId = null;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === body.fileId) {
      spreadsheetId = data[i][1];
      dataFileId = data[i][6];
      sh.deleteRow(i + 1);
    }
  }

  try { DriveApp.getFileById(body.fileId).setTrashed(true); } catch (e) { /* 既に削除済みなど */ }
  // 以下は旧バージョンで作成された報告書の後片付け（新規作成分は該当なし）
  if (spreadsheetId) {
    try { DriveApp.getFileById(spreadsheetId).setTrashed(true); } catch (e) { /* 既に削除済みなど */ }
  }
  if (dataFileId) {
    try { DriveApp.getFileById(dataFileId).setTrashed(true); } catch (e) { /* 既に削除済みなど */ }
  }

  return getReportsList_();
}

// ===================== 報告書生成 =====================

/**
 * body = {
 *   siteName, customerSuffix('様'固定でもOK), workContent, workDate,
 *   photos: [
 *     { dataUrl, shotAt, place, kata, seizo, koji }, ...
 *   ]
 * }
 */
function generateReport_(body) {
  const ss = SpreadsheetApp.create('写真報告書_' + (body.siteName || '無題') + '_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss'));

  buildCoverSheet_(ss, body);

  const photos = body.photos || [];
  const PER_SHEET = 3; // 3枚 = 印刷1ページ分（元xlsmの改ページ位置=49行目までを1シートに）
  const sheetCount = Math.max(1, Math.ceil(photos.length / PER_SHEET));

  const placementsBySheetName = {}; // 画像をセルにぴったり合わせるための後処理用データ
  const printRangesBySheetName = { '表紙': 'A1:L33' }; // 各シート=印刷1ページに収める範囲

  for (let s = 0; s < sheetCount; s++) {
    const sheetPhotos = photos.slice(s * PER_SHEET, s * PER_SHEET + PER_SHEET);
    const sheetName = sheetCount > 1 ? '写真貼り付け原紙' + (s + 1) : '写真貼り付け原紙';
    buildPhotoSheet_(ss, sheetName, sheetPhotos, placementsBySheetName);
    printRangesBySheetName[sheetName] = 'A1:AC49';
  }

  // 既定で残る空シートSheet1を削除
  const def = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (def) ss.deleteSheet(def);

  SpreadsheetApp.flush();

  // xlsxを組み立てる（画像をセルにぴったり合わせる後処理＋印刷範囲を1ページに固定）
  let xlsxBlob = exportAsXlsx_(ss.getId());
  xlsxBlob = stretchImagesToFillCells_(xlsxBlob, placementsBySheetName);
  xlsxBlob = applyPrintSettings_(xlsxBlob, printRangesBySheetName);

  const outputFormat = (body.outputFormat === 'pdf') ? 'pdf' : 'excel';
  const outFolder = DriveApp.getFolderById(getProps_().getProperty('OUTPUT_FOLDER_ID'));

  let file;
  if (outputFormat === 'pdf') {
    const pdfBlob = convertXlsxToPdf_(xlsxBlob);
    file = outFolder.createFile(pdfBlob).setName(ss.getName() + '.pdf');
  } else {
    xlsxBlob = embedReportData_(xlsxBlob, body); // 編集用データはxlsx形式の時だけ埋め込み可能
    file = outFolder.createFile(xlsxBlob).setName(ss.getName() + '.xlsx');
  }
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 中間生成物のGoogleスプレッドシートはもう不要なので破棄（Driveにファイルが増え続けるのを防ぐ）
  try { DriveApp.getFileById(ss.getId()).setTrashed(true); } catch (e) { /* 失敗しても致命的ではない */ }

  // 報告書一覧に記録（あとで一覧表示・編集・削除できるように）。
  // xlsx1ファイルで完結する新方式のため、spreadsheetId/dataFileIdは空にしておく（列は後方互換のため維持）。
  ensureReportsSheet_().appendRow([
    file.getId(),
    '',
    body.siteName || '',
    body.workDate || '',
    Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd HH:mm'),
    file.getName(),
    ''
  ]);

  // 編集モード（既存報告書の差し替え）の場合、古いファイルを削除
  if (body.editFileId) {
    try { deleteReport_({ fileId: body.editFileId }); } catch (e) { /* 削除失敗は無視 */ }
  }

  return {
    ok: true,
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId(),
    fileId: file.getId(),
    format: outputFormat
  };
}

function buildCoverSheet_(ss, body) {
  const sheet = ss.getSheets()[0];
  sheet.setName('表紙');

  // 列幅（元xlsmのA/B/G/H/L列幅を再現。他は既定幅のまま）
  sheet.setColumnWidth(1, 32);  // A
  sheet.setColumnWidth(2, 64);  // B
  sheet.setColumnWidth(7, 23);  // G
  sheet.setColumnWidth(8, 64);  // H
  sheet.setColumnWidth(12, 34); // L

  // 行高（元xlsmに合わせる。7行目のみタイトル用に高め）
  sheet.setRowHeight(1, 18);
  for (let r = 2; r <= 33; r++) sheet.setRowHeight(r, r === 7 ? 46 : 31);

  const FONT = 'HGP創英ﾌﾟﾚｾﾞﾝｽEB';
  const NAVY = '#1c3448';
  const AMBER = '#e8a13c';

  // タイトル「完 了 写 真」: 紺地に白文字＋アンバーの太枠で見出し感を強める
  const title = sheet.getRange('E7:H7');
  title.merge();
  title.setValue('完 了 写 真')
    .setFontFamily(FONT).setFontSize(24).setFontColor('#ffffff')
    .setBackground(NAVY)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, AMBER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // タイトル下にアンバーの帯線を引いてアクセントに
  sheet.getRange('C9:K9').setBackground(AMBER);
  sheet.setRowHeight(9, 4);

  // 物件名／様（値の下に下線を引き、記入欄らしく見せる）
  sheet.getRange('C13:D13').merge().setValue('物件名：')
    .setFontFamily(FONT).setFontSize(16).setHorizontalAlignment('right');
  sheet.getRange('E13:I13').merge().setValue(body.siteName || '')
    .setFontFamily(FONT).setFontSize(16).setHorizontalAlignment('left')
    .setBorder(false, false, true, false, false, false, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('J13').setValue('様').setFontFamily(FONT).setFontSize(16);

  // 作業内容
  sheet.getRange('C15:D15').merge().setValue('作業内容：')
    .setFontFamily(FONT).setFontSize(16).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange('E15:J15').merge().setValue(body.workContent || '')
    .setFontFamily(FONT).setFontSize(16).setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, '#999999', SpreadsheetApp.BorderStyle.SOLID);

  // 作業日（年月日表記に変換して表示）
  sheet.getRange('C17:D17').merge().setValue('作業日：')
    .setFontFamily(FONT).setFontSize(16).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange('E17:J17').merge().setValue(toJapaneseDateForSheet_(body.workDate))
    .setFontFamily(FONT).setFontSize(16).setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, '#999999', SpreadsheetApp.BorderStyle.SOLID);

  // 宛先会社（2行）。上に区切り線を入れて署名欄のように見せる
  sheet.getRange('D25:I25').setBorder(false, false, true, false, false, false, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('D26:I26').merge().setValue(body.companyLine1 || '')
    .setFontFamily(FONT).setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('D27:I27').merge().setValue(body.companyLine2 || '')
    .setFontFamily(FONT).setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle');
}

/**
 * "2026-07-22" のようなISO形式の日付文字列を "2026年7月22日" に変換する。
 * すでに変換済み・空・不正な形式の場合はそのまま返す。
 */
function toJapaneseDateForSheet_(dateStr) {
  if (!dateStr) return '';
  const m = String(dateStr).match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return dateStr;
  return m[1] + '年' + Number(m[2]) + '月' + Number(m[3]) + '日';
}

/**
 * 6枠分の写真ブロックを1シートに配置。元xlsmと同じ行構成:
 *  ブロック開始行: 4, 20, 36 (A列:写真, T列:情報)。3ブロック=49行目までで
 *  印刷1ページに収まる(元xlsmの改ページ位置と同じ)。
 *  各ブロック内 T列オフセット: 0=撮影日時ラベル,1=値,2=撮影場所ラベル,3=値,
 *                              4=品番,5=製造番号,6=空白,7=工事内容ラベル,8=値
 */
function buildPhotoSheet_(ss, sheetName, photos, placementsBySheetName) {
  const sheet = ss.insertSheet(sheetName);

  // デフォルトのシートは26列(A〜Z)までしかないため、AC列(29列目)まで使えるよう列を追加
  const neededCols = 29;
  const currentCols = sheet.getMaxColumns();
  if (currentCols < neededCols) {
    sheet.insertColumnsAfter(currentCols, neededCols - currentCols);
  }

  sheet.setColumnWidth(1, 20); // A列
  for (let c = 2; c <= 29; c++) sheet.setColumnWidth(c, 26); // B〜AC 概算

  const blockStarts = [4, 20, 36]; // 元xlsmの改ページ(49行目)までの3ブロック分＝印刷1ページ
  const PHOTO_COL_START = 1;   // A
  const PHOTO_COL_END = 18;    // R
  const INFO_COL_START = 20;   // T
  const INFO_COL_END = 29;     // AC

  sheet.getRange(2, 1).setValue('◆完了写真').setFontWeight('bold').setFontSize(14);

  blockStarts.forEach((startRow, idx) => {
    const endRow = startRow + 13; // 14行分(元xlsmと同じ)

    // 写真エリア（A:R を1ブロック分結合）
    const photoRange = sheet.getRange(startRow, PHOTO_COL_START, endRow - startRow + 1, PHOTO_COL_END - PHOTO_COL_START + 1);
    photoRange.merge();
    photoRange.setBorder(true, true, true, true, false, false);

    // 情報エリア（T:AC を1行ずつ結合）
    const labelRow = startRow;                 // 撮影日時
    const shotAtRow = startRow + 1;             // 値
    const placeLabelRow = startRow + 2;         // 撮影場所
    const placeRow = startRow + 3;              // 値
    const kataRow = startRow + 4;               // 品番：
    const seizoRow = startRow + 5;              // 製造番号：
    const kojiLabelRow = startRow + 7;          // 工事内容
    const kojiRow = startRow + 8;               // 値

    [labelRow, shotAtRow, placeLabelRow, placeRow, kataRow, seizoRow, startRow + 6, kojiLabelRow, kojiRow]
      .forEach(r => {
        const rg = sheet.getRange(r, INFO_COL_START, 1, INFO_COL_END - INFO_COL_START + 1);
        rg.merge();
        rg.setFontFamily('MS PGothic').setFontSize(9).setVerticalAlignment('middle');
        rg.setBorder(true, null, true, null, false, false, '#999999', SpreadsheetApp.BorderStyle.DOTTED);
      });

    // 撮影日時の値だけ中央揃え（元xlsmと同じ）
    sheet.getRange(shotAtRow, INFO_COL_START, 1, INFO_COL_END - INFO_COL_START + 1)
      .setHorizontalAlignment('center');

    sheet.getRange(labelRow, INFO_COL_START).setValue('撮影日時');
    sheet.getRange(placeLabelRow, INFO_COL_START).setValue('撮影場所');
    sheet.getRange(kojiLabelRow, INFO_COL_START).setValue('工事内容');

    const p = photos[idx];
    if (p) {
      sheet.getRange(shotAtRow, INFO_COL_START).setValue(p.shotAt || '');
      sheet.getRange(placeRow, INFO_COL_START).setValue(p.place || '');
      sheet.getRange(kataRow, INFO_COL_START).setValue('品　　　番：' + (p.kata || ''));
      sheet.getRange(seizoRow, INFO_COL_START).setValue('製造番号：' + (p.seizo || ''));
      sheet.getRange(kojiRow, INFO_COL_START).setValue(p.koji || '');

      if (p.dataUrl) {
        insertPhotoIntoRange_(sheet, p.dataUrl, photoRange);
        if (placementsBySheetName) {
          if (!placementsBySheetName[sheetName]) placementsBySheetName[sheetName] = [];
          // OOXML(0始まり)の"to"は「境界の位置」を表すため、終端行/終端列の次の位置を指定する
          placementsBySheetName[sheetName].push({
            fromCol: PHOTO_COL_START - 1,      // 0
            fromRow: startRow - 1,
            toCol: PHOTO_COL_END,               // R列(18)の右端 = S列の開始位置
            toRow: endRow                       // 最終行の下端 = 次の行の開始位置
          });
        }
      }
    }
  });
}

/**
 * 画像(dataURL)をセル範囲にサイズを合わせて中央寄せで挿入。
 * 元のVBAマクロ(Worksheet_BeforeDoubleClick)と同じ挙動を再現。
 *
 * 注意: Googleスプレッドシートの列幅(px指定)は、xlsxへの書き出し時に
 * Excelの文字幅単位に変換されるため、変換誤差でセル幅がわずかに縮む。
 * そのため、計算上のセル幅・高さに安全マージンをかけてから画像サイズを決める。
 */
const IMAGE_FIT_MARGIN = 0.99;

function insertPhotoIntoRange_(sheet, dataUrl, targetRange) {
  const blob = dataUrlToBlob_(dataUrl);
  const image = sheet.insertImage(blob, targetRange.getColumn(), targetRange.getRow());

  // ピクセル換算でのセル範囲サイズを取得（安全マージンを掛けて少し小さめに）
  const targetWidth = getRangeWidthPx_(sheet, targetRange) * IMAGE_FIT_MARGIN;
  const targetHeight = getRangeHeightPx_(sheet, targetRange) * IMAGE_FIT_MARGIN;

  const origWidth = image.getWidth();
  const origHeight = image.getHeight();
  const scale = Math.min(targetWidth / origWidth, targetHeight / origHeight, 1);

  const newWidth = origWidth * scale;
  const newHeight = origHeight * scale;
  image.setWidth(newWidth).setHeight(newHeight);

  // 中央寄せ
  const offsetX = Math.max(0, (targetWidth - newWidth) / 2);
  const offsetY = Math.max(0, (targetHeight - newHeight) / 2);
  image.setAnchorCell(sheet.getRange(targetRange.getRow(), targetRange.getColumn()));
  image.setAnchorCellXOffset(offsetX);
  image.setAnchorCellYOffset(offsetY);
}

function getRangeWidthPx_(sheet, range) {
  let w = 0;
  for (let c = range.getColumn(); c < range.getColumn() + range.getNumColumns(); c++) {
    w += sheet.getColumnWidth(c);
  }
  return w;
}

function getRangeHeightPx_(sheet, range) {
  let h = 0;
  for (let r = range.getRow(); r < range.getRow() + range.getNumRows(); r++) {
    h += sheet.getRowHeight(r);
  }
  return h;
}

function dataUrlToBlob_(dataUrl) {
  const match = dataUrl.match(/^data:(.+?);base64,(.+)$/);
  if (!match) throw new Error('不正な画像データです');
  const contentType = match[1];
  const base64 = match[2];
  const bytes = Utilities.base64Decode(base64);
  return Utilities.newBlob(bytes, contentType, 'photo.jpg');
}

function exportAsXlsx_(spreadsheetId) {
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=xlsx';
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token }
  });
  return response.getBlob();
}

/**
 * 完成したxlsx(印刷範囲・1ページ収まる設定を反映済み)をPDFに変換する。
 * 手順: xlsxを一時的にGoogleスプレッドシートとして変換アップロード→PDFエクスポート→一時ファイルを削除。
 * xlsxに書き込んだ印刷設定(印刷範囲・fitToPage等)をGoogleが取り込んでくれるため、
 * そのままPDF化しても表紙1ページ・写真3枚ごとに1ページという体裁を維持できる。
 *
 * 事前準備: GASエディタで「サービス」→「Drive API」を追加しておく必要がある（v2・v3どちらでも動く）。
 */
function convertXlsxToPdf_(xlsxBlob) {
  const title = '_temp_pdf_' + Utilities.getUuid();
  let tempSheetId = null;

  if (typeof Drive !== 'undefined' && Drive.Files) {
    if (typeof Drive.Files.insert === 'function') {
      // Drive API v2
      const resource = { title: title, mimeType: MimeType.GOOGLE_SHEETS };
      const converted = Drive.Files.insert(resource, xlsxBlob, { convert: true });
      tempSheetId = converted.id;
    } else if (typeof Drive.Files.create === 'function') {
      // Drive API v3
      const resource = { name: title, mimeType: MimeType.GOOGLE_SHEETS };
      const converted = Drive.Files.create(resource, xlsxBlob);
      tempSheetId = converted.id;
    }
  }

  if (!tempSheetId) {
    throw new Error('PDF出力にはDrive APIサービスが必要です。GASエディタ左側の「サービス」(＋)から「Drive API」を追加してから、もう一度お試しください。');
  }

  try {
    const url = 'https://docs.google.com/spreadsheets/d/' + tempSheetId +
      '/export?format=pdf&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false';
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    return response.getBlob().setName('report.pdf');
  } finally {
    try { DriveApp.getFileById(tempSheetId).setTrashed(true); } catch (e) { /* 失敗しても致命的ではない */ }
  }
}

/**
 * xlsx書き出し後の画像は「絶対サイズ(EMU)」で埋め込まれるため、
 * Googleスプレッドシート→xlsxの列幅換算誤差でセルからはみ出す/隙間ができる。
 * これを解消するため、xlsx内部のdrawing XMLを直接書き換え、
 * 画像のアンカーを「絶対サイズ指定(oneCellAnchor)」から
 * 「セル範囲に追従する指定(twoCellAnchor)」に変換し、
 * 実際のセル幅・高さに関わらず必ずぴったり収まるようにする。
 *
 * placementsBySheetName = {
 *   'シート名': [ { fromCol, fromRow, toCol, toRow }, ... ]  // 0始まり、写真の挿入順
 * }
 */
function stretchImagesToFillCells_(blob, placementsBySheetName) {
  if (!placementsBySheetName || Object.keys(placementsBySheetName).length === 0) return blob;

  // Utilities.unzip()はBlobのContentTypeが厳密に"application/zip"であることを要求するため変換する
  const zipBlob = blob.copyBlob().setContentType('application/zip');
  const files = Utilities.unzip(zipBlob);
  const fileMap = {};
  files.forEach(f => { fileMap[f.getName()] = f; });

  const getText = name => (fileMap[name] ? fileMap[name].getDataAsString('UTF-8') : null);

  const workbookXml = getText('xl/workbook.xml');
  const workbookRelsXml = getText('xl/_rels/workbook.xml.rels');
  if (!workbookXml || !workbookRelsXml) return blob;

  // シート名 -> r:id
  const sheetNameToRid = {};
  const sheetTagRe = /<sheet\b[^>]*\/>/g;
  let sm;
  while ((sm = sheetTagRe.exec(workbookXml)) !== null) {
    const tag = sm[0];
    const nameMatch = tag.match(/name="([^"]*)"/);
    const ridMatch = tag.match(/r:id="([^"]*)"/);
    if (nameMatch && ridMatch) sheetNameToRid[nameMatch[1]] = ridMatch[1];
  }

  // r:id -> worksheets/sheetN.xml
  const ridToTarget = {};
  const relTagRe = /<Relationship\b[^>]*\/>/g;
  let rm;
  while ((rm = relTagRe.exec(workbookRelsXml)) !== null) {
    const tag = rm[0];
    const idMatch = tag.match(/Id="([^"]*)"/);
    const targetMatch = tag.match(/Target="([^"]*)"/);
    if (idMatch && targetMatch) ridToTarget[idMatch[1]] = targetMatch[1];
  }

  Object.keys(placementsBySheetName).forEach(sheetName => {
    const placements = placementsBySheetName[sheetName];
    if (!placements || placements.length === 0) return;

    const rid = sheetNameToRid[sheetName];
    if (!rid) return;
    const target = ridToTarget[rid]; // 例: "worksheets/sheet2.xml"
    if (!target) return;

    const sheetFileName = target.split('/').pop();
    const sheetRelsPath = 'xl/worksheets/_rels/' + sheetFileName + '.rels';
    const sheetRelsXml = getText(sheetRelsPath);
    if (!sheetRelsXml) return;

    let drawingTarget = null;
    const relTagRe2 = /<Relationship\b[^>]*\/>/g;
    let dm;
    while ((dm = relTagRe2.exec(sheetRelsXml)) !== null) {
      const tag = dm[0];
      if (tag.indexOf('relationships/drawing') !== -1) {
        const targetMatch = tag.match(/Target="([^"]*)"/);
        if (targetMatch) drawingTarget = targetMatch[1];
      }
    }
    if (!drawingTarget) return;

    const drawingFileName = drawingTarget.split('/').pop();
    const drawingPath = 'xl/drawings/' + drawingFileName;
    let drawingXml = getText(drawingPath);
    if (!drawingXml) return;

    let idx = 0;
    drawingXml = drawingXml.replace(
      /<xdr:oneCellAnchor>(<xdr:from>[\s\S]*?<\/xdr:from>)<xdr:ext[^>]*\/>([\s\S]*?)<\/xdr:oneCellAnchor>/g,
      function (whole, fromXml, restXml) {
        const p = placements[idx++];
        if (!p) return whole;
        const toXml = '<xdr:to><xdr:col>' + p.toCol + '</xdr:col><xdr:colOff>0</xdr:colOff>' +
          '<xdr:row>' + p.toRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>';
        // editAs="oneCell"は「セルと一緒に伸縮しない(原寸のまま移動のみ)」という意味になり、
        // Googleのレンダラーでは画像が枠にフィットせず原寸で配置されてしまう。
        // 既定(editAs省略=twoCell、セルに合わせて伸縮)にすることで、実際のセル範囲に必ずフィットする。
        return '<xdr:twoCellAnchor>' + fromXml + toXml + restXml + '</xdr:twoCellAnchor>';
      }
    );

    fileMap[drawingPath] = Utilities.newBlob(drawingXml, 'application/xml', drawingPath);
  });

  const newFiles = Object.keys(fileMap).map(name => fileMap[name]);
  const zipped = Utilities.zip(newFiles, 'report.xlsx');
  return zipped.setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

/**
 * Apps ScriptのSpreadsheetApp APIには印刷範囲・改ページ・拡大縮小印刷を
 * 設定する手段がないため、xlsx内部のXMLを直接書き換えて設定する。
 * 各シートを「印刷範囲=指定レンジ」「1ページに収まるよう自動縮小」にする。
 *
 * sheetPrintRanges = { 'シート名': 'A1:AC49', ... }  （A1形式、シート名なし）
 */
function applyPrintSettings_(blob, sheetPrintRanges) {
  if (!sheetPrintRanges || Object.keys(sheetPrintRanges).length === 0) return blob;

  const zipBlob = blob.copyBlob().setContentType('application/zip');
  const files = Utilities.unzip(zipBlob);
  const fileMap = {};
  files.forEach(f => { fileMap[f.getName()] = f; });

  const getText = name => (fileMap[name] ? fileMap[name].getDataAsString('UTF-8') : null);

  let workbookXml = getText('xl/workbook.xml');
  const workbookRelsXml = getText('xl/_rels/workbook.xml.rels');
  if (!workbookXml || !workbookRelsXml) return blob;

  // シート名 -> { rid, index(0始まりのシート順) }
  const sheetInfo = {};
  const sheetTagRe = /<sheet\b[^>]*\/>/g;
  let sm;
  let sheetIndex = 0;
  while ((sm = sheetTagRe.exec(workbookXml)) !== null) {
    const tag = sm[0];
    const nameMatch = tag.match(/name="([^"]*)"/);
    const ridMatch = tag.match(/r:id="([^"]*)"/);
    if (nameMatch && ridMatch) {
      sheetInfo[nameMatch[1]] = { rid: ridMatch[1], index: sheetIndex };
    }
    sheetIndex++;
  }

  const ridToTarget = {};
  const relTagRe = /<Relationship\b[^>]*\/>/g;
  let rm;
  while ((rm = relTagRe.exec(workbookRelsXml)) !== null) {
    const tag = rm[0];
    const idMatch = tag.match(/Id="([^"]*)"/);
    const targetMatch = tag.match(/Target="([^"]*)"/);
    if (idMatch && targetMatch) ridToTarget[idMatch[1]] = targetMatch[1];
  }

  const definedNamesList = [];

  Object.keys(sheetPrintRanges).forEach(sheetName => {
    const info = sheetInfo[sheetName];
    if (!info) return;
    const target = ridToTarget[info.rid]; // 例: "worksheets/sheet2.xml"
    if (!target) return;

    const sheetPath = 'xl/' + target;
    let sheetXml = getText(sheetPath);
    if (!sheetXml) return;

    // 1. sheetPrに fitToPage を追加（1ページに収まるよう自動縮小）
    // 注意: <sheetPr>の子要素は tabColor?, outlinePr?, pageSetUpPr? の順で並んでいる必要があるため、
    // 必ず</sheetPr>の直前(=末尾)に挿入する。先頭に挿入すると要素順序違反でExcelが破損とみなす。
    if (sheetXml.indexOf('</sheetPr>') !== -1) {
      if (sheetXml.indexOf('pageSetUpPr') === -1) {
        sheetXml = sheetXml.replace('</sheetPr>', '<pageSetUpPr fitToPage="1"/></sheetPr>');
      }
    } else if (sheetXml.indexOf('<sheetPr/>') !== -1) {
      sheetXml = sheetXml.replace('<sheetPr/>', '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
    } else {
      sheetXml = sheetXml.replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
    }

    // 2. pageMargins/pageSetup を drawing の直前（なければ</worksheet>の直前）に挿入
    const pageXml =
      '<pageMargins left="0.51" right="0.51" top="0.55" bottom="0.55" header="0.3" footer="0.3"/>' +
      '<pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="1"/>';

    if (sheetXml.indexOf('<drawing ') !== -1) {
      sheetXml = sheetXml.replace('<drawing ', pageXml + '<drawing ');
    } else {
      sheetXml = sheetXml.replace('</worksheet>', pageXml + '</worksheet>');
    }

    fileMap[sheetPath] = Utilities.newBlob(sheetXml, 'application/xml', sheetPath);

    // 3. 印刷範囲（definedName）を組み立てる
    const range = sheetPrintRanges[sheetName];
    const absRange = range.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'); // A1:L33 -> $A$1:$L$33
    definedNamesList.push(
      '<definedName name="_xlnm.Print_Area" localSheetId="' + info.index + '">\'' +
      sheetName.replace(/'/g, "''") + '\'!' + absRange +
      '</definedName>'
    );
  });

  if (definedNamesList.length > 0) {
    const definedNamesXml = '<definedNames>' + definedNamesList.join('') + '</definedNames>';
    if (workbookXml.indexOf('<definedNames/>') !== -1) {
      workbookXml = workbookXml.replace('<definedNames/>', definedNamesXml);
    } else if (workbookXml.indexOf('<definedNames>') !== -1) {
      workbookXml = workbookXml.replace(/<definedNames>[\s\S]*?<\/definedNames>/, definedNamesXml);
    } else {
      workbookXml = workbookXml.replace('</sheets>', '</sheets>' + definedNamesXml);
    }
    fileMap['xl/workbook.xml'] = Utilities.newBlob(workbookXml, 'application/xml', 'xl/workbook.xml');
  }

  const newFiles = Object.keys(fileMap).map(name => fileMap[name]);
  const zipped = Utilities.zip(newFiles, 'report.xlsx');
  return zipped.setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

/**
 * 編集時に読み込み直せるよう、送信された元データ(現場情報＋写真)をxlsxファイル自体に
 * 追加のzipエントリとして埋め込む。Excel/Google的には未参照の部品なので無視される。
 * これにより「編集用の別ファイル」をDriveに作らずに済み、Driveのファイル数が増えない。
 *
 * 注意: 拡張子は".xml"にしている（中身はJSONテキストのまま）。xlsx内の[Content_Types].xmlには
 * 既に".xml"の既定コンテンツタイプが登録済みのため、追加の登録なしでそのまま安全に同梱できる。
 * ".json"のような未登録の拡張子で追加すると、Excelが「内容に問題がある」と検出し破損警告が出る。
 */
function embedReportData_(blob, dataObj) {
  const zipBlob = blob.copyBlob().setContentType('application/zip');
  const files = Utilities.unzip(zipBlob);
  const jsonBlob = Utilities.newBlob(JSON.stringify(dataObj), 'application/xml', 'customData/reportData.xml');
  const newFiles = files.concat([jsonBlob]);
  const zipped = Utilities.zip(newFiles, 'report.xlsx');
  return zipped.setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

/**
 * embedReportData_で埋め込んだ編集用データを取り出す。埋め込みがなければnullを返す。
 */
function extractEmbeddedReportData_(fileId) {
  const blob = DriveApp.getFileById(fileId).getBlob();
  const zipBlob = blob.copyBlob().setContentType('application/zip');
  const files = Utilities.unzip(zipBlob);
  const target = files.filter(f => f.getName() === 'customData/reportData.xml' || f.getName() === 'customData/reportData.json')[0];
  if (!target) return null;
  return JSON.parse(target.getDataAsString('UTF-8'));
}

// ════════════════════════════════════════════════
// AuthGuard.gs ― Googleログインの本人確認（GAS側）  v2
//
// ■ 許可リスト
//   ・共通の「アクセス許可」スプレッドシート（シート名「許可アカウント」）を、管理画面から編集します。
//     スクリプト プロパティ ACL_SHEET_ID にそのスプレッドシートのIDを登録してください。
//   ・予備として、スクリプト プロパティ ALLOWED_EMAILS（カンマ区切り）も有効です。
//   ・管理者：シートの「権限」列が「管理者」の人、またはスクリプト プロパティ ADMIN_EMAILS の人。
//   ・許可リストが空のときは全員を拒否します。
//
// ■ 初回だけ：エディタで関数 authorizeOnce を選んで「実行」し、権限を承認してください。
//   （名前の末尾が _ の関数は、実行メニューに表示されません）
//
// ■ doGet(e) と doPost(e) の先頭に、次の2行を入れる
//   const g = authGuard_(e); if (g) return g;
// ════════════════════════════════════════════════

const GOOGLE_CLIENT_ID = '782190614995-5d4fur401smd1gpnnvs43h3m2q2q2ar3.apps.googleusercontent.com';
const ACL_SHEET_NAME = '許可アカウント';

// 初回だけ実行する関数：外部通信とスプレッドシートの権限の承認画面を出す
function authorizeOnce() {
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=test', { muteHttpExceptions: true });
  CacheService.getScriptCache().put('authorize_test', '1', 5);
  const id = PropertiesService.getScriptProperties().getProperty('ACL_SHEET_ID');
  if (id) { const sh = aclSheet_(); Logger.log('許可シート: ' + sh.getParent().getName() + ' / 登録 ' + Math.max(0, sh.getLastRow() - 1) + ' 件'); }
  Logger.log('承認OK。ACL_SHEET_ID=' + (id ? '設定あり' : '未設定（ALLOWED_EMAILSだけで動作）'));
}

// ---------- 許可リスト ----------
function listProp_(name) {
  const v = PropertiesService.getScriptProperties().getProperty(name) || '';
  return v.split(/[\s,;]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
}
function aclSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('ACL_SHEET_ID');
  if (!id) return null;
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(ACL_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ACL_SHEET_NAME);
    sh.getRange(1, 1, 1, 4).setValues([['メールアドレス', '名前', '権限', '登録日時']]);
  }
  return sh;
}
function aclRows_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('acl_rows');
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* 再取得 */ } }
  let rows = [];
  try {
    const sh = aclSheet_();
    if (sh && sh.getLastRow() > 1) {
      rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
        .map(r => ({ email: String(r[0]).trim().toLowerCase(), name: String(r[1] || ''), role: String(r[2] || '') }))
        .filter(r => r.email);
    }
  } catch (err) { rows = []; }
  cache.put('acl_rows', JSON.stringify(rows), 60);
  return rows;
}
function allowedEmails_() { return listProp_('ALLOWED_EMAILS').concat(aclRows_().map(r => r.email)); }
function adminEmails_() { return listProp_('ADMIN_EMAILS').concat(aclRows_().filter(r => r.role === '管理者').map(r => r.email)); }

// ---------- IDトークンの確認 ----------
// 戻り値：{ ok:true, email } | { ok:false, reason:'unauthorized' } | { ok:false, reason:'forbidden', email }
function verifyIdToken_(idToken) {
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
  if (info.aud !== GOOGLE_CLIENT_ID) return { ok: false, reason: 'unauthorized' };
  if (iss !== 'accounts.google.com' && iss !== 'https://accounts.google.com') return { ok: false, reason: 'unauthorized' };
  if (Number(info.exp) <= now) return { ok: false, reason: 'unauthorized' };
  if (String(info.email_verified) !== 'true') return { ok: false, reason: 'unauthorized' };
  const email = String(info.email || '').toLowerCase();
  if (allowedEmails_().indexOf(email) < 0) return { ok: false, reason: 'forbidden', email: email };
  return { ok: true, email: email };
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- 管理画面用の操作（acl_list / acl_add / acl_remove）。管理者だけが使える ----------
function aclHandle_(action, p, a) {
  if (!a.email) return jsonOut_({ status: 'error', success: false, error: a.reason, data: a.reason, message: a.reason });
  if (adminEmails_().indexOf(a.email) < 0) {
    return jsonOut_({ status: 'error', success: false, error: 'not_admin', data: 'not_admin', message: '管理者のみ操作できます' });
  }
  const props = { me: a.email };
  if (action === 'acl_list') {
    props.rows = aclRows_();
    props.propEmails = listProp_('ALLOWED_EMAILS');
    props.propAdmins = listProp_('ADMIN_EMAILS');
    props.sheetReady = !!PropertiesService.getScriptProperties().getProperty('ACL_SHEET_ID');
    return jsonOut_({ status: 'success', success: true, data: props });
  }
  const sh = aclSheet_();
  if (!sh) return jsonOut_({ status: 'error', success: false, error: 'no_sheet', data: 'no_sheet', message: 'スクリプト プロパティ ACL_SHEET_ID が未設定です' });
  const email = String(p.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return jsonOut_({ status: 'error', success: false, error: 'bad_email', data: 'bad_email', message: 'メールアドレスの形式が正しくありません' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const last = sh.getLastRow();
    const emails = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase()) : [];
    const idx = emails.indexOf(email);
    if (action === 'acl_add') {
      const name = String(p.name || '').slice(0, 60);
      const role = p.role === '管理者' ? '管理者' : '';
      if (idx >= 0) sh.getRange(idx + 2, 1, 1, 3).setValues([[email, name, role]]);
      else sh.appendRow([email, name, role, new Date()]);
    } else if (action === 'acl_remove') {
      if (email === a.email) return jsonOut_({ status: 'error', success: false, error: 'self', data: 'self', message: '自分自身は削除できません' });
      if (idx >= 0) sh.deleteRow(idx + 2);
    } else {
      return jsonOut_({ status: 'error', success: false, error: 'bad_action', data: 'bad_action', message: '不明な操作です' });
    }
  } finally { lock.releaseLock(); }
  CacheService.getScriptCache().remove('acl_rows');
  props.rows = aclRows_();
  return jsonOut_({ status: 'success', success: true, data: props });
}

// 許可されていれば null、そうでなければエラー応答（JSON）を返す
function authGuard_(e) {
  const p = (e && e.parameter) || {};
  let token = p.idToken || '';
  if (!token && e && e.postData && e.postData.contents) {
    try { token = (JSON.parse(e.postData.contents) || {}).idToken || ''; } catch (err) { /* JSON以外は無視 */ }
  }
  const a = verifyIdToken_(token);
  if (/^acl_/.test(p.action || '')) return aclHandle_(p.action, p, a);
  if (a.ok) return null;
  return jsonOut_({ status: 'error', success: false, ok: false, error: a.reason, data: a.reason, message: a.reason });
}
