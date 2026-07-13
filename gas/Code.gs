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

  Logger.log('セットアップ完了');
}

function getProps_() {
  return PropertiesService.getScriptProperties();
}

// ===================== Web API =====================

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'master';
  if (action === 'master') {
    return jsonOut_(getMasterData_());
  }
  return jsonOut_({ error: 'unknown action' });
}

function doPost(e) {
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

  for (let s = 0; s < sheetCount; s++) {
    const sheetPhotos = photos.slice(s * PER_SHEET, s * PER_SHEET + PER_SHEET);
    const sheetName = sheetCount > 1 ? '写真貼り付け原紙' + (s + 1) : '写真貼り付け原紙';
    buildPhotoSheet_(ss, sheetName, sheetPhotos);
  }

  // 既定で残る空シートSheet1を削除
  const def = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (def) ss.deleteSheet(def);

  SpreadsheetApp.flush();

  // xlsx としてDriveに保存
  const xlsxBlob = exportAsXlsx_(ss.getId());
  const outFolder = DriveApp.getFolderById(getProps_().getProperty('OUTPUT_FOLDER_ID'));
  const file = outFolder.createFile(xlsxBlob).setName(ss.getName() + '.xlsx');
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    ok: true,
    spreadsheetUrl: ss.getUrl(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId(),
    fileId: file.getId()
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

  // タイトル「完 了 写 真」
  const title = sheet.getRange('E7:H7');
  title.merge();
  title.setValue('完 了 写 真')
    .setFontFamily(FONT).setFontSize(24)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false);

  // 物件名／様
  sheet.getRange('C13:D13').merge().setValue('物件名：')
    .setFontFamily(FONT).setFontSize(16).setHorizontalAlignment('right');
  sheet.getRange('E13:I13').merge().setValue(body.siteName || '')
    .setFontFamily(FONT).setFontSize(16).setHorizontalAlignment('left');
  sheet.getRange('J13').setValue('様').setFontFamily(FONT).setFontSize(16);

  // 作業内容
  sheet.getRange('C15:D15').merge().setValue('作業内容：')
    .setFontFamily(FONT).setFontSize(16).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange('E15:J15').merge().setValue(body.workContent || '')
    .setFontFamily(FONT).setFontSize(16).setHorizontalAlignment('left').setVerticalAlignment('middle');

  // 作業日
  sheet.getRange('C17:D17').merge().setValue('作業日：')
    .setFontFamily(FONT).setFontSize(16).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange('E17:J17').merge().setValue(body.workDate || '')
    .setFontFamily(FONT).setFontSize(16).setHorizontalAlignment('left').setVerticalAlignment('middle');

  // 宛先会社（2行）
  sheet.getRange('D26:I26').merge().setValue(body.companyLine1 || '')
    .setFontFamily(FONT).setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('D27:I27').merge().setValue(body.companyLine2 || '')
    .setFontFamily(FONT).setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle');
}

/**
 * 6枠分の写真ブロックを1シートに配置。元xlsmと同じ行構成:
 *  ブロック開始行: 4, 20, 36 (A列:写真, T列:情報)。3ブロック=49行目までで
 *  印刷1ページに収まる(元xlsmの改ページ位置と同じ)。
 *  各ブロック内 T列オフセット: 0=撮影日時ラベル,1=値,2=撮影場所ラベル,3=値,
 *                              4=品番,5=製造番号,6=空白,7=工事内容ラベル,8=値
 */
function buildPhotoSheet_(ss, sheetName, photos) {
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
const IMAGE_FIT_MARGIN = 0.96;

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
