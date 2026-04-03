// ============================================================
// Google Apps Script - らくらく原稿管理プラグイン バックエンド
// ============================================================
// GET:
//   ?action=ping
//   ?action=initWorkbook&fileName=&figmaPageName=&figmaUrl=
//   ?action=resetWorkbook&ssId=  → 既存ブック内シートをすべて削除し __ms_init__ のみ（上書き更新の前処理）
//   ?action=addFrameSheet&ssId=&frameSheetName=&figmaPageName=&figmaUrl=
//   ?action=addRows&ssId=&sheetName=&items=<JSON>
//   ?action=readSheetData&ssId=<ssId>  → 修正原稿読取（13列: M列ID・H列本文／12列:L列ID／11列: A列ID・H列／7列・旧形式も互換）
//   ?id=<ssId>  → 同上（後方互換。id が空のときはエラーになるため readSheetData を推奨）
// POST: insertPreview（sheetName 推奨） / resetWorkbook（JSON: { action, ssId }）※更新ボタンは POST 推奨
// ============================================================

/** 13列（A〜M）。A・E・M=30px、B〜D=150px、F〜K=表、L=空、M=プラグインID（非表示）。N列以降は削除。 */
var FRAME_SHEET_NUM_COLS = 13;
var FRAME_HEADER_ROW = 9;
var FRAME_DATA_START_ROW = 10;
var FRAME_ROW_HEIGHT_PX = 21;
var COL_REF_MOD = 8;  // H
var COL_PLUGIN_ID = 13; // M
/** 概要なしブック用の初期シート（最初の addFrameSheet で削除） */
var INIT_SHEET_NAME = '__ms_init__';
var ALT_ROW_BG_1 = '#faf6f0';
var ALT_ROW_BG_2 = '#ffffff';
var THIN_OUTLINE_COLOR = '#5f6368';
var MOD_COLUMN_BORDER_COLOR = '#b71c1c';
/** シートの使用行は 1〜200 行まで。201 行目以降のグリッドを削除 */
var SHEET_MAX_ROWS = 200;

/** タイトルから「原稿作成用」（括弧付き・単独）を除く */
function stripManuscriptPageLabel_(s) {
  var t = String(s || '');
  t = t.replace(/（\s*原稿作成用\s*）/g, '').replace(/\(\s*原稿作成用\s*\)/g, '');
  t = t.replace(/\s*原稿作成用\s*/g, '');
  return t.replace(/\n\s*\n+/g, '\n').trim();
}

/** 修正原稿列 H: データ行を1つの赤外枠のみ（表の枠線より最後に実行すること） */
function applyUnifiedModificationColumnBorder_(sheet) {
  var last = sheet.getLastRow();
  if (last < FRAME_DATA_START_ROW) return;
  var numRows = last - FRAME_DATA_START_ROW + 1;
  sheet.getRange(FRAME_DATA_START_ROW, COL_REF_MOD, numRows, 1)
    .setBorder(true, true, true, true, false, false, MOD_COLUMN_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THICK);
}

/** 列Iの条件付き書式（プラグイン用）を削除 */
function pruneColumnIConditionalRules_(sheet) {
  var kept = [];
  var rules = sheet.getConditionalFormatRules();
  for (var i = 0; i < rules.length; i++) {
    if (!ruleTouchesColumnIData_(rules[i])) kept.push(rules[i]);
  }
  sheet.setConditionalFormatRules(kept);
}

function ruleTouchesColumnIData_(rule) {
  var rgs = rule.getRanges();
  for (var j = 0; j < rgs.length; j++) {
    var r = rgs[j];
    var c0 = r.getColumn();
    var c1 = c0 + r.getNumColumns() - 1;
    var r0 = r.getRow();
    var r1 = r0 + r.getNumRows() - 1;
    if (c0 <= 9 && c1 >= 9 && r1 >= FRAME_DATA_START_ROW && r0 <= SHEET_MAX_ROWS) return true;
  }
  return false;
}

/** 修正原稿の文字数(I)が目安(J)超過時は赤文字（J は「NN字程度」または空＝目安なし） */
function applyLenExceedsTargetRed_(sheet) {
  var last = Math.min(sheet.getLastRow(), SHEET_MAX_ROWS);
  pruneColumnIConditionalRules_(sheet);
  if (last < FRAME_DATA_START_ROW) return;
  var numRows = last - FRAME_DATA_START_ROW + 1;
  var targetRng = sheet.getRange(FRAME_DATA_START_ROW, 9, numRows, 1);
  var formula =
    '=AND(INDIRECT("J"&ROW())<>"\u2014",INDIRECT("J"&ROW())<>"",LEN(INDIRECT("H"&ROW()))>0,LEN(INDIRECT("H"&ROW()))>IFERROR(VALUE(REGEXEXTRACT(INDIRECT("J"&ROW()),"^[0-9]+")),-1))';
  var newRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setRanges([targetRng])
    .setFontColor(MOD_COLUMN_BORDER_COLOR)
    .build();
  var kept = sheet.getConditionalFormatRules();
  kept.push(newRule);
  sheet.setConditionalFormatRules(kept);
}

/** D3:K3 の URL 表示（下線・塗りなし） */
function setUrlRowRich_(sheet, text) {
  var len = String(text).length;
  var st = SpreadsheetApp.newTextStyle()
    .setUnderline(true)
    .setFontSize(10)
    .setForegroundColor('#1155cc')
    .build();
  var rich = SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setTextStyle(0, len, st)
    .build();
  sheet.getRange(3, 4, 1, 8).merge()
    .setRichTextValue(rich)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setBackground('#ffffff');
}

/** 記入ルール〜タイトル周り B2:K7 を薄い外枠のみ */
function applyStaticBlockOutlineB2K7_(sheet) {
  sheet.getRange(2, 2, 6, 10).setBorder(
    true, true, true, true, false, false,
    THIN_OUTLINE_COLOR, SpreadsheetApp.BorderStyle.SOLID
  );
}

/** 表ヘッダー〜最終データ行 F9:K を薄い外枠のみ */
function refreshDataTableOutlineF9K_(sheet) {
  var last = sheet.getLastRow();
  if (last < FRAME_HEADER_ROW) return;
  var numRows = last - FRAME_HEADER_ROW + 1;
  sheet.getRange(FRAME_HEADER_ROW, 6, numRows, 6).setBorder(
    true, true, true, true, false, false,
    THIN_OUTLINE_COLOR, SpreadsheetApp.BorderStyle.SOLID
  );
}

/** B:D 列の行9からの縦結合をいったん解除（データ書き込み前に呼ぶ） */
function breakApartBthroughDFromRow9_(sheet) {
  var last = Math.max(sheet.getLastRow(), FRAME_HEADER_ROW);
  var numRows = last - FRAME_HEADER_ROW + 1;
  try {
    sheet.getRange(FRAME_HEADER_ROW, 2, numRows, 3).breakApart();
  } catch (eBr) { /* 未結合など */ }
}

/** B9〜最終行×D列を1セルに縦結合（左スペーサー） */
function mergeBthroughDFromRow9ToLast_(sheet) {
  var last = sheet.getLastRow();
  if (last < FRAME_HEADER_ROW) return;
  var numRows = last - FRAME_HEADER_ROW + 1;
  var r = sheet.getRange(FRAME_HEADER_ROW, 2, numRows, 3);
  try {
    r.breakApart();
  } catch (eM) { /* 未結合 */ }
  r.merge();
  r.setBackground('#cccccc')
   .setValue('ここにページの書き出し画像を配置')
   .setFontColor('#ffffff')
   .setFontWeight('bold')
   .setHorizontalAlignment('center')
   .setVerticalAlignment('middle');
}

/** 201行目〜シート末尾まで削除（1〜200行だけ残す） */
function trimSheetRowsFrom200_(sheet) {
  var maxR = sheet.getMaxRows();
  if (maxR <= SHEET_MAX_ROWS) return;
  sheet.deleteRows(SHEET_MAX_ROWS + 1, maxR - SHEET_MAX_ROWS);
}

/** N列（14列目）以降を削除し、グリッドを A〜M のみにする */
function trimColumnsBeyondM_(sheet) {
  var maxC = sheet.getMaxColumns();
  if (maxC <= FRAME_SHEET_NUM_COLS) return;
  sheet.deleteColumns(FRAME_SHEET_NUM_COLS + 1, maxC - FRAME_SHEET_NUM_COLS);
}

function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) {
      return jsonResponse({ error: 'POST body が空です' });
    }
    var data = JSON.parse(e.postData.contents);
    if (data.action === 'insertPreview') return insertPreviewImage(data);
    if (data.action === 'resetWorkbook') return resetWorkbook(data);
    return jsonResponse({ error: '不明な action: ' + data.action });
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

function insertPreviewImage(data) {
  var ssId      = data.ssId;
  var sheetKey  = data.sheetName || data.pageName || '';
  var b64       = data.imageBase64 || '';
  if (!ssId || !b64) return jsonResponse({ error: 'ssId と imageBase64 が必要です' });

  var ss    = SpreadsheetApp.openById(ssId);
  var sheet = sheetKey ? ss.getSheetByName(sanitizeSheetName(sheetKey)) : null;
  if (!sheet) {
    var all = ss.getSheets();
    for (var si = 0; si < all.length; si++) {
      var nm = all[si].getName();
      if (nm !== INIT_SHEET_NAME && nm !== '概要' && nm.indexOf('概要_') !== 0) {
        sheet = all[si];
        break;
      }
    }
    if (!sheet && all.length > 1) sheet = all[1];
    if (!sheet && all.length > 0) sheet = all[0];
  }
  if (!sheet) return jsonResponse({ error: 'シートが見つかりません' });

  var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/png', 'figma-preview.png');

  try {
    var imgs = sheet.getImages();
    if (imgs && imgs.length) {
      for (var j = 0; j < imgs.length; j++) {
        try {
          var cell = imgs[j].getAnchorCell();
          if (cell && cell.getRow() === 1) imgs[j].remove();
        } catch (x) { /* ignore */ }
      }
    }
  } catch (x) { /* ignore */ }

  sheet.insertImage(blob, 2, 1);
  sheet.setRowHeight(1, FRAME_ROW_HEIGHT_PX);

  return jsonResponse({ ok: true, message: 'プレビュー画像を挿入しました' });
}

/** e.parameter が空でも queryString から拾う（Webアプリ経由で action が落ちる事例の保険） */
function getParamFromEvent_(e, key) {
  var p = e.parameter[key];
  if (p != null && String(p).trim() !== '') return String(p).trim();
  var qs = e.queryString || '';
  if (!qs) return '';
  var esc = String(key).replace(/[\\^$.*+?()[\]{}|]/g, '\\$&');
  var re = new RegExp('(?:^|&)' + esc + '=([^&]*)');
  var m = re.exec(qs);
  if (!m) return '';
  try {
    return decodeURIComponent(m[1].replace(/\+/g, ' ')).trim();
  } catch (x) {
    return String(m[1]).trim();
  }
}

function doGet(e) {
  try {
    var action = getParamFromEvent_(e, 'action');

    if (action === 'ping')           return jsonResponse({ ok: true, message: 'GAS接続OK' });
    if (action === 'readSheetData') {
      var sidRead = e.parameter.ssId || e.parameter.id;
      if (sidRead != null && String(sidRead).trim() !== '') {
        return readSpreadsheet(String(sidRead).trim());
      }
      return jsonResponse({ error: 'readSheetData には ssId（スプレッドシート ID）が必要です' });
    }
    if (action === 'initWorkbook')   return initWorkbook(e.parameter);
    if (action === 'resetWorkbook')  return resetWorkbook(e.parameter);
    if (action === 'addFrameSheet')  return addFrameSheet(e.parameter);
    if (action === 'addRows')        return addRows(e.parameter);
    if (action === 'initSheet')      return initSheet(e.parameter);
    if (action === 'addSection')     return addSection(e.parameter);
    var idLegacy = e.parameter.id;
    if (idLegacy != null && String(idLegacy).trim() !== '') {
      return readSpreadsheet(String(idLegacy).trim());
    }

    return jsonResponse({ error: 'actionまたはidパラメータが必要です' });
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

// ---- 新規ブック（概要シートなし。最初のフレーム追加時にこのプレースホルダーを削除） ----
function initWorkbook(params) {
  var fileName = params.fileName || ('原稿シート_' + formatDate(new Date()));

  var ss = SpreadsheetApp.create(fileName);
  var sheet = ss.getSheets()[0];
  sheet.setName(INIT_SHEET_NAME);

  return jsonResponse({ id: ss.getId(), url: ss.getUrl(), coverSheetName: '' });
}

/** 既存スプレッドシートを空にし、新規作成と同じ初期状態（プレースホルダーシートのみ）にする */
function resetWorkbook(params) {
  var ssId = String(params.ssId || params.id || '').trim();
  if (!ssId) return jsonResponse({ error: 'ssId が必要です' });
  var ss = SpreadsheetApp.openById(ssId);
  var tmp = ss.insertSheet('__ms_wipe__', 0);
  var tmpId = tmp.getSheetId();
  while (ss.getSheets().length > 1) {
    var sheets = ss.getSheets();
    var deleted = false;
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() !== tmpId) {
        ss.deleteSheet(sheets[i]);
        deleted = true;
        break;
      }
    }
    if (!deleted) break;
  }
  tmp.setName(INIT_SHEET_NAME);
  return jsonResponse({ ok: true, id: ss.getId(), url: ss.getUrl() });
}

// ---- フレーム＝1シート（表紙風ヘッダー＋一覧テーブル） ----
function addFrameSheet(params) {
  var ss            = SpreadsheetApp.openById(params.ssId);
  var frameTitle    = params.frameSheetName || params.sectionName || 'ブロック';
  var figmaPageName = params.figmaPageName || '';
  var figmaUrl      = params.figmaUrl || '';

  var initSh = ss.getSheetByName(INIT_SHEET_NAME);
  var name = ensureUniqueSheetName(ss, frameTitle);
  var sheet = ss.insertSheet(name);

  if (initSh && ss.getSheets().length > 1) {
    ss.deleteSheet(initSh);
  }

  applyFrameSheetTemplate(sheet, frameTitle, figmaPageName, figmaUrl);

  return jsonResponse({ ok: true, sheetName: name });
}

/** 記入ルール用リッチテキスト（先頭「n. 」太字・本文赤 #ea4335／参照シート準拠） */
function buildRuleRichText_(text, prefixLen) {
  var body = SpreadsheetApp.newTextStyle()
    .setBold(false).setFontSize(10).setForegroundColor('#ea4335').build();
  var head = SpreadsheetApp.newTextStyle()
    .setBold(true).setFontSize(10).setForegroundColor('#ea4335').build();
  var len = text.length;
  var pl = Math.min(prefixLen, len);
  return SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setTextStyle(0, len, body)
    .setTextStyle(0, pl, head)
    .build();
}

function applyFrameSheetTemplate(sheet, frameTitle, figmaPageName, figmaUrl) {
  trimColumnsBeyondM_(sheet);
  // A30 B150 C150 D150 E30 | F60 G400 H400 I100 J100 K備考400 L30 M30(ID)
  var W = [30, 150, 150, 150, 30, 60, 400, 400, 100, 100, 400, 30, 30];
  for (var wi = 0; wi < W.length; wi++) sheet.setColumnWidth(wi + 1, W[wi]);

  var ft = stripManuscriptPageLabel_(frameTitle);
  var fp = stripManuscriptPageLabel_(figmaPageName);
  var title2 = ft + (fp ? '\n' + fp : '');

  // 行1: プレビュー帯 B1:K1（A1 は塗りなし）
  sheet.setRowHeight(1, FRAME_ROW_HEIGHT_PX);
  sheet.getRange(1, 1, 1, 1).setBackground('#ffffff');
  sheet.getRange(1, 2, 1, 10).merge()
    .setValue('')
    .setBackground('#ffffff')
    .setVerticalAlignment('middle').setWrap(true);

  // 行2: ページタイトル（中央）
  sheet.setRowHeight(2, FRAME_ROW_HEIGHT_PX);
  sheet.getRange(2, 2, 1, 10).merge()
    .setValue(title2)
    .setFontSize(12).setFontWeight('bold')
    .setBackground('#434343').setFontColor('#ffffff')
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);

  // 行3: デザインプレビュー（中央）+ URL
  sheet.setRowHeight(3, FRAME_ROW_HEIGHT_PX);
  sheet.getRange(3, 2, 1, 2).merge()
    .setValue('デザインプレビュー')
    .setFontWeight('bold').setFontSize(10).setFontColor('#000000')
    .setBackground('#f3f3f3')
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
  var urlPlaceholder = 'ここにFigmaのリンクまたはPDFのURLがある場合は記載';
  if (figmaUrl) {
    setUrlRowRich_(sheet, figmaUrl);
  } else {
    setUrlRowRich_(sheet, urlPlaceholder);
  }

  // 行4〜7: 記入ルール
  sheet.getRange(4, 2, 4, 2).merge()
    .setValue('記入ルール')
    .setFontWeight('bold').setFontSize(10).setFontColor('#000000')
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

  var rules = [
    '1. 参考原稿のままでよい箇所は「@」と記載してください。',
    '2. 文字数は目安です（±20%目安）。無理に合わせなくて構いません。',
    '3. デザインプレビューが画面に合わないときは、表示メニューの「ズーム」で調整してください。',
    '4. 赤枠の「修正原稿」欄に確定原稿を記載してください。',
  ];
  var prefixLen = 3;
  for (var ri = 0; ri < rules.length; ri++) {
    sheet.setRowHeight(4 + ri, FRAME_ROW_HEIGHT_PX);
    sheet.getRange(4 + ri, 4, 1, 8).merge()
      .setRichTextValue(buildRuleRichText_(rules[ri], prefixLen))
      .setWrap(true).setVerticalAlignment('middle').setBackground('#ffffff');
  }

  sheet.setRowHeight(8, FRAME_ROW_HEIGHT_PX);

  var hRow = FRAME_HEADER_ROW;
  sheet.setRowHeight(hRow, FRAME_ROW_HEIGHT_PX);
  sheet.getRange(hRow, 5, 1, 1).setValue('').setBackground('#ffffff').setVerticalAlignment('middle');
  sheet.getRange(hRow, 6, 1, 6).setValues([[
    '該当箇所', '参考原稿', '修正原稿', '修正原稿の\n文字数', '目安文字数', '備考',
  ]]);
  sheet.getRange(hRow, 6, 1, 6)
    .setBackground('#f3f3f3').setFontColor('#000000')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

  applyStaticBlockOutlineB2K7_(sheet);
  trimSheetRowsFrom200_(sheet);
  mergeBthroughDFromRow9ToLast_(sheet);
  trimColumnsBeyondM_(sheet);
  refreshDataTableOutlineF9K_(sheet);
  applyLenExceedsTargetRed_(sheet);
  applyUnifiedModificationColumnBorder_(sheet);

  try {
    sheet.hideColumns(COL_PLUGIN_ID, 1);
  } catch (hideErr) { /* 古いランタイム等 */ }
}

/**
 * J列: r=ROUND(LEN(G)*1.1,0) を帯にマップ（指定の重なりは解消済み）。
 * 0〜10→15字, 20〜30→25字, … のあいだ（11〜19等）は隣の帯に合わせる。
 * r>=300 は空（目安なし）。G 空は —。
 */
function formulaTargetCharHintFromG_(rowNum) {
  var g = 'G' + rowNum;
  return (
    '=IF(LEN(' + g + ')=0,"\u2014",LET(r,ROUND(LEN(' + g + ')*1.1,0),IF(r>=300,"",IF(AND(r>=200,r<300),"300\u5b57\u7a0b\u5ea6",IF(AND(r>=150,r<200),"200\u5b57\u7a0b\u5ea6",IF(AND(r>=120,r<150),"150\u5b57\u7a0b\u5ea6",IF(AND(r>=90,r<120),"100\u5b57\u7a0b\u5ea6",IF(AND(r>=60,r<90),"80\u5b57\u7a0b\u5ea6",IF(AND(r>=40,r<60),"50\u5b57\u7a0b\u5ea6",IF(AND(r>=20,r<40),"25\u5b57\u7a0b\u5ea6",IF(AND(r>=0,r<20),"15\u5b57\u7a0b\u5ea6","")))))))))))'
  );
}

// ============================================================
// データ行（sheetName で対象シートを指定）
// ============================================================
function addRows(params) {
  var ss        = SpreadsheetApp.openById(params.ssId);
  var sheetName = sanitizeSheetName(params.sheetName || params.pageName || '');
  var sheet     = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse({ error: 'シートが見つかりません: ' + sheetName });

  var items = JSON.parse(params.items || '[]');
  if (items.length === 0) return jsonResponse({ ok: true, written: 0 });

  trimColumnsBeyondM_(sheet);
  breakApartBthroughDFromRow9_(sheet);

  var startRow = Math.max(sheet.getLastRow() + 1, FRAME_DATA_START_ROW);
  var rows     = [];

  for (var i = 0; i < items.length; i++) {
    var item   = items[i];
    var rowNum = startRow + i;
    rows.push([
      '',
      '', '', '', '',
      String(item.seqNum != null ? item.seqNum : ''),
      item.refText,
      '',
      '=LEN(H' + rowNum + ')',
      formulaTargetCharHintFromG_(rowNum),
      '',
      '',
      item.pluginId,
    ]);
  }

  sheet.getRange(startRow, 1, rows.length, FRAME_SHEET_NUM_COLS).setValues(rows);

  for (var r = 0; r < rows.length; r++) {
    var rr = startRow + r;
    sheet.setRowHeight(rr, FRAME_ROW_HEIGHT_PX);
    var bg = ((rr - FRAME_DATA_START_ROW) % 2 === 0) ? ALT_ROW_BG_1 : ALT_ROW_BG_2;
    sheet.getRange(rr, 1, 1, 1)
      .setBackground('#ffffff').setVerticalAlignment('middle').setWrap(true);
    sheet.getRange(rr, 5, 1, 1)
      .setBackground('#ffffff').setVerticalAlignment('middle').setWrap(true);
    sheet.getRange(rr, 6, 1, 6)
      .setBackground(bg).setVerticalAlignment('middle').setWrap(true);
    sheet.getRange(rr, 6, 1, 1)
      .setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(rr, 9, 1, 2)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(rr, 13, 1, 1)
      .setBackground(bg).setFontSize(8).setFontColor('#999999')
      .setVerticalAlignment('middle').setWrap(true);
  }

  trimSheetRowsFrom200_(sheet);
  mergeBthroughDFromRow9ToLast_(sheet);
  trimColumnsBeyondM_(sheet);
  refreshDataTableOutlineF9K_(sheet);
  applyLenExceedsTargetRed_(sheet);
  applyUnifiedModificationColumnBorder_(sheet);

  try {
    sheet.hideColumns(COL_PLUGIN_ID, 1);
  } catch (hideErr2) { /* ignore */ }

  return jsonResponse({ ok: true, written: rows.length });
}

// ---- 旧API（互換） ----
function initSheet(params) {
  return initWorkbook({
    fileName:       params.fileName,
    figmaPageName:  params.pageName,
    figmaUrl:       params.figmaUrl,
  });
}

function addSection(params) {
  return addFrameSheet({
    ssId:            params.ssId,
    frameSheetName:  params.sectionName,
    figmaPageName:   params.pageName,
    figmaUrl:        '',
  });
}

function readSpreadsheet(ssId) {
  var ss     = SpreadsheetApp.openById(ssId);
  var result = {};
  var sheets = ss.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    var sheet   = sheets[i];
    if (sheet.getName() === INIT_SHEET_NAME) continue;
    var lastRow = sheet.getLastRow();
    if (lastRow < FRAME_DATA_START_ROW) continue;

    var values = sheet.getRange(1, 1, lastRow, FRAME_SHEET_NUM_COLS).getValues();
    for (var j = 0; j < values.length; j++) {
      var row = values[j];
      var pluginIdM = String(row[12] != null ? row[12] : '').trim();
      if (/^[A-Z0-9]{6}$/.test(pluginIdM)) {
        result[pluginIdM] = String(row[7] != null ? row[7] : '').trim();
        continue;
      }
      var pluginIdL = String(row[11] != null ? row[11] : '').trim();
      if (/^[A-Z0-9]{6}$/.test(pluginIdL)) {
        result[pluginIdL] = String(row[7] != null ? row[7] : '').trim();
        continue;
      }
      var pluginId = String(row[0] != null ? row[0] : '').trim();
      if (/^[A-Z0-9]{6}$/.test(pluginId)) {
        var colB = String(row[1] != null ? row[1] : '').trim();
        var modified = (colB === '')
          ? String(row[7] != null ? row[7] : '').trim()
          : String(row[3] != null ? row[3] : '').trim();
        result[pluginId] = modified;
        continue;
      }
      pluginId = String(row[6] != null ? row[6] : '').trim();
      var modifiedLegacy = String(row[2] != null ? row[2] : '').trim();
      if (/^[A-Z0-9]{6}$/.test(pluginId)) {
        result[pluginId] = modifiedLegacy;
      }
    }
  }

  return jsonResponse(result);
}

function jsonResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function sanitizeSheetName(name) {
  return String(name).replace(/[:\\/?\*\[\]]/g, '_').slice(0, 100);
}

function ensureUniqueSheetName(ss, base) {
  var n = sanitizeSheetName(base).slice(0, 90);
  if (!n.length) n = 'シート';
  if (n === INIT_SHEET_NAME) n = 'ブロック';
  if (!ss.getSheetByName(n)) return n;
  for (var i = 2; i < 99; i++) {
    var cand = n.slice(0, 88) + '_' + i;
    if (!ss.getSheetByName(cand)) return cand;
  }
  return n + '_' + String(Date.now()).slice(-6);
}

function formatDate(d) {
  return d.getFullYear() + '-'
    + String(d.getMonth() + 1).padStart(2, '0') + '-'
    + String(d.getDate()).padStart(2, '0');
}

function debug_ping() {
  var e = { parameter: { action: 'ping' } };
  Logger.log(doGet(e).getContent());
}
