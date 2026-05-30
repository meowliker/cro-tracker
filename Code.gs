// ── Minding Art · Upsell CRO Tracker · Backend ───────────────────────────────
// Paste this into Extensions → Apps Script → Code.gs
// Deploy as Web App → Execute as Me → Anyone can access
// ─────────────────────────────────────────────────────────────────────────────

var DATA_SHEET = 'CRO_Data';
var LOG_SHEET  = 'Activity_Log';
var BACKUP_SHEET = 'CRO_Backups';
var IMAGE_FOLDER = 'CRO Tracker Images';
var SPREADSHEET_ID = '';

// GET  ?action=getData   -> returns full tracker JSON
// POST { data, action }  -> saves data, logs action
// GET  (no action)       -> serves the tracker UI as an Apps Script web app

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : '';

  if (action === 'getData') {
    var data = _read();
    return _json(data);
  }

  return HtmlService
    .createHtmlOutputFromFile('tracker')
    .setTitle('Minding Art - Upsell CRO Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    if (body.action === 'uploadImage') return _json(_uploadImage(body.image || {}));
    var result = _write(body.data, body.action || {});
    return _json(result);
  } catch(err) {
    return _json({ ok: false, error: err.toString() });
  }
}

// Called by tracker.html when deployed as a single Apps Script web app.
function getData() {
  return _read();
}

function saveData(jsonStr, actionInfo) {
  return _write(jsonStr, actionInfo || {});
}

function uploadImage(image) {
  return _uploadImage(image || {});
}

function authDriveTest() {
  var folder = _imageFolder();
  Logger.log(folder.getName());
}

// ── Internal ─────────────────────────────────────────────────────────────────

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _read() {
  try {
    var ss    = _spreadsheet();
    var sheet = ss.getSheetByName(DATA_SHEET);
    if (!sheet) { sheet = ss.insertSheet(DATA_SHEET); _initLog(ss); return _seed(); }
    var val = sheet.getRange('A1').getValue();
    return val ? _parseTrackerJson(val) : _seed();
  } catch(e) { return { ok: false, error: e.toString() }; }
}

function _write(jsonStr, actionInfo) {
  try {
    var ss    = _spreadsheet();
    _validateTrackerData(jsonStr);
    var sheet = ss.getSheetByName(DATA_SHEET) || ss.insertSheet(DATA_SHEET);
    _backupCurrentData(ss, sheet);
    sheet.getRange('A1').setValue(jsonStr);

    var logSheet = ss.getSheetByName(LOG_SHEET) || _initLog(ss);
    var user = '';
    try { user = Session.getActiveUser().getEmail(); } catch(e) {}
    logSheet.appendRow([
      new Date(), user,
      actionInfo.action   || '',
      actionInfo.funnel   || '',
      actionInfo.position || '',
      actionInfo.element  || '',
      actionInfo.value    || '',
      actionInfo.result   || ''
    ]);
    return { ok: true };
  } catch(e) { return { ok: false, error: e.toString() }; }
}

function _validateTrackerData(jsonStr) {
  var data = _parseTrackerJson(jsonStr);
  if (!data || !Array.isArray(data.slots) || !data.slots.length) {
    throw new Error('Rejected save: tracker data must contain at least one slot.');
  }
  if (!data.active || !data.slots.some(function(slot) { return slot.id === data.active; })) {
    throw new Error('Rejected save: active slot is missing from tracker data.');
  }
}

function _parseTrackerJson(jsonStr) {
  var text = String(jsonStr || '').trim();
  try { return JSON.parse(text); } catch(e) {}

  var depth = 0;
  var inString = false;
  var escapeNext = false;
  for (var i = 0; i < text.length; i++) {
    var ch = text.charAt(i);
    if (escapeNext) { escapeNext = false; continue; }
    if (ch === '\\') { escapeNext = true; continue; }
    if (ch === '"') { inString = !inString; continue; }
    if (inString) continue;
    if (ch === '{') depth++;
    if (ch === '}') {
      depth--;
      if (depth === 0) return JSON.parse(text.slice(0, i + 1));
    }
  }
  return JSON.parse(text);
}

function _backupCurrentData(ss, dataSheet) {
  var current = dataSheet.getRange('A1').getValue();
  if (!current) return;
  var backupSheet = ss.getSheetByName(BACKUP_SHEET) || _initBackup(ss);
  var lastRow = backupSheet.getLastRow();
  var lastBackup = lastRow > 1 ? backupSheet.getRange(lastRow, 2).getValue() : '';
  if (lastBackup === current) return;
  backupSheet.appendRow([new Date(), current]);
}

function _spreadsheet() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet. Open the Sheet, go to Extensions -> Apps Script, paste Code.gs there, or set SPREADSHEET_ID in Code.gs.');
  return ss;
}

function _uploadImage(image) {
  try {
    if (!image.data || !image.mimeType) throw new Error('Missing image data.');
    var bytes = Utilities.base64Decode(image.data);
    var name = image.name || ('cro-image-' + new Date().getTime());
    var blob = Utilities.newBlob(bytes, image.mimeType, name);
    var folder = _imageFolder();
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return {
      ok: true,
      image: {
        id: file.getId(),
        name: name,
        url: 'https://drive.google.com/uc?export=view&id=' + file.getId()
      }
    };
  } catch(e) { return { ok: false, error: e.toString() }; }
}

function _imageFolder() {
  var folders = DriveApp.getFoldersByName(IMAGE_FOLDER);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(IMAGE_FOLDER);
}

function _initLog(ss) {
  var s = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
  s.getRange(1,1,1,8).setValues([['Timestamp','User','Action','Funnel','Position','Element','Value','Result']]);
  s.setFrozenRows(1);
  s.getRange(1,1,1,8).setFontWeight('bold');
  return s;
}

function _initBackup(ss) {
  var s = ss.getSheetByName(BACKUP_SHEET) || ss.insertSheet(BACKUP_SHEET);
  s.getRange(1,1,1,2).setValues([['Timestamp','CRO_Data_A1']]);
  s.setFrozenRows(1);
  s.getRange(1,1,1,2).setFontWeight('bold');
  return s;
}

function _seed() {
  return {
    active: 'kls-ocu-pre',
    slots: [
      { id:'kls-ocu-pre', funnel:'KLS', platform:'OCU', position:'Pre-Purchase', elements:{
        'Price':          { winner:'$13.95', variants:[{id:'p1',value:'$13.95',rpv:2.43,cvr:17.5,result:'win',date:'2025-05-01',notes:'Avg of A+B'}] },
        'Image / GIF':    { winner:null, variants:[
          {id:'ig1',value:'Original static image (Var A)',rpv:2.23,cvr:15.66,result:'running',date:'2025-05-21',notes:'396 views'},
          {id:'ig2',value:'Yellow-cover image (Var B)',rpv:2.64,cvr:18.85,result:'running',date:'2025-05-21',notes:'419 views — leading'}
        ]},
        'Headline / Copy':{ winner:'"Your child will thank you — 96% OFF"', variants:[{id:'h1',value:'"Your child will thank you — 96% OFF"',rpv:2.43,cvr:null,result:'win',date:'2025-05-01',notes:''}] },
        'Product':        { winner:'Big Mystery Box', variants:[{id:'pr1',value:'Big Mystery Box ($13.95)',rpv:2.43,cvr:17.5,result:'win',date:'2025-05-01',notes:''}] }
      }},
      { id:'kls-as-u1', funnel:'KLS', platform:'AfterSell', position:'Upsell 1', elements:{
        'Price':          { winner:'$8.00', variants:[{id:'ap1',value:'$8.00',rpv:1.09,cvr:13.6,result:'win',date:'2025-05-14',notes:''}] },
        'Layout':         { winner:null, variants:[
          {id:'al1',value:'Layout A (wider)',rpv:1.11,cvr:13.91,result:'running',date:'2025-05-14',notes:'956 views — leading'},
          {id:'al2',value:'Layout B (alternate)',rpv:1.07,cvr:13.34,result:'running',date:'2025-05-14',notes:'937 views'}
        ]},
        'Headline / Copy':{ winner:null, variants:[] },
        'Product':        { winner:'Kids Stress Response Bundle', variants:[{id:'apd1',value:'Kids Stress Response Bundle',rpv:1.09,cvr:13.6,result:'win',date:'2025-05-14',notes:''}] }
      }},
      { id:'art-as-u1', funnel:'Art Therapy', platform:'AfterSell', position:'Upsell 1', elements:{
        'Price':          { winner:'$8.00', variants:[{id:'av1',value:'$8.00',rpv:1.44,cvr:null,result:'win',date:'2025-05-01',notes:''}] },
        'Image / GIF':    { winner:null, variants:[] },
        'Headline / Copy':{ winner:null, variants:[] },
        'Product':        { winner:'Art Therapy Bundle', variants:[{id:'atp1',value:'Art Therapy Bundle',rpv:1.44,cvr:null,result:'win',date:'2025-05-01',notes:''}] }
      }},
      { id:'art-as-d2', funnel:'Art Therapy', platform:'AfterSell', position:'Downsell 2', elements:{
        'Price':          { winner:null, variants:[] },
        'Image / GIF':    { winner:null, variants:[] },
        'Headline / Copy':{ winner:null, variants:[] },
        'Product':        { winner:'Kids Nature & Nutrition', variants:[{id:'adp1',value:'Kids Nature & Nutrition Activity Book',rpv:1.03,cvr:25.66,result:'win',date:'2025-05-01',notes:''}] }
      }},
      { id:'phonics-ocu-pre', funnel:'Phonics', platform:'OCU', position:'Pre-Purchase', elements:{
        'Price':          { winner:null, variants:[] },
        'Image / GIF':    { winner:null, variants:[] },
        'Headline / Copy':{ winner:null, variants:[] },
        'Product':        { winner:null, variants:[] }
      }}
    ]
  };
}
