// ── Minding Art · Upsell CRO Tracker · Backend ───────────────────────────────
// Paste this into Extensions → Apps Script → Code.gs
// Deploy as Web App → Execute as Me → Anyone can access
// ─────────────────────────────────────────────────────────────────────────────

var DATA_SHEET = 'CRO_Data';
var LOG_SHEET  = 'Activity_Log';
var BACKUP_SHEET = 'CRO_Backups';
var STORE_INDEX_SHEET = 'CRO_Stores';
var DEFAULT_STORE_ID = 'minding-art';
var DEFAULT_STORE_NAME = 'Minding Art';
var IMAGE_FOLDER = 'CRO Tracker Images';
var SPREADSHEET_ID = '';

// GET  ?action=getData   -> returns full tracker JSON
// POST { data, action }  -> saves data, logs action
// GET  (no action)       -> serves the tracker UI as an Apps Script web app

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : '';
  var storeId = e && e.parameter ? e.parameter.storeId : '';

  if (action === 'getData') {
    var data = _read(storeId);
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
    if (body.action === 'createStore') return _json(_createStore(body.name || 'New Store'));
    if (body.action === 'deleteStore') return _json(_deleteStore(body.storeId || ''));
    var result = _write(body.data, body.action || {}, body.storeId || '');
    return _json(result);
  } catch(err) {
    return _json({ ok: false, error: err.toString() });
  }
}

// Called by tracker.html when deployed as a single Apps Script web app.
function getData(storeId) {
  return _read(storeId || '');
}

function saveData(jsonStr, actionInfo, storeId) {
  return _write(jsonStr, actionInfo || {}, storeId || '');
}

function createStore(name) {
  return _createStore(name || 'New Store');
}

function deleteStore(storeId) {
  return _deleteStore(storeId || '');
}

function uploadImage(image) {
  return _uploadImage(image || {});
}

function authSheetsTest() {
  var ss = _spreadsheet();
  Logger.log(ss.getName());
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

function _read(storeId) {
  try {
    var ss    = _spreadsheet();
    var store = _storeFor(ss, storeId);
    var sheet = ss.getSheetByName(store.sheetName);
    if (!sheet) { sheet = ss.insertSheet(store.sheetName); _initLog(ss); return _storePayload(ss, store, _defaultTrackerForStore(store)); }
    var val = sheet.getRange('A1').getValue();
    return _storePayload(ss, store, val ? _parseTrackerJson(val) : _defaultTrackerForStore(store));
  } catch(e) { return { ok: false, error: e.toString() }; }
}

function _write(jsonStr, actionInfo, storeId) {
  try {
    var ss    = _spreadsheet();
    var store = _storeFor(ss, storeId);
    _validateTrackerData(jsonStr);
    var sheet = ss.getSheetByName(store.sheetName) || ss.insertSheet(store.sheetName);
    _backupCurrentData(ss, sheet);
    sheet.getRange('A1').setValue(jsonStr);

    var logSheet = ss.getSheetByName(LOG_SHEET) || _initLog(ss);
    var user = '';
    try { user = Session.getActiveUser().getEmail(); } catch(e) {}
    logSheet.appendRow([
      new Date(), user,
      actionInfo.action   || '',
      actionInfo.store    || store.name || '',
      actionInfo.position || '',
      actionInfo.element  || '',
      actionInfo.value    || '',
      actionInfo.result   || ''
    ]);
    return { ok: true };
  } catch(e) { return { ok: false, error: e.toString() }; }
}

function _storePayload(ss, activeStore, tracker) {
  return {
    stores: _stores(ss),
    activeStoreId: activeStore.id,
    tracker: tracker
  };
}

function _stores(ss) {
  var sheet = ss.getSheetByName(STORE_INDEX_SHEET) || _initStores(ss);
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return _initStores(ss).getDataRange().getValues().slice(1).map(_storeRow);
  return values.slice(1).filter(function(row) { return row[0] && row[1] && row[2]; }).map(_storeRow);
}

function _storeRow(row) {
  return { id:String(row[0]), name:String(row[1]), sheetName:String(row[2]) };
}

function _storeFor(ss, storeId) {
  var stores = _stores(ss);
  var found = stores.filter(function(store) { return store.id === storeId; })[0];
  return found || stores[0];
}

function _initStores(ss) {
  var sheet = ss.getSheetByName(STORE_INDEX_SHEET) || ss.insertSheet(STORE_INDEX_SHEET);
  sheet.clear();
  sheet.getRange(1,1,1,4).setValues([['Store ID','Store Name','Data Sheet','Created At']]);
  sheet.getRange(2,1,1,4).setValues([[DEFAULT_STORE_ID, DEFAULT_STORE_NAME, DATA_SHEET, new Date()]]);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,4).setFontWeight('bold');
  if (!ss.getSheetByName(DATA_SHEET)) ss.insertSheet(DATA_SHEET);
  return sheet;
}

function _createStore(name) {
  try {
    var ss = _spreadsheet();
    var cleanName = String(name || '').trim();
    if (!cleanName) throw new Error('Store name is required.');
    var stores = _stores(ss);
    var id = _uniqueStoreId(stores, _slug(cleanName));
    var sheetName = _uniqueSheetName(ss, 'Store - ' + cleanName);
    var sheet = ss.insertSheet(sheetName);
    sheet.getRange('A1').setValue(JSON.stringify(_emptyProductTracker()));
    var index = ss.getSheetByName(STORE_INDEX_SHEET) || _initStores(ss);
    index.appendRow([id, cleanName, sheetName, new Date()]);
    return _read(id);
  } catch(e) { return { ok:false, error:e.toString() }; }
}

function _deleteStore(storeId) {
  try {
    var ss = _spreadsheet();
    var stores = _stores(ss);
    if (stores.length <= 1) throw new Error('Cannot delete the only store.');
    var target = stores.filter(function(store) { return store.id === storeId; })[0];
    if (!target) throw new Error('Store not found.');

    var dataSheet = ss.getSheetByName(target.sheetName);
    if (dataSheet && ss.getSheets().length > 1) ss.deleteSheet(dataSheet);

    var index = ss.getSheetByName(STORE_INDEX_SHEET) || _initStores(ss);
    var values = index.getDataRange().getValues();
    for (var i = values.length - 1; i >= 1; i--) {
      if (String(values[i][0]) === target.id) index.deleteRow(i + 1);
    }

    var remaining = _stores(ss);
    return _read(remaining[0].id);
  } catch(e) { return { ok:false, error:e.toString() }; }
}

function _slug(name) {
  return String(name || 'store').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-+|-+$/g,'') || 'store';
}

function _uniqueStoreId(stores, base) {
  var id = base;
  var n = 2;
  while (stores.some(function(store) { return store.id === id; })) id = base + '-' + n++;
  return id;
}

function _uniqueSheetName(ss, base) {
  var clean = String(base || 'Store').replace(/[\\\/\?\*\[\]:]/g,' ').slice(0, 90).trim() || 'Store';
  var name = clean;
  var n = 2;
  while (ss.getSheetByName(name)) name = (clean.slice(0, 85) + ' ' + n++);
  return name;
}

function _validateTrackerData(jsonStr) {
  var data = _parseTrackerJson(jsonStr);
  if (data && Array.isArray(data.products)) {
    if (data.activeProductId && !data.products.some(function(product) { return product.id === data.activeProductId; })) {
      throw new Error('Rejected save: active product is missing from tracker data.');
    }
    return;
  }
  if (!data || !Array.isArray(data.slots) || !data.slots.length) {
    throw new Error('Rejected save: tracker data must contain at least one slot.');
  }
  if (!data.active || !data.slots.some(function(slot) { return slot.id === data.active; })) {
    throw new Error('Rejected save: active slot is missing from tracker data.');
  }
}

function _defaultTrackerForStore(store) {
  return store && store.id === DEFAULT_STORE_ID ? _seed() : _emptyProductTracker();
}

function _emptyProductTracker() {
  return {
    schemaVersion: 2,
    activeProductId: '',
    activePlatform: 'OCU',
    activePosition: 'Pre-Purchase',
    products: []
  };
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
        url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200',
        previewUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view'
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
  s.getRange(1,1,1,8).setValues([['Timestamp','User','Action','Store','Position','Element','Value','Result']]);
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
