/**
 * PP FLOOR — Apps Script backend for the PrintPro floor tablets.
 * Bind to the "PP Floor" spreadsheet: Extensions → Apps Script.
 *
 * ---------------------------------------------------------------------------
 * WHY THIS EXISTS IN THIS SHAPE
 *
 * The tablets are hosted on GitHub Pages (https). They cannot reach the
 * pipeline at http://<shop-ip>:8002 — a browser blocks an https page from
 * fetching http, and no CORS header fixes that. So the pipeline PUSHES
 * instead of the tablet pulling: when the art room approves a rotation, the
 * Flask app POSTs it here (outbound https from the office machine, the same
 * move the PORT disagree forwarder already makes), and the tablet reads it
 * back from here over https.
 *
 * That inverts the dependency in the shop's favour. The tablet no longer
 * needs shop wifi, the office machine no longer needs to be awake for the
 * press to see its rotation, and the floor only ever sees rotations the art
 * room actually released.
 *
 * Deploy → New deployment → Web app → Execute as: Me · Who has access:
 * Anyone. "Anyone" is what lets the tablet and the pipeline POST without
 * OAuth; SHARED_TOKEN is what gates writes. The sheet stays private.
 *
 * ---------------------------------------------------------------------------
 * ROUTES
 *   POST {source:"release"}   → RELEASES tab, upserted on graphic|location
 *   POST {source:"press"}     → PRESS tab
 *   POST {source:"darkroom"}  → DARKROOM tab
 *   GET  ?fn=rotation&graphic=…       → the released rotation(s) for that job
 *   GET  ?fn=kpi&operator=…&station=… → the operator's day, the press rate
 *
 * COLUMN-MAPPING DISCIPLINE. No row is ever built by hand. COLUMNS is the
 * header and buildRow() walks it, pulling each named field off the payload.
 * Add a field by putting its name in COLUMNS where you want it — nothing can
 * silently shift a column, which is the failure that quietly corrupts every
 * report downstream.
 */

// ---------------------------------------------------------------- knobs
var SHARED_TOKEN = 'pp-floor-dzQCyQbIUaPF2uZfOlkezE_z';  // matches CFG.SHEET_TOKEN in pp_floor.html
var SETUP_CEILING_MIN_PER_SCREEN = 25;   // runaway threshold — EXCLUDED, never clipped
var RELEASE_KEEP_DAYS = 45;              // releases older than this are pruned on write
var TZ = 'America/New_York';

var PRESS_TAB    = 'PRESS';
var DARK_TAB     = 'DARKROOM';
var RELEASE_TAB  = 'RELEASES';

/* One job on the press = ONE row. The rotation the operator was shown,
   whether they ran it, and the setup minutes it cost sit on the same row —
   that join is what answers "does the released rotation cost more or less
   setup time than the one the floor picks?" */
var PRESS_COLUMNS = [
  'received','clientTs','station','operator',
  'order','colors','pieces','garment',
  'setupMin','runMin','waitMin','setupMinPerScreen',
  'setupRunaway','runRunaway','waitRunaway',
  'activities','delays','delayOther','problems','problemOther','notes',
  'portGraphic','portLocation','portApproved','portDecision','portReason',
  'portFlashes','portPredicted','portRan'
];

var DARK_COLUMNS = [
  'received','clientTs','station','operator','mode',
  // mesh is per-screen and IN SCREEN ORDER ("155, 155, 230"), not a set of
  // counts used — that ordering is what makes the column worth having
  'order','screens','mesh','reshoots','reshootScreens','prepMin','finishMin',
  'prepRunaway','finishRunaway',
  'coated','reclaimed','stretched','cleaned','shiftMin','shiftRunaway',
  'humidity','problems','problemOther','notes'
];

/* payload holds the whole location object the pipeline predicted — job,
   rotation, confidence, warnings, source — so the tablet gets back exactly
   the shape it would have got from /port/predict and needs no second code
   path for reading it. */
var RELEASE_COLUMNS = [
  'received','key','graphic','location','approvedBy','from','note',
  'flashes','body','method','rotationText','payload'
];

// ================================================================== write
function doPost(e) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    var p = JSON.parse(e.postData.contents);
    if (SHARED_TOKEN && p.token !== SHARED_TOKEN) {
      out.setContent(JSON.stringify({ok: false, message: 'bad token'}));
      return out;
    }
    if (p.source === 'release') { out.setContent(JSON.stringify(putRelease(p))); return out; }

    var isDark = (p.source === 'darkroom');
    var cols   = isDark ? DARK_COLUMNS : PRESS_COLUMNS;
    tab(isDark ? DARK_TAB : PRESS_TAB, cols).appendRow(buildRow(cols, p));
    out.setContent(JSON.stringify({ok: true}));
  } catch (err) {
    out.setContent(JSON.stringify({ok: false, message: String(err)}));
  }
  return out;
}

/* A release REPLACES the previous one for that graphic|location. Art
   re-approving after a correction must not leave the old order sitting
   above it for an operator to scroll to. */
function putRelease(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh  = tab(RELEASE_TAB, RELEASE_COLUMNS);
    var key = String(p.graphic || '') + '|' + String(p.location || '');
    var row = buildRow(RELEASE_COLUMNS, {
      key: key,
      graphic: p.graphic || '',
      location: p.location || '',
      approvedBy: p.approved_by || 'art',
      from: p.from || 'predicted',
      note: p.note || '',
      flashes: (p.payload && p.payload.result) ? p.payload.result.flashes : '',
      body: (p.payload && p.payload.result && p.payload.result.job)
            ? p.payload.result.job.body : '',
      method: (p.payload && p.payload.result && p.payload.result.job)
            ? p.payload.result.job.method : '',
      rotationText: (p.rotation || []).map(function (s, i) {
        return (i + 1) + '. ' + (s.plate || s.ink);
      }).join('  |  '),
      payload: JSON.stringify(p.payload || {})
    });

    var keyCol = RELEASE_COLUMNS.indexOf('key') + 1;
    var last = sh.getLastRow();
    var at = 0;
    if (last > 1) {
      var keys = sh.getRange(2, keyCol, last - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (String(keys[i][0]) === key) { at = i + 2; break; }
      }
    }
    if (at) sh.getRange(at, 1, 1, RELEASE_COLUMNS.length).setValues([row]);
    else    sh.appendRow(row);

    pruneReleases(sh);
    return {ok: true, key: key, replaced: !!at};
  } finally {
    lock.releaseLock();
  }
}

/* Releases are a working set, not an archive — the JSON payloads are big and
   a job released two months ago is not going back on a press. */
function pruneReleases(sh) {
  var last = sh.getLastRow();
  if (last < 400) return;
  var cutoff = new Date(new Date().getTime() - RELEASE_KEEP_DAYS * 86400000);
  var vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] instanceof Date && vals[i][0] >= cutoff) {
      if (i > 0) sh.deleteRows(2, i);
      return;
    }
  }
}

function buildRow(cols, p) {
  var row = [];
  for (var i = 0; i < cols.length; i++) {
    var k = cols[i];
    if (k === 'received') { row.push(new Date()); continue; }
    var v = p[k];
    if (v === undefined || v === null) v = '';
    if (v === true)  v = 'YES';
    if (v === false) v = '';
    row.push(v);
  }
  return row;
}

function tab(name, cols) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(cols);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    sh.setColumnWidth(1, 168);
  }
  return sh;
}

// =================================================================== read
function doGet(e) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  var q = (e && e.parameter) || {};
  if (q.fn === 'rotation') out.setContent(JSON.stringify(getRotation(q.graphic || '')));
  else if (q.fn === 'kpi') out.setContent(JSON.stringify(getKpi(q.operator || '', q.station || '')));
  else out.setContent(JSON.stringify({ok: false, message: 'unknown fn'}));
  return out;
}

/* Returns the same shape /port/predict returns, so the tablet has one way of
   drawing a rotation regardless of where it read it from. */
function getRotation(graphic) {
  graphic = String(graphic || '').replace(/[^0-9A-Za-z_-]/g, '');
  if (!graphic) return {ok: false, message: 'no graphic'};
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RELEASE_TAB);
  if (!sh || sh.getLastRow() < 2) {
    return {ok: false, message: 'Nothing released for ' + graphic + ' yet.'};
  }
  var vals = sh.getDataRange().getValues();
  var hdr = vals[0], idx = {};
  for (var c = 0; c < hdr.length; c++) idx[hdr[c]] = c;

  var locations = [];
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][idx['graphic']]) !== graphic) continue;
    var payload = {};
    try { payload = JSON.parse(vals[r][idx['payload']] || '{}'); } catch (err) { payload = {}; }
    payload.label = payload.label || String(vals[r][idx['location']] || '');
    payload.approved = {
      approved_by: String(vals[r][idx['approvedBy']] || 'art'),
      from: String(vals[r][idx['from']] || 'predicted'),
      note: String(vals[r][idx['note']] || ''),
      ts: fmt(vals[r][idx['received']]),
      rotation: (payload.result && payload.result.rotation) || []
    };
    locations.push(payload);
  }
  if (!locations.length) {
    return {ok: false, message: 'Nothing released for ' + graphic + ' yet.'};
  }
  return {ok: true, graphic: graphic, locations: locations,
          result: locations[0].result, source: locations[0].source || ''};
}

function fmt(d) {
  return (d instanceof Date) ? Utilities.formatDate(d, TZ, 'MMM d, h:mm a') : '';
}

/**
 * EXCLUDE, DON'T CLIP. A row over the per-screen ceiling, or one the tablet
 * already flagged, is dropped from every average here and counted in
 * `excluded` so the floor can see it happened. Both rates on the screen apply
 * the same exclusion, so no two numbers on it can contradict each other.
 */
function getKpi(operator, station) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRESS_TAB);
  if (!sh || sh.getLastRow() < 2) return {ok: true, myJobs: 0, myPieces: 0, excluded: 0};

  var vals = sh.getDataRange().getValues();
  var hdr = vals[0], idx = {};
  for (var c = 0; c < hdr.length; c++) idx[hdr[c]] = c;

  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var myJobs = 0, myPieces = 0, mySetup = 0, myScreens = 0, excluded = 0;
  var pressSetup = 0, pressScreens = 0, histSetup = 0, histScreens = 0;

  for (var r = 1; r < vals.length; r++) {
    var row = vals[r];
    var when = row[idx['received']];
    var day = (when instanceof Date) ? Utilities.formatDate(when, TZ, 'yyyy-MM-dd') : '';
    var op = String(row[idx['operator']] || '');
    var st = String(row[idx['station']] || '');
    var colors = Number(row[idx['colors']] || 0);
    var setup  = Number(row[idx['setupMin']] || 0);
    var rate   = colors > 0 ? setup / colors : 0;
    var runaway = String(row[idx['setupRunaway']] || '') === 'YES'
                  || rate > SETUP_CEILING_MIN_PER_SCREEN;

    if (op === operator && !runaway && colors > 0) { histSetup += setup; histScreens += colors; }
    if (day !== today) continue;
    if (op === operator) {
      myJobs++;
      myPieces += Number(row[idx['pieces']] || 0);
      if (runaway) excluded++;
      else if (colors > 0) { mySetup += setup; myScreens += colors; }
    }
    if (st === station && !runaway && colors > 0) { pressSetup += setup; pressScreens += colors; }
  }

  var myRate    = myScreens    ? round2(mySetup / myScreens)       : null;
  var histRate  = histScreens  ? round2(histSetup / histScreens)   : null;
  var pressRate = pressScreens ? round2(pressSetup / pressScreens) : null;

  return {ok: true, myJobs: myJobs, myPieces: myPieces, mySetupMin: mySetup,
          myRate: myRate,
          myRateVsAvg: (myRate != null && histRate != null) ? round2(myRate - histRate) : null,
          pressRate: pressRate, excluded: excluded};
}

function round2(n) { return Math.round(n * 100) / 100; }

/** Run once by hand after the first rows land, to widen the readable columns. */
function formatSheets() {
  [[PRESS_TAB, PRESS_COLUMNS], [DARK_TAB, DARK_COLUMNS], [RELEASE_TAB, RELEASE_COLUMNS]]
  .forEach(function (pair) {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(pair[0]);
    if (!sh) return;
    for (var i = 0; i < pair[1].length; i++) sh.setColumnWidth(i + 1, 110);
    sh.setColumnWidth(1, 168);
    ['notes', 'rotationText', 'portPredicted', 'portRan', 'note'].forEach(function (n) {
      var k = pair[1].indexOf(n);
      if (k >= 0) sh.setColumnWidth(k + 1, 420);
    });
    var pj = pair[1].indexOf('payload');
    if (pj >= 0) sh.setColumnWidth(pj + 1, 120);
  });
}
