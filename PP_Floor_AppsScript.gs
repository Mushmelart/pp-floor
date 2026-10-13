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
 *   GET  ?fn=dash                     → weekly rollups for the manager view
 *
 * COLUMN-MAPPING DISCIPLINE. No row is ever built by hand. COLUMNS is the
 * header and buildRow() walks it, pulling each named field off the payload.
 * Add a field by putting its name in COLUMNS where you want it — nothing can
 * silently shift a column, which is the failure that quietly corrupts every
 * report downstream.
 */

// ---------------------------------------------------------------- knobs
/* Bump this whenever you paste a new version in. Open the /exec URL with
   ?fn=version to see which build is actually deployed — saving the editor
   does not change what the URL serves; only Deploy > Manage deployments >
   Edit > New version does. This is the fastest way to tell the two apart. */
var SCRIPT_VERSION = '2026-10-13 dash+untimed';

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
  'portGraphic','portLocation','portApproved','portDecision',
  'portReason','portFlashes','portPredicted','portRan',
  // NEW COLUMNS GO AT THE END, ALWAYS. A sheet that already has rows keeps
  // its header; appending is the only change that leaves existing data
  // sitting under the right names. Inserting one mid-list shifts every value
  // after it by a column, silently, on every row already written.
  // portChangeType separates a real resequencing from a flash-only call.
  'portChangeType'
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
/* `thumb` is a base64 JPEG of the art, rendered by the pipeline at release
   time. It rides this token-gated push rather than sitting on a public URL,
   so the client's artwork is never world-readable. A sheet cell holds 50,000
   characters; the pipeline steps the render down until it fits and sends ""
   if it cannot. */
var RELEASE_COLUMNS = [
  'received','key','graphic','location','approvedBy','from','note',
  'flashes','body','method','rotationText','payload',
  'thumb'   // appended, never inserted — see the note in PRESS_COLUMNS
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
    var sh     = tab(isDark ? DARK_TAB : PRESS_TAB, cols);
    sh.appendRow(buildRow(headerOf(sh, cols), p));
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
    var cols = headerOf(sh, RELEASE_COLUMNS);
    var row = buildRow(cols, {
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
      thumb: p.thumb || '',
      payload: JSON.stringify(p.payload || {})
    });

    var keyCol = cols.indexOf('key') + 1;
    var last = sh.getLastRow();
    var at = 0;
    if (last > 1) {
      var keys = sh.getRange(2, keyCol, last - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (String(keys[i][0]) === key) { at = i + 2; break; }
      }
    }
    if (at) sh.getRange(at, 1, 1, cols.length).setValues([row]);
    else    sh.appendRow(row);

    pruneReleases(sh);
    try { CacheService.getScriptCache().remove('rot_' + String(p.graphic || '')); }
    catch (e) {}
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

/* The row is built against the SHEET'S OWN header, not against the constant
   in this file. If the two ever drift the sheet wins, so a value can never
   land in a column that is not named for it. */
function headerOf(sh, fallback) {
  var w = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, w).getValues()[0]
              .map(String).filter(function (s) { return s !== ''; });
  return hdr.length ? hdr : fallback;
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

/**
 * Returns the tab, creating it and reconciling its header.
 *
 * A tab written before a column was added keeps the old header, so every new
 * row would be one value longer than the header describes and the extra field
 * would land under a blank heading — exactly the silent corruption this file
 * is built to prevent. So: any name in COLUMNS that the sheet's header does
 * not have is appended to the header here. Combined with the append-only rule
 * for COLUMNS, old rows stay aligned and new ones are labelled.
 */
function tab(name, cols) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(cols);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    sh.setColumnWidth(1, 168);
    return sh;
  }
  var width = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, width).getValues()[0].map(String);
  var missing = cols.filter(function (c) { return hdr.indexOf(c) < 0; });
  if (missing.length) {
    sh.getRange(1, width + 1, 1, missing.length)
      .setValues([missing]).setFontWeight('bold');
  }
  return sh;
}

// =================================================================== read
function doGet(e) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  var q = (e && e.parameter) || {};
  if (q.fn === 'version') {
    out.setContent(JSON.stringify({ok: true, version: SCRIPT_VERSION,
                                   releaseColumns: RELEASE_COLUMNS.length,
                                   hasThumb: RELEASE_COLUMNS.indexOf('thumb') > -1}));
    return out;
  }
  if (q.fn === 'rotation') out.setContent(JSON.stringify(getRotation(q.graphic || '')));
  else if (q.fn === 'kpi') out.setContent(JSON.stringify(getKpi(q.operator || '', q.station || '')));
  else if (q.fn === 'dash') out.setContent(JSON.stringify(getDash(q.force === '1')));
  else out.setContent(JSON.stringify({ok: false, message: 'unknown fn'}));
  return out;
}

/* Returns the same shape /port/predict returns, so the tablet has one way of
   drawing a rotation regardless of where it read it from. */
function getRotation(graphic) {
  graphic = String(graphic || '').replace(/[^0-9A-Za-z_-]/g, '');
  if (!graphic) return {ok: false, message: 'no graphic'};

  // A repeat lookup of the same job is common at a press — answer it from
  // cache instead of touching the sheet at all.
  var cache = CacheService.getScriptCache();
  var hit = cache.get('rot_' + graphic);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RELEASE_TAB);
  if (!sh || sh.getLastRow() < 2) {
    return {ok: false, message: 'Nothing released for ' + graphic + ' yet.'};
  }

  /* NARROW READS. getDataRange().getValues() pulls every release row with its
     whole payload and its base64 thumbnail — tens of thousands of characters
     per row — just to find one job. Read the graphic column alone, find the
     rows that match, then read only those. */
  var width = sh.getLastColumn();
  var hdr = sh.getRange(1, 1, 1, width).getValues()[0];
  var idx = {};
  for (var c = 0; c < hdr.length; c++) idx[String(hdr[c])] = c;
  if (idx['graphic'] === undefined) {
    return {ok: false, message: 'Nothing released for ' + graphic + ' yet.'};
  }

  var last = sh.getLastRow();
  var gcol = sh.getRange(2, idx['graphic'] + 1, last - 1, 1).getValues();
  var rows = [];
  for (var r = 0; r < gcol.length; r++) {
    if (String(gcol[r][0]) === graphic) rows.push(r + 2);
  }
  if (!rows.length) {
    return {ok: false, message: 'Nothing released for ' + graphic + ' yet.'};
  }

  var locations = [];
  for (var k = 0; k < rows.length; k++) {
    var v = sh.getRange(rows[k], 1, 1, width).getValues()[0];
    var payload = {};
    try { payload = JSON.parse(v[idx['payload']] || '{}'); } catch (err) { payload = {}; }
    payload.label = payload.label || String(v[idx['location']] || '');
    payload.thumb = idx['thumb'] !== undefined ? String(v[idx['thumb']] || '') : '';
    payload.approved = {
      approved_by: String(v[idx['approvedBy']] || 'art'),
      from: String(v[idx['from']] || 'predicted'),
      note: String(v[idx['note']] || ''),
      ts: fmt(v[idx['received']]),
      rotation: (payload.result && payload.result.rotation) || []
    };
    locations.push(payload);
  }

  var out = {ok: true, graphic: graphic, locations: locations,
             result: locations[0].result, source: locations[0].source || ''};
  try { cache.put('rot_' + graphic, JSON.stringify(out), 300); } catch (e) {}
  return out;
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

/* ==========================================================================
   fn=dash — THE MANAGER DASHBOARD'S DATA
   --------------------------------------------------------------------------
   One GET returns every number the manager view draws, already rolled up by
   week. The tablet never sees a raw row: a season of press logs is tens of
   thousands of cells, and an 11" tablet on shop wifi should not be pulling
   that to work out an average it could have been handed.

   Weeks start Monday and are named by that Monday, so a week never changes
   its label and two charts drawn from different tabs line up on the axis.

   EXCLUDE, NEVER CLIP. A runaway row — a setup timer left running over a
   lunch, a wait that ran past a shift — is dropped from every average and
   counted in `excluded`, exactly as the weekly report and the operator's own
   NUMBERS screen do it. Clipping it to a ceiling would quietly invent a
   number nobody worked; dropping it and saying so is honest and leaves the
   row in the sheet for anyone who wants to look.
   ========================================================================== */
var SETUP_TARGET_MIN_PER_SCREEN = 8;   // the shop target — same knob as K on the tablet
var DASH_CACHE_SEC = 600;

function weekOf_(when) {
  if (!(when instanceof Date)) return '';
  var ymd = Utilities.formatDate(when, TZ, 'yyyy-MM-dd').split('-');
  var d = new Date(Number(ymd[0]), Number(ymd[1]) - 1, Number(ymd[2]));
  d.setDate(d.getDate() - ((d.getDay() + 6) % 7));   // back up to Monday
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}
/* Narrow read: pull only the named columns, not the whole sheet. */
function cols_(sh, names) {
  var out = {rows: [], idx: {}};
  if (!sh || sh.getLastRow() < 2) return out;
  var width = sh.getLastColumn();
  var hdr = sh.getRange(1, 1, 1, width).getValues()[0];
  var pos = {};
  for (var c = 0; c < hdr.length; c++) pos[String(hdr[c])] = c;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  for (var i = 0; i < names.length; i++) out.idx[names[i]] = pos[names[i]];
  out.rows = vals;
  return out;
}
function v_(row, idx, name) { var c = idx[name]; return c === undefined ? '' : row[c]; }

function getDash(force) {
  var cache = CacheService.getScriptCache();
  if (!force) {
    var hit = cache.get('dash_v1');
    if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var T = SETUP_TARGET_MIN_PER_SCREEN;

  var W = {};          // week -> bucket
  var OP = {};         // operator -> week -> bucket
  var ST = {};         // station  -> week -> bucket
  var DLY = {};        // week -> reason -> jobs that hit it
  var pressOrders = {};// week -> {order: true}
  var darkOrders  = {};// week -> {order: true}
  var excluded = 0, untimed = 0, through = '';

  function bucket(store, k) {
    if (!store[k]) store[k] = {setup: 0, screens: 0, jobs: 0, atTarget: 0,
                              pieces: 0, run: 0, wait: 0,
                              dScreens: 0, dReshoots: 0,
                              portRan: 0, portChanged: 0,
                              /* how much of the row was actually filled in —
                                 an average is only worth as much as the share
                                 of jobs that fed it */
                              timed: 0, counted: 0, answered: 0};
    return store[k];
  }

  // ---------------------------------------------------------------- press
  var P = cols_(ss.getSheetByName(PRESS_TAB),
    ['received', 'station', 'operator', 'order', 'colors', 'pieces',
     'setupMin', 'runMin', 'waitMin', 'setupRunaway', 'runRunaway', 'waitRunaway',
     'delays', 'delayOther', 'portDecision']);
  for (var r = 0; r < P.rows.length; r++) {
    var row = P.rows[r];
    var when = v_(row, P.idx, 'received');
    var wk = weekOf_(when);
    if (!wk) continue;
    var day = Utilities.formatDate(when, TZ, 'yyyy-MM-dd');
    if (day > through) through = day;

    var colors = Number(v_(row, P.idx, 'colors') || 0);
    var setup  = Number(v_(row, P.idx, 'setupMin') || 0);
    var run    = Number(v_(row, P.idx, 'runMin')  || 0);
    var wait   = Number(v_(row, P.idx, 'waitMin') || 0);
    var rate   = colors > 0 ? setup / colors : 0;
    var bad    = String(v_(row, P.idx, 'setupRunaway') || '') === 'YES'
                 || rate > SETUP_CEILING_MIN_PER_SCREEN;
    var badRun = String(v_(row, P.idx, 'runRunaway')  || '') === 'YES';
    var badWt  = String(v_(row, P.idx, 'waitRunaway') || '') === 'YES';

    var b = bucket(W, wk);
    b.jobs++;
    b.pieces += Number(v_(row, P.idx, 'pieces') || 0);

    var dec = String(v_(row, P.idx, 'portDecision') || '');
    if (dec === 'ran') b.portRan++;
    else if (dec) b.portChanged++;

    var ord = String(v_(row, P.idx, 'order') || '').trim();
    if (ord) { if (!pressOrders[wk]) pressOrders[wk] = {}; pressOrders[wk][ord] = true; }

    // delay reasons: one job hitting a reason counts once, however long it ran
    var reasons = String(v_(row, P.idx, 'delays') || '').split(';');
    var other = String(v_(row, P.idx, 'delayOther') || '').trim();
    if (!DLY[wk]) DLY[wk] = {};
    for (var q = 0; q < reasons.length; q++) {
      var nm = reasons[q].trim();
      if (!nm) continue;
      if (nm === 'Other' && other) nm = other;        // the write-in is the finding
      DLY[wk][nm] = (DLY[wk][nm] || 0) + 1;
    }

    if (!badRun) b.run  += run;
    if (!badWt)  b.wait += wait;

    /* WHAT WAS ACTUALLY LOGGED. Counted before anything is excluded, because
       the point of these three is to say how much of the week the averages
       above are standing on. */
    if (setup > 0) b.timed++;
    if (Number(v_(row, P.idx, 'pieces') || 0) > 0) b.counted++;
    if (dec) b.answered++;

    /* A job submitted without the setup clock ever started comes through as
       setup = 0, which divides to 0 min/screen and would be counted as a job
       that HIT the target. That is the one way this dashboard could lie in
       the shop's favour, so an untimed job is excluded from the rate and from
       the at-target share, and reported on its own. It still counts as a job
       and its pieces still count. */
    if (bad || colors <= 0) { excluded++; continue; }
    if (setup <= 0) { untimed++; continue; }

    b.setup += setup; b.screens += colors;
    if (rate <= T) b.atTarget++;

    var op = String(v_(row, P.idx, 'operator') || '').trim();
    if (op) { if (!OP[op]) OP[op] = {}; var ob = bucket(OP[op], wk);
              ob.setup += setup; ob.screens += colors; ob.jobs++;
              if (rate <= T) ob.atTarget++; }
    var st = String(v_(row, P.idx, 'station') || '').trim();
    if (st) { if (!ST[st]) ST[st] = {}; var sb = bucket(ST[st], wk);
              sb.setup += setup; sb.screens += colors; sb.jobs++;
              if (rate <= T) sb.atTarget++; }
  }

  // ----------------------------------------------------------- screen room
  var D = cols_(ss.getSheetByName(DARK_TAB),
    ['received', 'mode', 'order', 'screens', 'reshoots']);
  for (var r2 = 0; r2 < D.rows.length; r2++) {
    var dr = D.rows[r2];
    var wk2 = weekOf_(v_(dr, D.idx, 'received'));
    if (!wk2) continue;
    if (String(v_(dr, D.idx, 'mode') || '') !== 'job') continue;
    var db = bucket(W, wk2);
    db.dScreens  += Number(v_(dr, D.idx, 'screens')  || 0);
    db.dReshoots += Number(v_(dr, D.idx, 'reshoots') || 0);
    var dord = String(v_(dr, D.idx, 'order') || '').trim();
    if (dord) { if (!darkOrders[wk2]) darkOrders[wk2] = {}; darkOrders[wk2][dord] = true; }
  }

  // --------------------------------------------------------------- shape it
  function rate_(b) { return b.screens ? round2(b.setup / b.screens) : null; }
  var keys = Object.keys(W).sort();
  var weeks = keys.map(function (k) {
    var b = W[k];
    var logged = b.setup + b.run + b.wait;
    var po = Object.keys(pressOrders[k] || {});
    var hit = 0;
    for (var i = 0; i < po.length; i++) if ((darkOrders[k] || {})[po[i]]) hit++;
    return {
      w: k,
      setupRate: rate_(b),
      atTarget:  b.jobs ? round2(100 * b.atTarget / b.jobs) : null,
      delayPct:  logged ? round2(100 * b.wait / logged)     : null,
      pieces:    b.pieces,
      jobs:      b.jobs,
      screens:   b.dScreens,
      reshoots:  b.dReshoots,
      coverage:  po.length ? round2(100 * hit / po.length)  : null,
      portRan:   b.portRan,
      portChanged: b.portChanged,
      timed:     b.timed,
      counted:   b.counted,
      answered:  b.answered
    };
  });
  function series(store) {
    var out = {};
    Object.keys(store).forEach(function (name) {
      out[name] = keys.map(function (k) {
        var b = store[name][k];
        return b ? {w: k, rate: rate_(b), jobs: b.jobs,
                    atTarget: b.jobs ? round2(100 * b.atTarget / b.jobs) : null}
                 : {w: k, rate: null, jobs: 0, atTarget: null};
      });
    });
    return out;
  }
  var delays = {};
  keys.forEach(function (k) {
    delays[k] = Object.keys(DLY[k] || {})
      .map(function (n) { return {name: n, jobs: DLY[k][n]}; })
      .sort(function (a, b) { return b.jobs - a.jobs; })
      .slice(0, 12);
  });

  var res = {ok: true, target: T, ceiling: SETUP_CEILING_MIN_PER_SCREEN,
             weeks: weeks, byOperator: series(OP), byStation: series(ST),
             delays: delays, excluded: excluded, untimed: untimed,
             through: through,
             updated: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'),
             version: SCRIPT_VERSION};
  try { cache.put('dash_v1', JSON.stringify(res), DASH_CACHE_SEC); } catch (e) {}
  return res;
}

/** Cosmetic only — widen the readable columns. Run by hand whenever.
 *  Clamped to the columns the sheet actually has: the COLUMNS list can be
 *  longer than a tab that predates a new field, and setColumnWidth past the
 *  end throws "Those columns are out of bounds." */
function formatSheets() {
  [[PRESS_TAB, PRESS_COLUMNS], [DARK_TAB, DARK_COLUMNS], [RELEASE_TAB, RELEASE_COLUMNS]]
  .forEach(function (pair) {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(pair[0]);
    if (!sh) return;
    var maxc = sh.getMaxColumns();
    var wide = function (col, px) { if (col >= 1 && col <= maxc) sh.setColumnWidth(col, px); };
    for (var i = 0; i < pair[1].length; i++) wide(i + 1, 110);
    wide(1, 168);
    ['notes', 'rotationText', 'portPredicted', 'portRan', 'note'].forEach(function (n) {
      wide(pair[1].indexOf(n) + 1, 420);
    });
    wide(pair[1].indexOf('payload') + 1, 120);
    wide(pair[1].indexOf('thumb') + 1, 90);
  });
}
