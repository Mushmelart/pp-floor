/**
 * PP FLOOR — weekly report.
 * Same spreadsheet, second script file. Builds an HTML report, converts it to
 * PDF and emails it on a schedule.
 *
 * SETUP
 *   1. Paste alongside PP_Floor_AppsScript.gs in the same Apps Script project.
 *   2. Set RECIPIENTS below.
 *   3. Run installWeeklyTrigger() once by hand (it clears its own duplicates).
 *   4. sendWeeklyReport() runs it immediately if you want to see one now.
 *
 * TWO RULES THIS REPORT KEEPS
 *   EXCLUDE, DON'T CLIP. A runaway row is dropped from every average and
 *   listed openly by order number at the bottom. Nothing is silently trimmed.
 *   ONE EXCLUSION EVERYWHERE. Every average that touches a runaway-prone
 *   value applies the same test, so no two numbers in the report contradict.
 */

// ---------------------------------------------------------------- knobs
var RECIPIENTS = 'michellemoxley@gmail.com';   // comma-separated
var REPORT_NAME = 'PrintPro — press & screen room';
var SETUP_CEILING_MIN_PER_SCREEN = 25;   // runaway: setup ÷ colors above this
var RUN_CEILING_MIN   = 600;             // runaway: a single run over 10h
var WAIT_CEILING_MIN  = 480;             // runaway: a single wait over a shift
var SETUP_TARGET_MIN_PER_SCREEN = 8;     // the shop's target, drawn as the line
var WEEK_STARTS_ON = 1;                  // 1 = Monday
var TZ2 = 'America/New_York';

var PRESS_TAB2 = 'PRESS';
var DARK_TAB2  = 'DARKROOM';

// ---------------------------------------------------------------- trigger
function installWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendWeeklyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
}

function sendWeeklyReport() {
  var now = new Date();
  var thisStart = weekStart(now);
  var lastStart = new Date(thisStart.getTime() - 7 * 86400000);

  var cur  = collect(lastStartPlus(thisStart, 0), thisStart, now);
  var prev = collect(lastStart, lastStart, thisStart);

  var html = render(cur, prev, thisStart, now);
  var blob = Utilities.newBlob(html, 'text/html', 'report.html').getAs('application/pdf')
              .setName('PP Floor ' + fmtD(thisStart) + '.pdf');
  MailApp.sendEmail({
    to: RECIPIENTS,
    subject: REPORT_NAME + ' — week of ' + fmtD(thisStart),
    htmlBody: html,
    attachments: [blob]
  });
}

function lastStartPlus(d) { return d; }

// ---------------------------------------------------------------- gather
function collect(from, windowStart, windowEnd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = {
    jobs: 0, pieces: 0, setupMin: 0, screens: 0, runMin: 0, waitMin: 0,
    excluded: [], delays: {}, problems: {}, operators: {},
    portAgree: 0, portOverride: 0, portNone: 0,
    agreeSetup: 0, agreeScreens: 0, overSetup: 0, overScreens: 0,
    dkJobs: 0, dkScreens: 0, dkReshoots: 0, dkPrep: 0, dkShiftTasks: 0,
    coveredOrders: {}, pressOrders: {}
  };

  var ps = ss.getSheetByName(PRESS_TAB2);
  if (ps && ps.getLastRow() > 1) {
    var v = ps.getDataRange().getValues(), h = hdr(v[0]);
    for (var r = 1; r < v.length; r++) {
      var row = v[r], when = row[h.received];
      if (!(when instanceof Date) || when < windowStart || when >= windowEnd) continue;

      var colors = num(row[h.colors]), setup = num(row[h.setupMin]);
      var run = num(row[h.runMin]), wait = num(row[h.waitMin]);
      var rate = colors > 0 ? setup / colors : 0;
      var bad = yes(row[h.setupRunaway]) || rate > SETUP_CEILING_MIN_PER_SCREEN
                || yes(row[h.runRunaway]) || run > RUN_CEILING_MIN
                || yes(row[h.waitRunaway]) || wait > WAIT_CEILING_MIN;

      out.jobs++;
      out.pieces += num(row[h.pieces]);
      out.pressOrders[String(row[h.order])] = true;

      var d = String(row[h.portDecision] || '');
      if (d === 'agree') out.portAgree++;
      else if (d === 'override') out.portOverride++;
      else out.portNone++;

      if (bad) {
        out.excluded.push({order: String(row[h.order]), rate: round1(rate),
                           run: round1(run), wait: round1(wait),
                           operator: String(row[h.operator] || '')});
      } else {
        if (colors > 0) { out.setupMin += setup; out.screens += colors; }
        out.runMin += run; out.waitMin += wait;
        if (colors > 0 && d === 'agree')    { out.agreeSetup += setup; out.agreeScreens += colors; }
        if (colors > 0 && d === 'override') { out.overSetup  += setup; out.overScreens  += colors; }
        var op = String(row[h.operator] || '—');
        var o = out.operators[op] || (out.operators[op] = {jobs: 0, setup: 0, screens: 0, pieces: 0});
        o.jobs++; o.setup += setup; o.screens += colors; o.pieces += num(row[h.pieces]);
      }
      tally(out.delays,   row[h.delays],   row[h.delayOther]);
      tally(out.problems, row[h.problems], row[h.problemOther]);
    }
  }

  var dk = ss.getSheetByName(DARK_TAB2);
  if (dk && dk.getLastRow() > 1) {
    var dv = dk.getDataRange().getValues(), dh = hdr(dv[0]);
    for (var q = 1; q < dv.length; q++) {
      var drow = dv[q], dwhen = drow[dh.received];
      if (!(dwhen instanceof Date) || dwhen < windowStart || dwhen >= windowEnd) continue;
      if (String(drow[dh.mode]) === 'shift') { out.dkShiftTasks++; continue; }
      out.dkJobs++;
      out.dkScreens  += num(drow[dh.screens]);
      out.dkReshoots += num(drow[dh.reshoots]);
      out.dkPrep     += num(drow[dh.prepMin]) + num(drow[dh.finishMin]);
      out.coveredOrders[String(drow[dh.order])] = true;
    }
  }
  return out;
}

function hdr(row) { var m = {}; for (var i = 0; i < row.length; i++) m[row[i]] = i; return m; }
function num(v)  { return Number(v || 0) || 0; }
function yes(v)  { return String(v || '') === 'YES'; }
function tally(bag, csv, other) {
  String(csv || '').split(';').forEach(function (s) {
    s = s.trim(); if (!s) return;
    if (s === 'Other' && other) s = 'Other: ' + other;
    bag[s] = (bag[s] || 0) + 1;
  });
}
function round1(n) { return Math.round(n * 10) / 10; }
function rate(setup, screens) { return screens ? round1(setup / screens) : null; }
function weekStart(d) {
  var x = new Date(d); x.setHours(0, 0, 0, 0);
  var diff = (x.getDay() - WEEK_STARTS_ON + 7) % 7;
  x.setDate(x.getDate() - diff); return x;
}
function fmtD(d) { return Utilities.formatDate(d, TZ2, 'MMM d, yyyy'); }

// ---------------------------------------------------------------- render
function render(c, p, start, end) {
  var s = '';
  s += '<div style="font:14px/1.5 Helvetica,Arial,sans-serif;color:#1a1a1a;max-width:760px">';
  s += '<h1 style="font-size:22px;margin:0 0 4px">' + REPORT_NAME + '</h1>';
  s += '<div style="color:#666;margin-bottom:22px">Week of ' + fmtD(start) +
       ' &middot; through ' + fmtD(end) + '</div>';

  s += row4([
    ['Jobs',   c.jobs,   p.jobs],
    ['Pieces', c.pieces, p.pieces],
    ['Min / screen', rate(c.setupMin, c.screens), rate(p.setupMin, p.screens), true],
    ['Run hours', round1(c.runMin / 60), round1(p.runMin / 60)]
  ]);

  // --- the question the whole build exists to answer ---
  var ra = rate(c.agreeSetup, c.agreeScreens), ro = rate(c.overSetup, c.overScreens);
  var answered = c.portAgree + c.portOverride;
  s += h2('The released rotation');
  s += '<table style="width:100%;border-collapse:collapse">' +
       tr(['Ran as released', c.portAgree + (answered ? ' (' + pct(c.portAgree, answered) + ')' : '')]) +
       tr(['Changed on the floor', c.portOverride + (answered ? ' (' + pct(c.portOverride, answered) + ')' : '')]) +
       tr(['No answer given', c.portNone]) +
       tr(['Setup when run as released', ra != null ? ra + ' min/screen' : '—']) +
       tr(['Setup when changed', ro != null ? ro + ' min/screen' : '—']) +
       '</table>';
  if (ra != null && ro != null) {
    var better = ra <= ro ? 'the released order' : 'the floor\'s order';
    s += '<p style="margin:8px 0 0;color:#444">Cheaper this week: <b>' + better + '</b>, by ' +
         round1(Math.abs(ra - ro)) + ' min/screen. Target is ' +
         SETUP_TARGET_MIN_PER_SCREEN + '.</p>';
  } else {
    s += '<p style="margin:8px 0 0;color:#888">Not enough answered jobs on both sides to compare yet.</p>';
  }

  s += h2('Delays');
  s += bag(c.delays);
  var loggedMin = c.setupMin + c.runMin;
  s += '<p style="color:#444;margin:6px 0 0">Waiting was ' +
       (loggedMin ? pct(c.waitMin, loggedMin + c.waitMin) : '—') +
       ' of logged floor time.</p>';

  s += h2('Problems');
  s += bag(c.problems);

  s += h2('By operator');
  s += '<table style="width:100%;border-collapse:collapse">' +
       trh(['Operator', 'Jobs', 'Pieces', 'Min/screen']);
  Object.keys(c.operators).sort().forEach(function (k) {
    var o = c.operators[k];
    s += trh([k, o.jobs, o.pieces, rate(o.setup, o.screens) || '—'], true);
  });
  s += '</table>';

  s += h2('Screen room');
  var covered = 0, total = 0;
  Object.keys(c.pressOrders).forEach(function (k) {
    total++; if (c.coveredOrders[k]) covered++;
  });
  s += '<table style="width:100%;border-collapse:collapse">' +
       tr(['Jobs prepped', c.dkJobs]) +
       tr(['Screens', c.dkScreens]) +
       tr(['Reshoots', c.dkReshoots + (c.dkScreens ? ' (' + pct(c.dkReshoots, c.dkScreens) + ' of screens)' : '')]) +
       tr(['Prep + finish hours', round1(c.dkPrep / 60)]) +
       tr(['Shift tasks logged', c.dkShiftTasks]) +
       tr(['Press jobs with a screen-room log', total ? covered + ' of ' + total : '—']) +
       '</table>';

  s += h2('Excluded from the averages');
  if (!c.excluded.length) {
    s += '<p style="color:#666">None this week.</p>';
  } else {
    s += '<p style="color:#444">These ran past a ceiling and were left out of every ' +
         'average above — listed, not trimmed.</p>';
    s += '<table style="width:100%;border-collapse:collapse">' +
         trh(['Order', 'Operator', 'Min/screen', 'Run min', 'Wait min']);
    c.excluded.forEach(function (x) {
      s += trh([x.order, x.operator, x.rate, x.run, x.wait], true);
    });
    s += '</table>';
  }

  s += '<p style="color:#999;font-size:12px;margin-top:26px">Ceilings: ' +
       SETUP_CEILING_MIN_PER_SCREEN + ' min/screen setup, ' + RUN_CEILING_MIN +
       ' min run, ' + WAIT_CEILING_MIN + ' min wait. Change them at the top of the report script.</p>';
  s += '</div>';
  return s;
}

function h2(t) {
  return '<h2 style="font-size:16px;margin:26px 0 8px;padding-bottom:5px;' +
         'border-bottom:1px solid #ddd">' + t + '</h2>';
}
function tr(cells) {
  return '<tr><td style="padding:5px 0;color:#555">' + cells[0] +
         '</td><td style="padding:5px 0;text-align:right;font-weight:600">' +
         (cells[1] === null || cells[1] === undefined || cells[1] === '' ? '—' : cells[1]) +
         '</td></tr>';
}
function trh(cells, body) {
  var st = body ? 'padding:5px 8px;border-top:1px solid #eee'
                : 'padding:5px 8px;color:#888;font-size:12px;text-transform:uppercase';
  return '<tr>' + cells.map(function (c, i) {
    return '<td style="' + st + (i ? ';text-align:right' : '') + '">' +
           (c === null || c === undefined || c === '' ? '—' : c) + '</td>';
  }).join('') + '</tr>';
}
function row4(items) {
  return '<table style="width:100%;border-collapse:separate;border-spacing:8px 0"><tr>' +
    items.map(function (it) {
      var now = it[1], was = it[2], lowerIsBetter = it[3];
      var d = '';
      if (now != null && was != null && was !== 0 && now !== was) {
        var up = now > was;
        var good = lowerIsBetter ? !up : up;
        d = '<div style="font-size:12px;color:' + (good ? '#137333' : '#b3261e') + '">' +
            (up ? '▲' : '▼') + ' ' + round1(Math.abs(now - was)) + ' vs last week</div>';
      }
      return '<td style="background:#f6f7f9;border-radius:8px;padding:12px 14px;width:25%;vertical-align:top">' +
             '<div style="font-size:26px;font-weight:700">' +
             (now === null || now === undefined ? '—' : now) + '</div>' +
             '<div style="font-size:12px;color:#666;margin-top:2px">' + it[0] + '</div>' + d + '</td>';
    }).join('') + '</tr></table>';
}
function bag(o) {
  var keys = Object.keys(o).sort(function (a, b) { return o[b] - o[a]; });
  if (!keys.length) return '<p style="color:#666">None logged.</p>';
  return '<table style="width:100%;border-collapse:collapse">' +
    keys.map(function (k) { return tr([k, o[k]]); }).join('') + '</table>';
}
function pct(n, d) { return d ? Math.round((n / d) * 100) + '%' : '—'; }
