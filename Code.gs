// ============================================================
//  Standalone Reports App — Code.gs
// ============================================================

// ── CONFIG ──────────────────────────────────────────────────
// Spreadsheet holding the "Order report" / "OrderStatusProgression
// Report" sheets. Defaulted to your existing one so this app reads
// the same data as your main tool — change it if you want this app
// fully independent with its own spreadsheet.
var DATA_SS_ID = '1Szp6QrHGTAwtjy36tpPjsi9HLfsFMXKml4gBaagSjPk';

var GEMINI_API_KEY_PROP = 'GEMINI_API_KEY';
var GEMINI_MODEL        = 'gemini-2.5-flash';

// ── Web App Entry ─────────────────────────────────────────────
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getData') {
    try {
      var sheetName = e.parameter.sheet;
      var startRow  = e.parameter.startRow ? parseInt(e.parameter.startRow) : 2;
      var batchSize = 2000; // was 500 — 4x fewer round trips per load

      var ss    = SpreadsheetApp.openById(DATA_SS_ID);
      var sheet = ss.getSheetByName(sheetName);

      if (!sheet) {
        return ContentService
          .createTextOutput(JSON.stringify({ data: [], nextRow: null, totalRows: 0 }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();

      if (lastRow < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ data: [], nextRow: null, totalRows: 0 }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var endRow = Math.min(startRow + batchSize - 1, lastRow);
      var data = sheet.getRange(startRow, 1, endRow - startRow + 1, lastCol).getDisplayValues();

      return ContentService
        .createTextOutput(JSON.stringify({
          data: data,
          nextRow: endRow < lastRow ? endRow + 1 : null,
          totalRows: lastRow - 1
        }))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: err.message, data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (e && e.parameter && e.parameter.action === 'getLastUpdated') {
    var meta = PropertiesService.getScriptProperties();
    return ContentService.createTextOutput(JSON.stringify({
      lastUpdated: meta.getProperty('rpt_lastUpdated') || null,
      uploadedBy:  meta.getProperty('rpt_uploadedBy')  || 'Unknown'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  if (e && e.parameter && e.parameter.action === 'getSharedCache') {
    return ContentService.createTextOutput(JSON.stringify(getSharedCache()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var rptShareState = _buildRptShareStateFromParams(e);
  rptShareState.isApprovedUser = true; // no login/whitelist in this standalone app

  var t = HtmlService.createTemplateFromFile('UI');
  t.rptShareState = JSON.stringify(rptShareState);
  return t.evaluate()
    .setTitle('Reports')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// Reads rtab/rfrom/rto/rch/rbr query params (from "Copy Report Link")
// so a recipient opening the link lands on the same tab/date/filters.
function _buildRptShareStateFromParams(e) {
  var p = (e && e.parameter) || {};
  var hasReport = !!(p.rtab || p.rfrom || p.rto || p.rch || p.rbr);
  return {
    hasReport: hasReport,
    tab:       p.rtab || '',
    from:      p.rfrom || '',
    to:        p.rto || '',
    channel:   p.rch || '',
    brands:    p.rbr ? p.rbr.split('|') : []
  };
}

// ── doPost — receives the ZIP-parsed data from Settings > Upload ──
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    if (payload.action === 'setLastUpdated') {
      var meta = PropertiesService.getScriptProperties();
      meta.setProperty('rpt_lastUpdated', payload.lastUpdated);
      meta.setProperty('rpt_uploadedBy',  payload.uploadedBy || 'Unknown');
      return ContentService.createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.action === 'setSharedCache') {
      return ContentService.createTextOutput(JSON.stringify(setSharedCache(payload)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.action === 'uploadData') {
      var ss = SpreadsheetApp.openById(DATA_SS_ID);

      var targetSheetName = (payload.sheet || '').toString().trim();
      if (!targetSheetName) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Missing sheet name' }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      // Only ever write to these two sheets — never create random ones.
      var ALLOWED_SHEETS = ['Order report', 'OrderStatusProgression Report'];
      var isAllowed = false;
      for (var ai = 0; ai < ALLOWED_SHEETS.length; ai++) {
        if (ALLOWED_SHEETS[ai].toLowerCase() === targetSheetName.toLowerCase()) {
          targetSheetName = ALLOWED_SHEETS[ai];
          isAllowed = true;
          break;
        }
      }
      if (!isAllowed) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Sheet not allowed: ' + targetSheetName }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var sheet = ss.getSheetByName(targetSheetName);
      if (!sheet) {
        sheet = ss.insertSheet(targetSheetName);
        SpreadsheetApp.flush();
      }

      // `clear` is now a standalone flag, not tied to "chunk 0" — a normal
      // diff-based upload (only rows that are actually new/changed) never
      // wipes rows it isn't touching.
      if (payload.clear) {
        sheet.clearContents();
        SpreadsheetApp.flush();
      }

      if (payload.csv && payload.csv.length > 0) {
        var rows = Utilities.parseCsv(payload.csv);
        if (rows.length > 0) {
          var startRow = payload.startRow || (sheet.getLastRow() + 1);
          var numCols  = rows[0].length;
          sheet.getRange(startRow, 1, rows.length, numCols).setValues(rows);
        }
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Shared cache — lets a brand-new visitor see data instantly ─
function setSharedCache(payload) {
  try {
    var cache = CacheService.getScriptCache();
    var data  = {
      orderRows:  payload.orderRows  || [],
      statusRows: payload.statusRows || [],
      savedAt:    payload.savedAt    || Date.now()
    };
    var json      = JSON.stringify(data);
    var chunkSize = 90000;
    var chunks    = [];
    for (var i = 0; i < json.length; i += chunkSize) {
      chunks.push(json.substring(i, i + chunkSize));
    }
    cache.put('shared_cache_chunks', String(chunks.length), 21600);
    for (var c = 0; c < chunks.length; c++) {
      cache.put('shared_cache_' + c, chunks[c], 21600);
    }

    var meta = PropertiesService.getScriptProperties();
    meta.setProperty('rpt_lastUpdated', new Date(data.savedAt).toISOString());

    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function getSharedCache() {
  try {
    var cache    = CacheService.getScriptCache();
    var totalStr = cache.get('shared_cache_chunks');
    if (!totalStr) return { error: 'no cache' };
    var total = parseInt(totalStr);
    var json  = '';
    for (var c = 0; c < total; c++) {
      var part = cache.get('shared_cache_' + c);
      if (!part) return { error: 'cache expired' };
      json += part;
    }
    return JSON.parse(json);
  } catch (e) {
    return { error: e.toString() };
  }
}

// ── Called via google.script.run by Reports.html ──
function getSheetDataBatch(sheetName, startRow) {
  try {
    startRow = startRow || 2;
    var batchSize = 2000; // was 500 — 4x fewer round trips per load
    var ss    = SpreadsheetApp.openById(DATA_SS_ID);
    var sheet = ss.getSheetByName(sheetName);

    // Nothing uploaded yet — return empty instead of an error, so the
    // dashboard just shows zeros until the first upload happens.
    if (!sheet) return { data: [], nextRow: null, totalRows: 0 };

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { data: [], nextRow: null, totalRows: 0 };

    var endRow = Math.min(startRow + batchSize - 1, lastRow);
    var data = sheet.getRange(startRow, 1, endRow - startRow + 1, lastCol).getDisplayValues();

    if (sheetName === 'Order report') data = _fixMultibrandChannel(data);

    return {
      data: data,
      nextRow: endRow < lastRow ? endRow + 1 : null,
      totalRows: lastRow - 1
    };
  } catch (err) {
    return { data: [], error: err.toString() };
  }
}

// Multibrand FFW/FFW2-Zalora orders don't carry a real brand name in
// Column C — the real brand lives in Column AX. Rewrite Column C so
// Reports can filter these by their real brand.
function _fixMultibrandChannel(rows) {
  var AX_INDEX = 49;
  return rows.map(function(row) {
    var ch = (row[2] || '').toString().trim();
    if (/^Multibrand\s+FFW2?\s*-\s*Zalora$/i.test(ch)) {
      var brand = (row[AX_INDEX] || '').toString().trim();
      if (brand) {
        row = row.slice();
        row[2] = brand + ' - Zalora';
      }
    }
    return row;
  });
}

// ── REPORTS AI — "Ask mAIk" ─────────────────────────────────
function askReportAI(question, contextJson, historyJson) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty(GEMINI_API_KEY_PROP);
    if (!apiKey) {
      return { success: false, error: 'No API key set. Add GEMINI_API_KEY in Project Settings → Script Properties.' };
    }

    var history = [];
    try { history = JSON.parse(historyJson || '[]'); } catch (e) {}

    var systemPrompt =
      'You are the Reports assistant embedded in a logistics dashboard. ' +
      'The JSON snapshot below contains data for all three report tabs — 3MP, Brandcom, and FFS — ' +
      'for the currently selected date range ("dateRangeLabel"), under a "tabs" object keyed by tab name. ' +
      'The "currentTabOnScreen" field tells you which tab the user is currently looking at, and "today" is today\'s date (YYYY-MM-DD). ' +
      'If the question does not name a tab, answer using currentTabOnScreen; if it names a different tab ' +
      '(e.g. "how about Brandcom" or "and FFS?"), answer using that tab\'s data instead. ' +
      '\n\n' +
      'DATE RANGE REQUESTS: you cannot change the date filter yourself, but the app can. ' +
      'If the user asks about a date range that is NOT the one covered by "dateRangeLabel" (a specific day, a different ' +
      'month, "last week", "yesterday", etc — resolve relative dates using "today"), do not answer from the current ' +
      'snapshot. Instead reply with ONLY this exact text and nothing else — no HTML, no extra words: ' +
      '[[SET_DATE:YYYY-MM-DD:YYYY-MM-DD]] using the resolved start and end dates. The app will apply that range and ' +
      'ask you again automatically, and the snapshot will then match, so answer normally at that point. Only do this ' +
      'once per question — if the snapshot still does not match after that, answer with whatever is closest and say so. ' +
      '\n\n' +
      'Answer only using numbers present in this snapshot — never invent numbers that are not in it. ' +
      'If the question needs data that is not present for a reason other than the date range, say so plainly instead of guessing. ' +
      'Respond as clean HTML fragments only (no markdown, no <html>/<body>): use <p>, <strong>, <ul>/<li>, ' +
      'and for tabular answers use <table><tr><th>...</th></tr><tr><td>...</td></tr></table>. ' +
      'Keep answers short and skimmable.\n\nREPORT SNAPSHOT:\n' + contextJson;

    var contents = history.concat([{ role: 'user', content: String(question) }]).map(function(m) {
      return { role: (m.role === 'assistant' ? 'model' : 'user'), parts: [{ text: String(m.content || '') }] };
    });

    var payload = {
      system_instruction: { parts: [{ text: systemPrompt }] },
      contents: contents,
      generationConfig: { maxOutputTokens: 1024 }
    };

    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent';

    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    var body = JSON.parse(resp.getContentText());

    if (code !== 200) {
      return { success: false, error: (body && body.error && body.error.message) || ('HTTP ' + code) };
    }

    var textOut = '';
    var cand = body.candidates && body.candidates[0];
    if (cand && cand.content && cand.content.parts) {
      cand.content.parts.forEach(function(part) {
        if (part.text) textOut += part.text;
      });
    }

    return { success: true, html: textOut, text: textOut };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}