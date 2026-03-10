/***************
 * Job Review Confirmation.gs (FULL REPLACEMENT)
 ***************/

/** CONFIG **/
const REVIEW_SHEET = 'Review Queue';
const DATA_SHEET   = 'Data';

/** ENTRY: open the selection dialog */
function openReviewModal() {
  var html = HtmlService.createHtmlOutputFromFile('reviewModal')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Jobs to Review');
}

/** NEW: open Post-Project Notes dialog (writes to Data!AD) */
function openPostProjectNotesModal() {
  var html = HtmlService.createHtmlOutputFromFile('postProjectNotes')
    .setWidth(1100)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Post-Project Notes');
}

/** Native Sheets message */
function showUiMessage(level, message) {
  var ui = SpreadsheetApp.getUi();
  var title = level ? String(level).toUpperCase() : 'INFO';
  ui.alert(title, String(message || ''), ui.ButtonSet.OK);
}

/** Helpers **/
function htmlEscape_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function normalizeUrl_(s) {
  if (!s) return '';
  var t = String(s).trim();
  if (!t) return '';
  if (/^https?:\/\//i.test(t)) return t;
  if (/^www\./i.test(t)) return 'https://' + t;
  if (/^[A-Za-z0-9.-]+\.[A-Za-z]{2,}($|[\/?#])/i.test(t)) return 'https://' + t;
  return t;
}
function fmtDate_(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function parseMoney_(v) {
  if (v == null) return null;
  var s = String(v).replace(/[^0-9.\-]/g,'');
  if (!s) return null;
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function fmtCurrency_(n) {
  if (n == null) return '';
  try { return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(n); }
  catch (e) { return '$' + (Math.round(n * 100) / 100).toLocaleString(); }
}
function nameKey_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^a-z0-9 ]/g, '')
    .trim();
}

/** Extract embedded link from rich text or HYPERLINK() */
function getEmbeddedUrl_(range) {
  try {
    var rtv = range.getRichTextValue();
    if (rtv) {
      var whole = rtv.getLinkUrl();
      if (whole) return whole;
      var runs = rtv.getRuns();
      if (runs && runs.length) {
        for (var i = 0; i < runs.length; i++) {
          var u = runs[i].getLinkUrl();
          if (u) return u;
        }
      }
    }
  } catch (e) {}

  try {
    var f = range.getFormula();
    if (f && /^=HYPERLINK\(/i.test(f)) {
      var m = f.match(/^=HYPERLINK\(\s*"([^"]+)"/i);
      if (m && m[1]) return m[1];
    }
  } catch (e2) {}

  try {
    var dv = range.getDisplayValue();
    return normalizeUrl_(dv);
  } catch (e3) {}

  return '';
}

/** Find highest blank row within A:L, filling holes */
function findHighestBlankRow_(sh) {
  var last = Math.max(sh.getLastRow(), 2);
  var num = last - 1;
  if (num < 1) return 2;

  var rng = sh.getRange(2, 1, num, 12).getValues(); // A..L
  for (var i = 0; i < rng.length; i++) {
    var row = rng[i], allBlank = true;
    for (var c = 0; c < 12; c++) {
      if (row[c] !== '' && row[c] !== null) { allBlank = false; break; }
    }
    if (allBlank) return i + 2;
  }
  return last + 1;
}

/** Find highest blank row in AA:AC (fills holes) */
function findHighestBlankRowDraft_(sh) {
  var last = Math.max(sh.getLastRow(), 2);
  var num = last - 1;
  if (num < 1) return 2;

  var rng = sh.getRange(2, 27, num, 3).getValues(); // AA:AC
  for (var i = 0; i < rng.length; i++) {
    var row = rng[i];
    var allBlank = true;
    for (var c = 0; c < 3; c++) {
      if (row[c] !== '' && row[c] !== null) { allBlank = false; break; }
    }
    if (allBlank) return i + 2;
  }
  return last + 1;
}

/** Get draft notes by exact job name match in Data!AA */
function getDraftNotesByJobName_(jobName) {
  var name = String(jobName || '').trim();
  if (!name) return '';
  var sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET);
  if (!sh) return '';
  var last = sh.getLastRow();
  if (last < 2) return '';

  var aa = sh.getRange(2, 27, last - 1, 1).getValues(); // AA
  for (var i = 0; i < aa.length; i++) {
    var v = String(aa[i][0] || '').trim();
    if (v === name) return String(sh.getRange(i + 2, 29).getDisplayValue() || ''); // AC
  }
  return '';
}

/** Remove draft row by exact job name match in Data!AA (clears AA:AC on that row) */
function removeDraftByJobName_(jobName) {
  var name = String(jobName || '').trim();
  if (!name) return 0;
  var sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET);
  if (!sh) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;

  var aa = sh.getRange(2, 27, last - 1, 1).getValues(); // AA
  for (var i = 0; i < aa.length; i++) {
    var v = String(aa[i][0] || '').trim();
    if (v === name) {
      sh.getRange(i + 2, 27, 1, 3).clearContent(); // AA:AC
      return 1;
    }
  }
  return 0;
}

/** Upsert draft row by exact job name match in Data!AA */
function upsertDraft_(jobName, dealValue, notes) {
  var name = String(jobName || '').trim();
  if (!name) return { ok:false, reason:'missing_name' };

  var sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET);
  if (!sh) return { ok:false, reason:'no_data_sheet' };

  var last = sh.getLastRow();
  if (last < 2) last = 2;

  var aa = (last >= 2) ? sh.getRange(2, 27, Math.max(last - 1, 1), 1).getValues() : [];
  for (var i = 0; i < aa.length; i++) {
    var v = String(aa[i][0] || '').trim();
    if (v === name) {
      sh.getRange(i + 2, 27, 1, 3).setValues([[name, dealValue || '', notes || '']]);
      return { ok:true, action:'updated', row:i+2 };
    }
  }

  var r = findHighestBlankRowDraft_(sh);
  sh.getRange(r, 27, 1, 3).setValues([[name, dealValue || '', notes || '']]);
  return { ok:true, action:'inserted', row:r };
}

/** Clear ONLY matching rows in Data!M2:U where Data!N (Pipedrive ID) matches */
function clearDataMUByPipedriveIds_(pipedriveIds) {
  var ids = (pipedriveIds || [])
    .map(function(x){ return String(x || '').trim(); })
    .filter(Boolean);

  if (!ids.length) return;

  var sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET);
  if (!sh) return;

  var last = sh.getLastRow();
  if (last < 2) return;

  // M..U = 9 cols. In that block: N is the 2nd col.
  var block = sh.getRange(2, 13, last - 1, 9).getValues(); // M2:U
  var idSet = {};
  ids.forEach(function(id){ idSet[id] = true; });

  for (var i = 0; i < block.length; i++) {
    var pid = String(block[i][1] || '').trim();
    var source = String(block[i][8] || '').trim(); // U column (source marker)
    if (pid && idSet[pid] && source !== 'MANUAL_ENTRY') {
      sh.getRange(i + 2, 13, 1, 9).clearContent();
    }
  }
}

/** recipients = all emails in W2:W, Y2 only, and shop Z per salesman present (match using X->Z) */
function collectRecipients_(salesmenInJobs) {
  const sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET);
  if (!sh) return [];

  const last = sh.getLastRow();
  if (last < 2) return [];

  const COL_W = 23, COL_X = 24, COL_Y = 25, COL_Z = 26; // W X Y Z
  const set = {};

  function addEmails(s) {
    if (!s) return;
    String(s).split(/[;,]/).forEach(p => {
      const e = p.trim();
      if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) set[e.toLowerCase()] = true;
    });
  }

  // All emails in W2:W
  const wVals = sh.getRange(2, COL_W, last - 1, 1).getValues();
  for (let i = 0; i < wVals.length; i++) addEmails(wVals[i][0]);

  // Head of Sales in Y2 only
  addEmails(String(sh.getRange(2, COL_Y).getDisplayValue() || '').trim());

  // Build map X -> Z
  const xzVals = sh.getRange(2, COL_X, last - 1, 2).getValues(); // X and Z
  const map = [];
  for (let i = 0; i < xzVals.length; i++) {
    const rawName = String(xzVals[i][0] || '').trim();
    const shop    = String(xzVals[i][1] || '').trim();
    const k = nameKey_(rawName);
    if (!k || !shop) continue;
    map.push({ key: k, shop: shop });
  }

  const sales = Array.isArray(salesmenInJobs) ? salesmenInJobs : [];
  for (let s = 0; s < sales.length; s++) {
    const wantedKey = nameKey_(sales[s]);
    if (!wantedKey) continue;

    let matched = map.filter(m => m.key === wantedKey);
    if (!matched.length) matched = map.filter(m => m.key.includes(wantedKey) || wantedKey.includes(m.key));

    for (let j = 0; j < matched.length; j++) addEmails(matched[j].shop);
  }

  return Object.keys(set);
}

/** Initial payload for first dialog */
function getInitialData() {
  var ss = SpreadsheetApp.getActive();
  var shQ = ss.getSheetByName(REVIEW_SHEET);
  var shD = ss.getSheetByName(DATA_SHEET);

  var owner = '';
  var reviewers = [];
  var jobs = [];

  if (!shQ) return { owner:'', reviewers:[], jobs:[] };

  owner = String(shQ.getRange('A2').getDisplayValue() || '').trim();

  // Review Queue: B..J
  var lastRow = shQ.getLastRow();
  if (lastRow >= 2) {
    var vals = shQ.getRange(2, 2, lastRow - 1, 9).getValues(); // B..J
    for (var i = 0; i < vals.length; i++) {
      var r = vals[i];
      var projectName = String(r[0] || '').trim(); // B
      var salesman    = String(r[2] || '').trim(); // D
      var startDate   = fmtDate_(r[3]);            // E
      if (!projectName) continue;

      if (owner) {
        if (nameKey_(salesman) !== nameKey_(owner)) continue;
      }

      jobs.push({ row: i + 2, projectName: projectName, startDate: startDate });
    }
  }

  // Reviewers list from Data!V2:V
  if (shD) {
    var raw = shD.getRange('V2:V').getValues();
    var seen = {};
    for (var j = 0; j < raw.length; j++) {
      var v = String(raw[j][0] || '').trim();
      if (!v) continue;
      var k = v.toLowerCase();
      if (!seen[k]) { seen[k] = true; reviewers.push(v); }
    }
    reviewers.sort(function(a,b){ return a.localeCompare(b); });
  }

  return { owner: owner, reviewers: reviewers, jobs: jobs };
}

/** Get full details for selected rows only (Review Queue B:J) */
function getJobDetailsByRows(rows) {
  if (!Array.isArray(rows) || !rows.length) return [];
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REVIEW_SHEET);
  if (!sh) return [];

  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r < 2) continue;

    // B..J
    var v = sh.getRange(r, 2, 1, 9).getValues()[0];

    var projectName       = String(v[0] || '').trim(); // B
    var pipedriveId       = String(v[1] || '').trim(); // C
    var salesmanDisplay   = String(v[2] || '').trim(); // D
    var startDate         = fmtDate_(v[3]);            // E
    var dealValue         = v[4];                      // F
    var estMaterialCost   = v[5];                      // G
    var estMandays        = v[6];                      // H
    var materialListLink  = normalizeUrl_(getEmbeddedUrl_(sh.getRange(r, 9))); // I
    var materialListNotes = String(v[8] || '').trim(); // J

    if (!projectName) continue;

    var draftNotes = getDraftNotesByJobName_(projectName);

    out.push({
      row: r,
      projectName: projectName,
      pipedriveId: pipedriveId,
      salesmanDisplay: salesmanDisplay,
      startDate: startDate,
      dealValue: dealValue,
      estMaterialCost: estMaterialCost,
      estMandays: estMandays,
      materialListLink: materialListLink,
      materialListNotes: materialListNotes,
      draftNotes: draftNotes || ''
    });
  }
  return out;
}

/** Open the review dialog */
function openReviewSummary(selectedRows, reviewers) {
  if (!Array.isArray(selectedRows) || !selectedRows.length) { showUiMessage('warning', 'Select at least one job.'); return; }
  if (!Array.isArray(reviewers) || !reviewers.length) { showUiMessage('warning', 'Select at least one reviewer.'); return; }

  var selected = getJobDetailsByRows(selectedRows);
  if (!selected.length) { showUiMessage('warning', 'Selected rows not found.'); return; }

  var t = HtmlService.createTemplateFromFile('reviewSummary');
  t.payload = { reviewers: reviewers, jobs: selected };
  var html = t.evaluate().setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Review');
}

/** Build HTML email body */
function buildEmailBody_(records, reviewers) {
  var esc = htmlEscape_;

  var rows = records.map(function(r, idx){
    var mat = r.materialListLink ? '<a href="' + esc(r.materialListLink) + '">Open</a>' : '';
    var deal = fmtCurrency_(parseMoney_(r.dealValue));
    var estMat = fmtCurrency_(parseMoney_(r.estMaterialCost));
    var estMandays = (r.estMandays == null ? '' : esc(String(r.estMandays)));

    var matNotes = esc(r.materialListNotes || '');
    var comments = esc(r.comment || '');

    var typeText = (String(r.mode || '').toLowerCase() === 'draft') ? 'DRAFT' : 'COMPLETED';
    var typeStyle = (typeText === 'DRAFT')
      ? 'background:#fff3cd;color:#856404;border:1px solid #ffeeba;'
      : 'background:#d4edda;color:#155724;border:1px solid #c3e6cb;';

    return '' +
      '<tr style="border-top:1px solid #ddd">' +
        '<td style="padding:6px 8px">'+(idx+1)+'</td>' +
        '<td style="padding:6px 8px"><b>'+esc(r.projectName||'')+'</b><br><span style="color:#555">Start: '+esc(r.startDate||'')+'</span></td>' +
        '<td style="padding:6px 8px">'+esc(r.salesmanDisplay||'')+'</td>' +
        '<td style="padding:6px 8px">'+esc(r.pipedriveId||'')+'</td>' +
        '<td style="padding:6px 8px; white-space:nowrap">'+esc(deal)+'</td>' +
        '<td style="padding:6px 8px">'+mat+'</td>' +
        '<td style="padding:6px 8px; white-space:nowrap">'+esc(estMat)+'</td>' +
        '<td style="padding:6px 8px; white-space:nowrap">'+estMandays+'</td>' +
        '<td style="padding:6px 8px; white-space:nowrap"><span style="display:inline-block;padding:2px 8px;border-radius:999px;font-weight:700;font-size:12px;'+typeStyle+'">'+typeText+'</span></td>' +
      '</tr>' +
      '<tr>' +
        '<td></td>' +
        '<td colspan="8" style="padding:0 8px 10px 8px">' +
          '<div style="background:#fafafa;border:1px solid #eee;padding:8px;border-radius:6px">' +
            '<b>Material List Notes</b><br>' + (matNotes ? matNotes.replace(/\n/g,'<br>') : '-') + '<br><br>' +
            '<b>Project Manager Notes</b><br>' + (comments ? comments.replace(/\n/g,'<br>') : '-') +
          '</div>' +
        '</td>' +
      '</tr>';
  }).join('');

  var reviewerLine = esc((reviewers || []).join(', '));

  return '' +
    '<div style="font:14px Arial, sans-serif">' +
      '<h2 style="margin:0 0 6px 0">Install Review Summary</h2>' +
      '<div style="margin:0 0 10px 0">Reviewed by: '+ reviewerLine +'</div>' +
      '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse; width:100%">' +
        '<thead><tr style="text-align:left;background:#f3f3f3">' +
          '<th style="padding:6px 8px">#</th>' +
          '<th style="padding:6px 8px">Project</th>' +
          '<th style="padding:6px 8px">Salesman</th>' +
          '<th style="padding:6px 8px">Pipedrive ID</th>' +
          '<th style="padding:6px 8px">Deal Value</th>' +
          '<th style="padding:6px 8px">Material List</th>' +
          '<th style="padding:6px 8px">Estimated Material</th>' +
          '<th style="padding:6px 8px">Estimated Mandays</th>' +
          '<th style="padding:6px 8px">Review Type</th>' +
        '</tr></thead>' +
        '<tbody>' + rows + '</tbody>' +
      '</table>' +
    '</div>';
}

/** Process a SINGLE job (called by UI one-at-a-time) */
function processSingleJobReview(submission) {
  try {
    var reviewers = (submission && submission.reviewers) || [];
    var row = (submission && submission.row);
    var comment = String((submission && submission.comment) || '').trim();
    var overrideML = normalizeUrl_((submission && submission.materialListLink) || '');
    var mode = String((submission && submission.mode) || 'completed').toLowerCase();
    if (mode !== 'draft' && mode !== 'completed') mode = 'completed';

    if (!row || row < 2) {
      showUiMessage('warning', 'Invalid job row.');
      return { ok:false, reason:'bad_row' };
    }

    var detailsArr = getJobDetailsByRows([row]);
    if (!detailsArr.length) {
      showUiMessage('warning', 'Job not found for row ' + row);
      return { ok:false, reason:'not_found' };
    }

    var j = detailsArr[0];
    var sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET);
    if (!sh) {
      showUiMessage('error', 'Sheet "' + DATA_SHEET + '" not found.');
      return { ok:false, reason:'no_data_sheet' };
    }

    var reviewedBy = Array.isArray(reviewers) ? reviewers.join(', ') : String(reviewers || '');
    var matLink = overrideML || j.materialListLink || '';

    if (mode === 'draft') {
      var up = upsertDraft_(j.projectName, j.dealValue, comment);
      if (!up.ok) {
        showUiMessage('error', 'Draft save failed: ' + (up.reason || 'unknown'));
        return { ok:false, reason:'draft_failed' };
      }
    } else {
      var targetRow = findHighestBlankRow_(sh);

      sh.getRange(targetRow, 1, 1, 8).setValues([[
        j.projectName || '',     // A
        j.pipedriveId || '',     // B
        j.salesmanDisplay || '', // C
        reviewedBy,              // D
        '',                      // E Project Status blank
        j.startDate || '',       // F
        matLink,                 // G
        j.estMaterialCost || ''  // H
      ]]);

      sh.getRange(targetRow, 10).setValue(j.estMandays || ''); // J
      sh.getRange(targetRow, 12).setValue(comment);            // L

      removeDraftByJobName_(j.projectName);

      if (j.pipedriveId) clearDataMUByPipedriveIds_([j.pipedriveId]);
    }

    // Email recipients for this job only
    var recipients = collectRecipients_([j.salesmanDisplay]);
    if (recipients.length) {
      var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yy');
      var subject = 'Install Review Summary - ' + dateStr;
      var htmlBody = buildEmailBody_([{
        projectName: j.projectName,
        startDate: j.startDate,
        salesmanDisplay: j.salesmanDisplay,
        pipedriveId: j.pipedriveId,
        dealValue: j.dealValue,
        materialListLink: matLink,
        estMaterialCost: j.estMaterialCost,
        estMandays: j.estMandays,
        materialListNotes: j.materialListNotes,
        comment: comment,
        mode: mode
      }], reviewers);

      MailApp.sendEmail({ to: recipients.join(','), subject: subject, htmlBody: htmlBody });
    } else {
      showUiMessage('warning', 'No recipient emails found for salesman. Email not sent.');
    }

    return { ok:true, mode:mode };
  } catch (e) {
    showUiMessage('error', 'Submit failed: ' + (e && e.message ? e.message : e));
    return { ok:false, error:String(e) };
  }
}

/* =========================
 * POST-PROJECT NOTES (Data!AD)
 * - No emails
 * - Pick a completed job from Data!A
 * - Write note into Data!AD on that same row
 * ========================= */

/** List jobs from Data that have a Project Name in column A */
function getPostProjectJobs() {
  var sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];

  // Read A..AD (30 cols)
  var vals = sh.getRange(2, 1, last - 1, 30).getValues(); // A2:AD
  var out = [];

  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var projectName = String(r[0] || '').trim(); // A
    if (!projectName) continue;

    out.push({
      row: i + 2,
      projectName: projectName,
      pipedriveId: String(r[1] || '').trim(),     // B
      salesman: String(r[2] || '').trim(),        // C
      startDate: fmtDate_(r[5]),                  // F
      postNotes: String(r[29] || '').trim()       // AD
    });
  }

  out.sort(function(a,b){ return a.projectName.localeCompare(b.projectName); });
  return out;
}

/** Get details for one Data row (A..AD) */
function getPostProjectJobDetails(row) {
  var rr = Number(row);
  if (!rr || rr < 2) return null;

  var sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET);
  if (!sh) return null;

  var last = sh.getLastRow();
  if (rr > last) return null;

  var v = sh.getRange(rr, 1, 1, 30).getValues()[0]; // A..AD
  var projectName = String(v[0] || '').trim();
  if (!projectName) return null;

  return {
    row: rr,
    projectName: projectName,
    pipedriveId: String(v[1] || '').trim(),     // B
    salesman: String(v[2] || '').trim(),        // C
    reviewedBy: String(v[3] || '').trim(),      // D
    startDate: fmtDate_(v[5]),                  // F
    materialList: normalizeUrl_(getEmbeddedUrl_(sh.getRange(rr, 7))), // G
    estMaterial: v[7],                          // H
    estMandays: v[9],                           // J
    pmNotes: String(v[11] || '').trim(),        // L
    postNotes: String(v[29] || '').trim()       // AD
  };
}

/* =========================
 * CUSTOMER FOLLOW UP (Customer Follow Up tab A:F)
 * ========================= */

/** Save a follow-up entry to "Customer Follow Up" tab */
function saveFollowUp(data) {
  if (!data) return { ok: false, reason: 'No data provided.' };

  var customerName  = String(data.customerName || '').trim();
  var phone         = String(data.phone || '').trim();
  var email         = String(data.email || '').trim();
  var address       = String(data.address || '').trim();
  var pmResponsible = String(data.pmResponsible || '').trim();
  var reason        = String(data.reason || '').trim();

  if (!reason) return { ok: false, reason: 'Reason for follow up is required.' };

  var TAB_NAME = 'Customer Follow Up';
  var sh = SpreadsheetApp.getActive().getSheetByName(TAB_NAME);
  if (!sh) return { ok: false, reason: 'Tab "' + TAB_NAME + '" not found.' };

  // Find next blank row (headers in row 1)
  var last = sh.getLastRow();
  var targetRow = last + 1;
  if (last < 1) targetRow = 2; // skip header row

  sh.getRange(targetRow, 1, 1, 6).setValues([[
    customerName,   // A
    phone,          // B
    email,          // C
    address,        // D
    pmResponsible,  // E
    reason          // F
  ]]);

  return { ok: true, row: targetRow };
}

/* =========================
 * MANUAL JOB ENTRY (Data!M:U)
 * - Writes to M:T with U = "MANUAL_ENTRY" marker
 * - Prevents import from overwriting manual rows
 * ========================= */

/** Open Manual Job Entry dialog */
function openManualJobEntryModal() {
  var html = HtmlService.createHtmlOutputFromFile('manualJobEntry')
    .setWidth(640)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manual Job Entry');
}

/** Return salesman names from Data!X for the dropdown */
function getManualEntrySalesmen() {
  var sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];

  var raw = sh.getRange(2, 24, last - 1, 1).getValues(); // X2:X
  var seen = {}, out = [];
  for (var i = 0; i < raw.length; i++) {
    var v = String(raw[i][0] || '').trim();
    if (!v) continue;
    var k = v.toLowerCase();
    if (!seen[k]) { seen[k] = true; out.push(v); }
  }
  out.sort(function(a, b) { return a.localeCompare(b); });
  return out;
}

/** Find the next blank row in M:U (col 13-21), filling holes */
function findHighestBlankRowMU_(sh) {
  var last = Math.max(sh.getLastRow(), 2);
  var num = last - 1;
  if (num < 1) return 2;

  var rng = sh.getRange(2, 13, num, 9).getValues(); // M:U
  for (var i = 0; i < rng.length; i++) {
    var row = rng[i], allBlank = true;
    for (var c = 0; c < 9; c++) {
      if (row[c] !== '' && row[c] !== null) { allBlank = false; break; }
    }
    if (allBlank) return i + 2;
  }
  return last + 1;
}

/** Save a manually-entered job to Data!M:U */
function saveManualJobEntry(data) {
  if (!data) return { ok: false, reason: 'No data provided.' };

  var jobName     = String(data.jobName || '').trim();
  var pipedriveId = String(data.pipedriveId || '').trim();
  var salesman    = String(data.salesman || '').trim();
  var startDate   = data.startDate || '';
  var dealValue   = data.dealValue;
  var estMaterial = data.estMaterial;
  var estMandays  = data.estMandays;
  var materialLink = normalizeUrl_(data.materialLink || '');

  if (!jobName)     return { ok: false, reason: 'Job Name is required.' };
  if (!pipedriveId) return { ok: false, reason: 'Pipedrive ID is required.' };
  if (!salesman)    return { ok: false, reason: 'Salesman is required.' };

  var sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET);
  if (!sh) return { ok: false, reason: 'Data sheet not found.' };

  // Check for duplicate Pipedrive ID in N column
  var last = sh.getLastRow();
  if (last >= 2) {
    var nVals = sh.getRange(2, 14, last - 1, 1).getValues(); // N2:N
    for (var i = 0; i < nVals.length; i++) {
      if (String(nVals[i][0] || '').trim() === pipedriveId) {
        return { ok: false, reason: 'Pipedrive ID ' + pipedriveId + ' already exists in Data!N (row ' + (i + 2) + ').' };
      }
    }
  }

  var targetRow = findHighestBlankRowMU_(sh);

  sh.getRange(targetRow, 13, 1, 9).setValues([[
    jobName,                        // M
    pipedriveId,                    // N
    salesman,                       // O
    startDate || '',                // P
    dealValue != null ? dealValue : '',   // Q
    estMaterial != null ? estMaterial : '',// R
    estMandays != null ? estMandays : '', // S
    '',                             // T (link set below)
    'MANUAL_ENTRY'                  // U (source marker)
  ]]);

  // Formatting
  sh.getRange(targetRow, 17, 1, 1).setNumberFormat('"$"#,##0.00'); // Q
  sh.getRange(targetRow, 18, 1, 1).setNumberFormat('"$"#,##0.00'); // R
  sh.getRange(targetRow, 19, 1, 1).setNumberFormat('0.00');        // S
  if (startDate) sh.getRange(targetRow, 16, 1, 1).setNumberFormat('yyyy-mm-dd'); // P

  // Material List link as HYPERLINK
  if (materialLink) {
    sh.getRange(targetRow, 20)
      .setFormula('=HYPERLINK("' + materialLink.replace(/"/g, '""') + '","Open")');
  }

  return { ok: true, row: targetRow };
}

/** Save Data!AD for the selected job row */
function savePostProjectNote(row, noteText) {
  var rr = Number(row);
  if (!rr || rr < 2) { showUiMessage('warning', 'Invalid row.'); return { ok:false }; }

  var sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET);
  if (!sh) { showUiMessage('error', 'Data sheet not found.'); return { ok:false }; }

  var last = sh.getLastRow();
  if (rr > last) { showUiMessage('warning', 'Row out of range.'); return { ok:false }; }

  // AD = 30
  sh.getRange(rr, 30).setValue(String(noteText || '').trim());
  return { ok:true };
}