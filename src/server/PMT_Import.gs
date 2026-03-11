/** IMPORT Job Tracking -> Data!M:U (append-only, no clearing) **/
function loadJobTrackingIntoData() {
  console.log("=== loadJobTrackingIntoData START ===");

  var SOURCE_ID  = '1MnkDZSNcNR4RFiPE3-QfSVtEvZ_M_qS_KdRHGiFyAvY';
  var SOURCE_TAB = 'Job Tracking';
  var TARGET_TAB = 'Data';

  var targetSs = SpreadsheetApp.getActive();
  var dataSheet = targetSs.getSheetByName(TARGET_TAB);
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('ERROR', 'Data sheet not found.', SpreadsheetApp.ButtonSet.OK);
    return;
  }

  var srcSs = SpreadsheetApp.openById(SOURCE_ID);
  var srcSh = srcSs.getSheetByName(SOURCE_TAB);
  if (!srcSh) {
    SpreadsheetApp.getUi().alert('ERROR', 'Source tab "Job Tracking" not found.', SpreadsheetApp.ButtonSet.OK);
    return;
  }

  /** Headers **/
  var headers = [
    'Job Name',           // M
    'Pipedrive ID',       // N
    'Salesman',           // O
    'Start Date',         // P
    'Deal Value',         // Q
    'Estimate Material',  // R
    'Estimate Mandays',   // S
    'Material List Link', // T
    'Issue Notes'         // U
  ];
  dataSheet.getRange(1, 13, 1, 9).setValues([headers]);

  var lastSrcRow = srcSh.getLastRow();
  if (lastSrcRow < 2) return;

  var numRows = lastSrcRow - 1;

  // Source columns
  // A Job Name, B Pipedrive ID, C Start Date, E Salesman
  // G Deal Value, L Estimate Material, O Estimate Mandays
  // X Issue Notes, Y Material List Link
  var valsAtoO = srcSh.getRange(2, 1, numRows, 15).getValues(); // A..O
  var valsX = srcSh.getRange(2, 24, numRows, 1).getValues();   // X
  var valsY = srcSh.getRange(2, 25, numRows, 1).getValues();   // Y
  var rtvY  = srcSh.getRange(2, 25, numRows, 1).getRichTextValues();

  // Build set of Pipedrive IDs already in Data!N (col 14) to avoid duplicates
  var existingNIds = {};
  var lastDataRow = dataSheet.getLastRow();
  if (lastDataRow >= 2) {
    var nVals = dataSheet.getRange(2, 14, lastDataRow - 1, 1).getValues(); // N column
    for (var i = 0; i < nVals.length; i++) {
      var v = String(nVals[i][0] || '').trim();
      if (v) existingNIds[v] = true;
    }
  }

  // Also check Data!B for IDs already reviewed/in the system
  var existingBIds = {};
  if (lastDataRow >= 2) {
    var bVals = dataSheet.getRange(2, 2, lastDataRow - 1, 1).getValues();
    for (var i = 0; i < bVals.length; i++) {
      var v = String(bVals[i][0] || '').trim();
      if (v) existingBIds[v] = true;
    }
  }

  function extractUrl_(rtv, fallback) {
    try {
      if (rtv) {
        var u = rtv.getLinkUrl();
        if (u) return u;
        var runs = rtv.getRuns();
        if (runs) {
          for (var i = 0; i < runs.length; i++) {
            if (runs[i].getLinkUrl()) return runs[i].getLinkUrl();
          }
        }
      }
    } catch (e) {}
    return (/^https?:\/\//i.test(fallback || '')) ? fallback : '';
  }

  var out = [];
  var linkUrls = [];

  for (var r = 0; r < numRows; r++) {
    var row = valsAtoO[r];

    var jobName    = String(row[0] || '').trim();  // A
    var pipedrive  = String(row[1] || '').trim();  // B
    var startDate  = row[2];                        // C
    var salesman   = String(row[4] || '').trim();  // E
    var dealValue  = row[6];                        // G
    var estMat     = row[11];                       // L
    var estMandays = row[14];                       // O
    var issueNotes = String(valsX[r][0] || '').trim(); // X

    if (!pipedrive) continue;
    // Skip if already in M:U (col N) or already reviewed (col B)
    if (existingNIds[pipedrive]) continue;
    if (existingBIds[pipedrive]) continue;

    var linkUrl = extractUrl_(rtvY[r][0], valsY[r][0]);

    out.push([
      jobName,
      pipedrive,
      salesman,
      startDate || '',
      dealValue || '',
      estMat || '',
      estMandays || '',
      '',           // T filled after
      issueNotes
    ]);
    linkUrls.push(linkUrl);
  }

  if (!out.length) {
    console.log("No new jobs to import.");
    return;
  }

  // Append new rows after existing data (no clearing)
  var appendRow = Math.max(lastDataRow + 1, 2);
  dataSheet.getRange(appendRow, 13, out.length, 9).setValues(out);

  // Formatting
  dataSheet.getRange(appendRow, 17, out.length, 1).setNumberFormat('"$"#,##0.00'); // Q
  dataSheet.getRange(appendRow, 18, out.length, 1).setNumberFormat('"$"#,##0.00'); // R
  dataSheet.getRange(appendRow, 19, out.length, 1).setNumberFormat('0.00');        // S
  dataSheet.getRange(appendRow, 16, out.length, 1).setNumberFormat('yyyy-mm-dd');  // P

  // Apply Material List links
  for (var i = 0; i < linkUrls.length; i++) {
    if (linkUrls[i]) {
      dataSheet
        .getRange(appendRow + i, 20)
        .setFormula('=HYPERLINK("' + linkUrls[i].replace(/"/g, '""') + '","Open")');
    }
  }

  console.log("Appended " + out.length + " new jobs.");
  console.log("=== loadJobTrackingIntoData COMPLETE ===");
}
