/** IMPORT Job Tracking -> Data!M:U (with headers) **/
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

  // Preserve manually-entered rows (U = "MANUAL_ENTRY") before clearing
  var manualRows = [];
  var manualLinks = [];
  var lastDataRow = dataSheet.getLastRow();
  if (lastDataRow >= 2) {
    var existingMU = dataSheet.getRange(2, 13, lastDataRow - 1, 9).getValues(); // M:U
    for (var mr = 0; mr < existingMU.length; mr++) {
      if (String(existingMU[mr][8] || '').trim() === 'MANUAL_ENTRY') {
        manualRows.push(existingMU[mr]);
        // Capture HYPERLINK formula from T column if present
        var tCell = dataSheet.getRange(mr + 2, 20);
        var formula = tCell.getFormula();
        manualLinks.push(formula || '');
      }
    }
    dataSheet.getRange(2, 13, lastDataRow - 1, 9).clearContent();
  }

  if (lastSrcRow < 2) return;

  var numRows = lastSrcRow - 1;

  // Source columns
  // A Job Name
  // B Pipedrive ID
  // C Start Date
  // E Salesman
  // G Deal Value
  // L Estimate Material
  // O Estimate Mandays
  // X Issue Notes
  // Y Material List Link
  var valsAtoO = srcSh.getRange(2, 1, numRows, 15).getValues(); // A..O
  var valsX = srcSh.getRange(2, 24, numRows, 1).getValues();   // X
  var valsY = srcSh.getRange(2, 25, numRows, 1).getValues();   // Y
  var rtvY  = srcSh.getRange(2, 25, numRows, 1).getRichTextValues();

  // Existing IDs from Data!B
  var existingIds = {};
  var lastMainRow = dataSheet.getLastRow();
  if (lastMainRow >= 2) {
    var ids = dataSheet.getRange(2, 2, lastMainRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      var v = (ids[i][0] || '').toString().trim();
      if (v) existingIds[v] = true;
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
    if (existingIds[pipedrive]) continue;

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

  var startRow = 2;

  if (out.length) {
    dataSheet.getRange(startRow, 13, out.length, 9).setValues(out);

    // Formatting
    dataSheet.getRange(startRow, 17, out.length, 1).setNumberFormat('"$"#,##0.00'); // Q
    dataSheet.getRange(startRow, 18, out.length, 1).setNumberFormat('"$"#,##0.00'); // R
    dataSheet.getRange(startRow, 19, out.length, 1).setNumberFormat('0.00');        // S
    dataSheet.getRange(startRow, 16, out.length, 1).setNumberFormat('yyyy-mm-dd');  // P

    // Apply Material List links
    for (var i = 0; i < linkUrls.length; i++) {
      if (linkUrls[i]) {
        dataSheet
          .getRange(startRow + i, 20)
          .setFormula('=HYPERLINK("' + linkUrls[i].replace(/"/g, '""') + '","Open")');
      }
    }
  }

  // Restore manually-entered rows after imported data
  if (manualRows.length) {
    var manualStart = startRow + out.length;
    dataSheet.getRange(manualStart, 13, manualRows.length, 9).setValues(manualRows);
    // Restore HYPERLINK formulas for T column
    for (var m = 0; m < manualLinks.length; m++) {
      if (manualLinks[m]) {
        dataSheet.getRange(manualStart + m, 20).setFormula(manualLinks[m]);
      }
    }
    // Re-apply formatting for manual rows
    dataSheet.getRange(manualStart, 17, manualRows.length, 1).setNumberFormat('"$"#,##0.00'); // Q
    dataSheet.getRange(manualStart, 18, manualRows.length, 1).setNumberFormat('"$"#,##0.00'); // R
    dataSheet.getRange(manualStart, 19, manualRows.length, 1).setNumberFormat('0.00');        // S
    dataSheet.getRange(manualStart, 16, manualRows.length, 1).setNumberFormat('yyyy-mm-dd');  // P
  }

  console.log("=== loadJobTrackingIntoData COMPLETE ===");
}
