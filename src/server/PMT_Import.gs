/** IMPORT from PMT (priority) + Job Tracking (fallback) -> Data!M:U **/
function loadJobTrackingIntoData() {
  console.log("=== loadJobTrackingIntoData START ===");

  var PMT_ID     = '1idlaLvh8uXW1_fe3PfJirXcTHXYev5YrUSisLCUDCWk';
  var PMT_TAB    = 'Project Tracker';
  var JT_ID      = '1MnkDZSNcNR4RFiPE3-QfSVtEvZ_M_qS_KdRHGiFyAvY';
  var JT_TAB     = 'Job Tracking';
  var TARGET_TAB = 'Data';

  var targetSs = SpreadsheetApp.getActive();
  var dataSheet = targetSs.getSheetByName(TARGET_TAB);
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('ERROR', 'Data sheet not found.', SpreadsheetApp.ButtonSet.OK);
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

  function isBlank_(val) {
    if (val === null || val === undefined) return true;
    if (typeof val === 'string') return val.trim() === '';
    if (typeof val === 'number') return false;
    if (val instanceof Date) return false;
    return String(val).trim() === '';
  }

  // ── Read PMT (priority source) ──
  // A=Job Name, C=Pipedrive ID, E=Salesman, G=Start Date, I=Deal Value
  // N=Estimate Material, R=Estimate Mandays, AM=Issue Notes (39), AN=Material List Link (40)
  var pmtJobs = {};  // keyed by Pipedrive ID
  var pmtLinks = {}; // link URLs keyed by Pipedrive ID
  try {
    var pmtSs = SpreadsheetApp.openById(PMT_ID);
    var pmtSh = pmtSs.getSheetByName(PMT_TAB);
    if (pmtSh) {
      var pmtLastRow = pmtSh.getLastRow();
      if (pmtLastRow >= 2) {
        var pmtRows = pmtLastRow - 1;
        var pmtAtoR  = pmtSh.getRange(2, 1, pmtRows, 18).getValues();  // A..R
        var pmtAM    = pmtSh.getRange(2, 39, pmtRows, 1).getValues();  // AM
        var pmtAN    = pmtSh.getRange(2, 40, pmtRows, 1).getValues();  // AN
        var pmtRtvAN = pmtSh.getRange(2, 40, pmtRows, 1).getRichTextValues();

        for (var r = 0; r < pmtRows; r++) {
          var pid = String(pmtAtoR[r][2] || '').trim();  // C = Pipedrive ID
          if (!pid) continue;
          pmtJobs[pid] = {
            jobName:    String(pmtAtoR[r][0] || '').trim(),   // A
            salesman:   String(pmtAtoR[r][4] || '').trim(),   // E
            startDate:  pmtAtoR[r][6],                         // G
            dealValue:  pmtAtoR[r][8],                         // I
            estMat:     pmtAtoR[r][13],                        // N
            estMandays: pmtAtoR[r][17],                        // R
            issueNotes: String(pmtAM[r][0] || '').trim()       // AM
          };
          pmtLinks[pid] = extractUrl_(pmtRtvAN[r][0], pmtAN[r][0]);
        }
      }
    } else {
      console.log('WARNING: PMT tab "Project Tracker" not found, using Job Tracking only.');
    }
  } catch (e) {
    console.log('WARNING: Could not open PMT spreadsheet: ' + e.message);
  }

  // ── Read Job Tracking (fallback source) ──
  // A=Job Name, B=Pipedrive ID, C=Start Date, E=Salesman, G=Deal Value
  // L=Estimate Material, O=Estimate Mandays, X=Issue Notes, Y=Material List Link
  var jtJobs = {};
  var jtLinks = {};
  try {
    var jtSs = SpreadsheetApp.openById(JT_ID);
    var jtSh = jtSs.getSheetByName(JT_TAB);
    if (jtSh) {
      var jtLastRow = jtSh.getLastRow();
      if (jtLastRow >= 2) {
        var jtRows = jtLastRow - 1;
        var jtAtoO = jtSh.getRange(2, 1, jtRows, 15).getValues(); // A..O
        var jtX    = jtSh.getRange(2, 24, jtRows, 1).getValues();  // X
        var jtY    = jtSh.getRange(2, 25, jtRows, 1).getValues();  // Y
        var jtRtvY = jtSh.getRange(2, 25, jtRows, 1).getRichTextValues();

        for (var r = 0; r < jtRows; r++) {
          var pid = String(jtAtoO[r][1] || '').trim();  // B = Pipedrive ID
          if (!pid) continue;
          jtJobs[pid] = {
            jobName:    String(jtAtoO[r][0] || '').trim(),   // A
            salesman:   String(jtAtoO[r][4] || '').trim(),   // E
            startDate:  jtAtoO[r][2],                         // C
            dealValue:  jtAtoO[r][6],                         // G
            estMat:     jtAtoO[r][11],                        // L
            estMandays: jtAtoO[r][14],                        // O
            issueNotes: String(jtX[r][0] || '').trim()         // X
          };
          jtLinks[pid] = extractUrl_(jtRtvY[r][0], jtY[r][0]);
        }
      }
    } else {
      console.log('WARNING: Job Tracking tab not found.');
    }
  } catch (e) {
    console.log('WARNING: Could not open Job Tracking spreadsheet: ' + e.message);
  }

  // ── Merge: PMT fields win unless blank, then fall back to Job Tracking ──
  var allIds = {};
  for (var id in pmtJobs) allIds[id] = true;
  for (var id in jtJobs)  allIds[id] = true;

  // merged[pipedriveId] = { jobName, salesman, startDate, dealValue, estMat, estMandays, issueNotes, linkUrl }
  var merged = {};
  var fields = ['jobName', 'salesman', 'startDate', 'dealValue', 'estMat', 'estMandays', 'issueNotes'];

  for (var id in allIds) {
    var pmt = pmtJobs[id] || {};
    var jt  = jtJobs[id] || {};
    var m = {};
    for (var f = 0; f < fields.length; f++) {
      var key = fields[f];
      m[key] = !isBlank_(pmt[key]) ? pmt[key] : (jt[key] || '');
    }
    // Link: PMT wins if non-empty
    var pmtLink = pmtLinks[id] || '';
    var jtLink  = jtLinks[id] || '';
    m.linkUrl = pmtLink || jtLink;
    merged[id] = m;
  }

  // ── Read existing Data!N column to find rows to update vs append ──
  var lastDataRow = dataSheet.getLastRow();
  var existingNMap = {};  // pipedriveId -> row number
  if (lastDataRow >= 2) {
    var nVals = dataSheet.getRange(2, 14, lastDataRow - 1, 1).getValues();
    for (var i = 0; i < nVals.length; i++) {
      var v = String(nVals[i][0] || '').trim();
      if (v) existingNMap[v] = i + 2; // sheet row number
    }
  }

  // Also check Data!B for IDs already reviewed
  var existingBIds = {};
  if (lastDataRow >= 2) {
    var bVals = dataSheet.getRange(2, 2, lastDataRow - 1, 1).getValues();
    for (var i = 0; i < bVals.length; i++) {
      var v = String(bVals[i][0] || '').trim();
      if (v) existingBIds[v] = true;
    }
  }

  var updatedCount = 0;
  var newRows = [];
  var newLinks = [];

  for (var id in merged) {
    // Skip entirely if this Pipedrive ID already exists in Data!B
    if (existingBIds[id]) continue;

    var m = merged[id];
    var rowData = [
      m.jobName,
      id,
      m.salesman,
      m.startDate || '',
      m.dealValue || '',
      m.estMat || '',
      m.estMandays || '',
      '',           // T filled separately
      m.issueNotes
    ];

    if (existingNMap[id]) {
      // Update existing row in place
      var row = existingNMap[id];
      dataSheet.getRange(row, 13, 1, 9).setValues([rowData]);
      applyRowFormat_(dataSheet, row);
      if (m.linkUrl) {
        dataSheet.getRange(row, 20)
          .setFormula('=HYPERLINK("' + m.linkUrl.replace(/"/g, '""') + '","Open")');
      }
      updatedCount++;
    } else if (!existingBIds[id]) {
      // New job, not yet reviewed — queue for appending
      newRows.push(rowData);
      newLinks.push(m.linkUrl);
    }
  }

  // Find first blank row in M:U (cols 13-21) to append new data
  if (newRows.length) {
    var appendRow = 2; // default if no data in M:U
    var lastSheetRow = dataSheet.getLastRow();
    if (lastSheetRow >= 2) {
      var mColVals = dataSheet.getRange(2, 13, lastSheetRow - 1, 1).getValues(); // M column
      for (var i = 0; i < mColVals.length; i++) {
        if (String(mColVals[i][0] || '').trim() !== '') {
          appendRow = i + 3; // next row after this occupied one
        }
      }
    }
    dataSheet.getRange(appendRow, 13, newRows.length, 9).setValues(newRows);

    for (var i = 0; i < newRows.length; i++) {
      applyRowFormat_(dataSheet, appendRow + i);
      if (newLinks[i]) {
        dataSheet.getRange(appendRow + i, 20)
          .setFormula('=HYPERLINK("' + newLinks[i].replace(/"/g, '""') + '","Open")');
      }
    }
  }

  console.log("Updated " + updatedCount + " existing rows, appended " + newRows.length + " new rows.");
  console.log("=== loadJobTrackingIntoData COMPLETE ===");
}

/** Apply standard number formats to a single row in M:U **/
function applyRowFormat_(sheet, row) {
  sheet.getRange(row, 16, 1, 1).setNumberFormat('yyyy-mm-dd');     // P
  sheet.getRange(row, 17, 1, 1).setNumberFormat('"$"#,##0.00');   // Q
  sheet.getRange(row, 18, 1, 1).setNumberFormat('"$"#,##0.00');   // R
  sheet.getRange(row, 19, 1, 1).setNumberFormat('0.00');           // S
}
