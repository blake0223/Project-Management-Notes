/** IMPORT from PMT -> Data!M:U + Data!I,K on reviewed rows **/
function loadJobTrackingIntoData() {
  console.log("=== loadJobTrackingIntoData START ===");

  var PMT_ID     = '1idlaLvh8uXW1_fe3PfJirXcTHXYev5YrUSisLCUDCWk';
  var PMT_TAB    = 'Project Tracker';
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

  // ── Read PMT ──
  // A=Job Name, C=Pipedrive ID, E=Salesman, G=Start Date, I=Deal Value
  // N=Estimate Material, O=Actual Material, R=Estimate Mandays, S=Actual Mandays
  // AM=Issue Notes (39), AN=Material List Link (40)
  var pmtJobs = {};  // keyed by Pipedrive ID
  var pmtLinks = {}; // link URLs keyed by Pipedrive ID
  try {
    var pmtSs = SpreadsheetApp.openById(PMT_ID);
    var pmtSh = pmtSs.getSheetByName(PMT_TAB);
    if (pmtSh) {
      var pmtLastRow = pmtSh.getLastRow();
      if (pmtLastRow >= 2) {
        var pmtRows = pmtLastRow - 1;
        var pmtAtoS  = pmtSh.getRange(2, 1, pmtRows, 19).getValues();  // A..S
        var pmtAM    = pmtSh.getRange(2, 39, pmtRows, 1).getValues();  // AM
        var pmtAN    = pmtSh.getRange(2, 40, pmtRows, 1).getValues();  // AN
        var pmtRtvAN = pmtSh.getRange(2, 40, pmtRows, 1).getRichTextValues();

        for (var r = 0; r < pmtRows; r++) {
          var pid = String(pmtAtoS[r][2] || '').trim();  // C = Pipedrive ID
          if (!pid) continue;
          pmtJobs[pid] = {
            jobName:     String(pmtAtoS[r][0] || '').trim(),   // A
            salesman:    String(pmtAtoS[r][4] || '').trim(),   // E
            startDate:   pmtAtoS[r][6],                         // G
            dealValue:   pmtAtoS[r][8],                         // I
            estMat:      pmtAtoS[r][13],                        // N
            actMat:      pmtAtoS[r][14],                        // O  → Data!I
            estMandays:  pmtAtoS[r][17],                        // R
            actMandays:  pmtAtoS[r][18],                        // S  → Data!K
            issueNotes:  String(pmtAM[r][0] || '').trim()       // AM
          };
          pmtLinks[pid] = extractUrl_(pmtRtvAN[r][0], pmtAN[r][0]);
        }
      }
    } else {
      console.log('WARNING: PMT tab "Project Tracker" not found.');
    }
  } catch (e) {
    console.log('WARNING: Could not open PMT spreadsheet: ' + e.message);
  }

  // ── Build merged map from PMT only ──
  var merged = {};
  for (var id in pmtJobs) {
    var p = pmtJobs[id];
    var m = {
      jobName:    p.jobName,
      salesman:   p.salesman,
      startDate:  p.startDate,
      dealValue:  p.dealValue,
      estMat:     p.estMat,
      actMat:     p.actMat,
      estMandays: p.estMandays,
      actMandays: p.actMandays,
      issueNotes: p.issueNotes,
      linkUrl:    pmtLinks[id] || ''
    };
    // Normalize salesman name
    if (typeof m.salesman === 'string' && m.salesman.toUpperCase() === 'POG') {
      m.salesman = "Patrick O'Gara";
    }
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

  // Find first blank row in M column, scanning from row 2 downward
  if (newRows.length) {
    var appendRow = 2; // default if no data in M:U
    var maxRow = dataSheet.getMaxRows();
    if (maxRow >= 2) {
      var mColVals = dataSheet.getRange(2, 13, maxRow - 1, 1).getValues(); // M column
      for (var i = 0; i < mColVals.length; i++) {
        if (String(mColVals[i][0] || '').trim() === '') {
          appendRow = i + 2; // this is the first blank row
          break;
        }
      }
      // If no blank found, append after last row
      if (appendRow === 2 && String(mColVals[0][0] || '').trim() !== '') {
        appendRow = mColVals.length + 2;
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

  // ── Update Data!I (col 9) and K (col 11) on reviewed rows matched by Pipedrive ID in B ──
  var actualUpdated = 0;
  if (lastDataRow >= 2) {
    var bValsRefresh = dataSheet.getRange(2, 2, lastDataRow - 1, 1).getValues();
    for (var i = 0; i < bValsRefresh.length; i++) {
      var pid = String(bValsRefresh[i][0] || '').trim();
      if (!pid || !merged[pid]) continue;
      var job = merged[pid];
      if (!isBlank_(job.actMat))     dataSheet.getRange(i + 2, 9).setValue(job.actMat);    // I
      if (!isBlank_(job.actMandays)) dataSheet.getRange(i + 2, 11).setValue(job.actMandays); // K
      actualUpdated++;
    }
  }

  console.log("Updated " + updatedCount + " existing rows, appended " + newRows.length + " new rows, refreshed actuals on " + actualUpdated + " reviewed rows.");
  console.log("=== loadJobTrackingIntoData COMPLETE ===");
}

/** Apply standard number formats to a single row in M:U **/
function applyRowFormat_(sheet, row) {
  sheet.getRange(row, 16, 1, 1).setNumberFormat('yyyy-mm-dd');     // P
  sheet.getRange(row, 17, 1, 1).setNumberFormat('"$"#,##0.00');   // Q
  sheet.getRange(row, 18, 1, 1).setNumberFormat('"$"#,##0.00');   // R
  sheet.getRange(row, 19, 1, 1).setNumberFormat('0.00');           // S
}
