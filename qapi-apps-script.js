/**
 * COFFMAN QAPI — Google Apps Script
 *
 * SETUP INSTRUCTIONS:
 * 1. Upload Coffman_QAPI_Monthly_Input_Template.xlsx to Google Drive
 * 2. Open it in Google Sheets
 * 3. Go to Extensions > Apps Script
 * 4. Delete any existing code and paste this entire file
 * 5. Click Deploy > New deployment
 * 6. Type: Web app
 * 7. Execute as: Me
 * 8. Who has access: Anyone
 * 9. Click Deploy and copy the Web App URL
 * 10. Paste that URL into the QAPI web form setup screen
 *
 * The web form will use this script to:
 * - READ KPI definitions from the Settings tab
 * - READ existing data from department tabs
 * - WRITE form submissions back to department tabs
 */

// ============================================================
// GET handler — returns KPI config + existing data
// ============================================================
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = e.parameter.action || 'config';

    if (action === 'config') {
      return sendJson(getConfig(ss));
    }

    if (action === 'data') {
      var dept = e.parameter.dept;
      if (!dept) return sendJson({ error: 'Missing dept parameter' });
      return sendJson(getDeptData(ss, dept));
    }

    if (action === 'all') {
      return sendJson(getAllData(ss));
    }

    return sendJson({ error: 'Unknown action: ' + action });
  } catch (err) {
    return sendJson({ error: err.toString() });
  }
}

// ============================================================
// POST handler — writes form data to department tab
// ============================================================
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var payload = JSON.parse(e.postData.contents);
    var dept = payload.dept;
    var data = payload.data;

    if (!dept || !data) {
      return sendJson({ error: 'Missing dept or data' });
    }

    var result = writeDeptData(ss, dept, data);
    return sendJson(result);
  } catch (err) {
    return sendJson({ error: err.toString() });
  }
}

// ============================================================
// Get KPI config from Settings tab
// ============================================================
function getConfig(ss) {
  var ws = ss.getSheetByName('Settings');
  if (!ws) return { error: 'Settings tab not found' };

  var data = ws.getDataRange().getValues();
  var departments = {};

  // Skip header rows (rows 1-4, index 0-3)
  for (var i = 4; i < data.length; i++) {
    var row = data[i];
    var deptName = String(row[0]).trim();
    var metric = String(row[1]).trim();
    var threshold = String(row[2]).trim();
    var source = String(row[3]).trim();
    var active = String(row[4]).trim().toLowerCase();
    var sortOrder = row[5] || 999;

    if (!deptName || !metric) continue;
    if (active === 'no') continue;

    if (!departments[deptName]) {
      departments[deptName] = {
        name: deptName,
        kpis: [],
        dataSources: source || ''
      };
    }

    departments[deptName].kpis.push({
      metric: metric,
      threshold: threshold,
      sortOrder: sortOrder
    });

    // Update data sources if this row has one
    if (source && !departments[deptName].dataSources) {
      departments[deptName].dataSources = source;
    }
  }

  // Sort KPIs by sort order
  for (var key in departments) {
    departments[key].kpis.sort(function(a, b) { return a.sortOrder - b.sortOrder; });
  }

  return {
    success: true,
    facility: ss.getName(),
    departments: departments,
    sheetId: ss.getId()
  };
}

// ============================================================
// Get existing data from a department tab
// ============================================================
function getDeptData(ss, deptName) {
  var ws = ss.getSheetByName(deptName);
  if (!ws) return { error: 'Tab not found: ' + deptName };

  var data = ws.getDataRange().getValues();
  var result = {
    dept: deptName,
    reportingMonth: '',
    kpis: [],
    opportunities: [],
    plans: []
  };

  // Row 2 (index 1): reporting month in merged D2:E2
  if (data.length > 1) {
    result.reportingMonth = String(data[1][3] || '').trim();
  }

  // Find KPI data (rows 5+, index 4+)
  // KPIs run from row 5 until we hit the Opportunities header
  var kpiStart = 4; // 0-indexed row 5
  for (var i = kpiStart; i < data.length; i++) {
    var col1 = String(data[i][0] || '').trim();

    // Stop at opportunities section
    if (col1.toUpperCase().indexOf('OPPORTUNITIES') >= 0) {
      // Read opportunity rows after this
      for (var j = i + 1; j < Math.min(i + 4, data.length); j++) {
        var oppText = String(data[j][1] || '').trim();
        if (oppText || String(data[j][0] || '').match(/^\d\./)) {
          result.opportunities.push(oppText);
        }
      }

      // Find Plan section
      for (var j = i + 4; j < data.length; j++) {
        var planCol = String(data[j][0] || '').trim();
        if (planCol.toUpperCase().indexOf('PLAN FOR') >= 0) {
          // Skip the header row after "PLAN FOR IMPROVEMENT"
          for (var k = j + 2; k < Math.min(j + 6, data.length); k++) {
            var numCol = data[k][0];
            if (typeof numCol === 'number' || String(numCol).match(/^\d+$/)) {
              result.plans.push({
                action: String(data[k][1] || '').trim(),
                responsible: String(data[k][2] || '').trim(),
                dueDate: data[k][3] ? formatDate(data[k][3]) : '',
                status: String(data[k][4] || '').trim()
              });
            }
          }
          break;
        }
      }
      break;
    }

    // This is a KPI row
    if (col1 && col1 !== 'Performance Metric / KPI') {
      result.kpis.push({
        metric: col1,
        prior: String(data[i][1] || '').trim(),
        current: String(data[i][2] || '').trim(),
        threshold: String(data[i][3] || '').trim(),
        notes: String(data[i][4] || '').trim()
      });
    }
  }

  return { success: true, data: result };
}

// ============================================================
// Get summary data from ALL department tabs
// ============================================================
function getAllData(ss) {
  var config = getConfig(ss);
  if (config.error) return config;

  var allDepts = {};
  for (var key in config.departments) {
    var deptData = getDeptData(ss, key);
    if (deptData.success) {
      allDepts[key] = deptData.data;
    }
  }

  return {
    success: true,
    facility: config.facility,
    departments: config.departments,
    data: allDepts
  };
}

// ============================================================
// Write form data to a department tab
// ============================================================
function writeDeptData(ss, deptName, formData) {
  var ws = ss.getSheetByName(deptName);
  if (!ws) return { error: 'Tab not found: ' + deptName };

  var data = ws.getDataRange().getValues();

  // Write reporting month to D2
  if (formData.reportingMonth) {
    ws.getRange('D2').setValue(formData.reportingMonth);
  }

  // Write KPI data
  var kpis = formData.kpis || [];
  var kpiRowStart = 5; // 1-indexed
  for (var i = 0; i < kpis.length; i++) {
    var r = kpiRowStart + i;
    var kpi = kpis[i];
    if (kpi.prior !== undefined && kpi.prior !== '') ws.getRange(r, 2).setValue(kpi.prior);
    if (kpi.current !== undefined && kpi.current !== '') ws.getRange(r, 3).setValue(kpi.current);
    if (kpi.notes !== undefined) ws.getRange(r, 5).setValue(kpi.notes);
  }

  // Find opportunities section and write
  for (var i = 4; i < data.length; i++) {
    var col1 = String(data[i][0] || '').trim();
    if (col1.toUpperCase().indexOf('OPPORTUNITIES') >= 0) {
      var opps = formData.opportunities || [];
      for (var j = 0; j < Math.min(opps.length, 3); j++) {
        ws.getRange(i + 2 + j, 2).setValue(opps[j] || '');
      }

      // Find and write plan data
      for (var j = i + 4; j < data.length; j++) {
        var planCol = String(data[j][0] || '').trim();
        if (planCol.toUpperCase().indexOf('PLAN FOR') >= 0) {
          var plans = formData.plans || [];
          for (var k = 0; k < Math.min(plans.length, 4); k++) {
            var planRow = j + 3 + k; // 1-indexed, skip header + column headers
            var plan = plans[k];
            if (plan.action) ws.getRange(planRow, 2).setValue(plan.action);
            if (plan.responsible) ws.getRange(planRow, 3).setValue(plan.responsible);
            if (plan.dueDate) ws.getRange(planRow, 4).setValue(plan.dueDate);
            if (plan.status) ws.getRange(planRow, 5).setValue(plan.status);
          }
          break;
        }
      }
      break;
    }
  }

  return { success: true, message: deptName + ' data saved successfully' };
}

// ============================================================
// Helpers
// ============================================================
function sendJson(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDate(val) {
  if (val instanceof Date) {
    var m = val.getMonth() + 1;
    var d = val.getDate();
    var y = val.getFullYear();
    return (m < 10 ? '0' : '') + m + '/' + (d < 10 ? '0' : '') + d + '/' + y;
  }
  return String(val);
}

// ============================================================
// Menu for manual Brief to DON
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('QAPI Tools')
    .addItem('Generate Brief to DON', 'generateBriefToDON')
    .addItem('Duplicate for New Month', 'duplicateForNewMonth')
    .addToUi();
}

function generateBriefToDON() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getConfig(ss);
  var briefSheet = ss.getSheetByName('Brief to DON');
  if (!briefSheet) {
    SpreadsheetApp.getUi().alert('Brief to DON tab not found.');
    return;
  }

  var row = 5;
  for (var key in config.departments) {
    var deptData = getDeptData(ss, key);
    if (!deptData.success) {
      briefSheet.getRange(row, 2).setValue('No data');
      briefSheet.getRange(row, 3).setValue('-');
      briefSheet.getRange(row, 4).setValue('');
      row++;
      continue;
    }

    var d = deptData.data;
    var filled = d.kpis.filter(function(k) { return k.current; }).length;
    var notMet = [];

    d.kpis.forEach(function(k) {
      if (k.current && k.threshold) {
        var met = checkThreshold(k.current, k.threshold);
        if (met === false) {
          notMet.push(k.metric + ': ' + k.current + ' (target: ' + k.threshold + ')');
        }
      }
    });

    briefSheet.getRange(row, 2).setValue(filled + ' / ' + d.kpis.length);
    briefSheet.getRange(row, 3).setValue(notMet.length > 0 ? notMet.length : 'All met');
    briefSheet.getRange(row, 4).setValue(notMet.join('\n'));

    // Color code
    if (notMet.length > 0) {
      briefSheet.getRange(row, 3).setBackground('#FADBD8');
    } else if (filled > 0) {
      briefSheet.getRange(row, 3).setBackground('#D5F5E3');
    }

    row++;
  }

  SpreadsheetApp.getUi().alert('Brief to DON updated!');
}

function checkThreshold(value, threshold) {
  var v = parseFloat(value);
  if (isNaN(v)) {
    var vl = value.toLowerCase().trim();
    var tl = threshold.toLowerCase().trim();
    if (tl === 'yes') return vl === 'yes';
    if (tl === 'none') return vl === 'none' || vl === '0' || vl === '';
    return null; // Can't determine
  }

  var t = threshold.replace(/[^0-9.<>%\-+]/g, '');
  if (t.indexOf('<') === 0) return v < parseFloat(t.substring(1));
  if (t.indexOf('>') === 0) return v > parseFloat(t.substring(1));
  var tv = parseFloat(t);
  if (!isNaN(tv)) {
    if (threshold.indexOf('+') >= 0) return v >= tv;
    return v <= tv;
  }
  return null;
}

function duplicateForNewMonth() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'New Month',
    'Enter the reporting month (e.g., "April 2026"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var monthName = response.getResponseText().trim();
  if (!monthName) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var newSS = ss.copy('QAPI Monthly Input - ' + monthName);

  // Clear data from each department tab in the new copy
  var config = getConfig(ss);
  for (var key in config.departments) {
    var ws = newSS.getSheetByName(key);
    if (!ws) continue;

    // Set reporting month
    ws.getRange('D2').setValue(monthName);

    // Clear KPI data (columns B, C, E from row 5 down)
    var data = ws.getDataRange().getValues();
    for (var i = 4; i < data.length; i++) {
      var col1 = String(data[i][0] || '').trim();
      if (col1.toUpperCase().indexOf('OPPORTUNITIES') >= 0) break;
      if (col1 && col1 !== 'Performance Metric / KPI') {
        // Move current to prior, clear current and notes
        var currentVal = ws.getRange(i + 1, 3).getValue();
        ws.getRange(i + 1, 2).setValue(currentVal); // prior = last month's current
        ws.getRange(i + 1, 3).setValue(''); // clear current
        ws.getRange(i + 1, 5).setValue(''); // clear notes
      }
    }

    // Clear opportunities and plans
    for (var i = 4; i < data.length; i++) {
      var col1 = String(data[i][0] || '').trim();
      if (col1.toUpperCase().indexOf('OPPORTUNITIES') >= 0) {
        for (var j = 1; j <= 3; j++) {
          ws.getRange(i + 1 + j, 2).setValue('');
        }
      }
      if (col1.toUpperCase().indexOf('PLAN FOR') >= 0) {
        for (var j = 2; j <= 5; j++) {
          for (var c = 2; c <= 5; c++) {
            ws.getRange(i + 1 + j, c).setValue('');
          }
        }
        break;
      }
    }
  }

  // Clear Brief to DON
  var briefSheet = newSS.getSheetByName('Brief to DON');
  if (briefSheet) {
    for (var r = 5; r <= 16; r++) {
      for (var c = 2; c <= 4; c++) {
        briefSheet.getRange(r, c).setValue('');
        briefSheet.getRange(r, c).setBackground('#FFF9E6');
      }
    }
  }

  ui.alert('New workbook created: "QAPI Monthly Input - ' + monthName + '"\n\nPrior month data has been carried forward. Current month is blank and ready for input.\n\nOpen it from Google Drive.');
}
