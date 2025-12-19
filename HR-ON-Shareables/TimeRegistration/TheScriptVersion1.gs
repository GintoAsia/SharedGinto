// ==========================================
//       HR-ON AUTOMATION (3-BREAK EDITION)
// ==========================================

// --- CONFIGURATION: SHEET NAMES ---
const SETTINGS_SHEET = 'Settings';
const EMPLOYEE_SHEET = 'Employee_Database';
const CALENDAR_SHEET = 'Planning_Calendar';
const SQL_SHEET = 'SQL_Output';

// --- MENU CREATION ---
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('HR-ON Automation')
    .addItem('🚀 INITIALIZE SHEET (Run First)', 'initializeSheetStructure')
    .addSeparator()
    .addSubMenu(ui.createMenu('Step 1: Data Setup')
      .addItem('1. Refresh Employees (Sorted by Dept)', 'refreshEmployeeData')
      .addItem('2. Refresh Dropdowns (Internal)', 'updateSettingsDropdowns'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Step 2: Scheduling')
      .addItem('Create/Reset Calendar', 'setupCalendar')
      .addSeparator()
      .addItem('📅 Bulk Assign Shifts (Pattern-Based)', 'bulkAssignShifts')
      .addItem('🏢 Assign Shifts by Department', 'assignByDepartment')
      .addItem('📋 Copy Week Pattern', 'copyWeekPattern')
      .addItem('🗑️ Clear Calendar', 'clearCalendar')
      .addSeparator()
      .addItem('✅ Validate Calendar Shifts', 'validateCalendarShifts'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Step 3: Export')
      .addItem('Process Calendar to SQL', 'processCalendar')
      .addItem('Email SQL to IT', 'emailSql'))
    .addToUi();
}

// ==========================================
//      CORE FUNCTIONS
// ==========================================

// --- 1. REFRESH EMPLOYEES (WITH DEPARTMENTS) ---
function refreshEmployeeData() {
  const ui = SpreadsheetApp.getUi();
  try {
    const config = getConfiguration(); 
    if (!config.USERS_API_URL) { ui.alert('Error: Missing "User API URL" in Settings.'); return; }

    const token = getAccessToken(config);
    
    // A. Fetch Company ID
    let companyId = "";
    if (config.COMPANY_API_URL) {
      try {
        const cResp = UrlFetchApp.fetch(config.COMPANY_API_URL, { headers: { 'Authorization': `Bearer ${token}` }, muteHttpExceptions:true });
        const cJson = JSON.parse(cResp.getContentText());
        const cData = Array.isArray(cJson) ? cJson[0] : (cJson.data ? (Array.isArray(cJson.data) ? cJson.data[0] : cJson.data) : cJson);
        if (cData && cData.id) companyId = cData.id;
      } catch (e) { console.warn("Company fetch failed."); }
    }
    if (!companyId) companyId = config['Default Company ID'] || "";

    // B. Fetch Employees
    const resp = UrlFetchApp.fetch(config.USERS_API_URL, { headers: { 'Authorization': `Bearer ${token}` }, muteHttpExceptions:true });
    const json = JSON.parse(resp.getContentText());
    
    let data = json;
    if (!Array.isArray(data)) {
      if (data.data) data = data.data; else if (data.items) data = data.items; else if (data.employees) data = data.employees;
    }
    if (!Array.isArray(data)) throw new Error("API Response was not a list.");

    // C. Process Data (Extract Department)
    const values = data.map(u => {
      const sys = u.systemFields || {};
      // Name
      let name = (sys.firstName || sys.lastName) ? `${sys.firstName||''} ${sys.lastName||''}`.trim() : (u.name || "Unknown");
      
      // Department Logic: API usually returns an array of departments
      let departmentName = "No Department";
      if (u.departments && Array.isArray(u.departments) && u.departments.length > 0) {
        departmentName = u.departments[0].name || u.departments[0].departmentId || "Unknown"; 
      }

      return [u.id, name, departmentName, companyId || u.companyId || ""];
    });

    // D. Sort: 1. Department, 2. Name
    values.sort((a, b) => {
      if (a[2] < b[2]) return -1; // Dept A-Z
      if (a[2] > b[2]) return 1;
      if (a[1] < b[1]) return -1; // Name A-Z
      if (a[1] > b[1]) return 1;
      return 0;
    });
    
    // Add Header
    values.unshift(['EmployeeID', 'EmployeeName', 'Department', 'CompanyID']);

    // E. Write & Style
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EMPLOYEE_SHEET);
    if (!sheet) return;

    sheet.clear();
    if (sheet.getFilter()) sheet.getFilter().remove();

    const range = sheet.getRange(1,1,values.length,4);
    range.setValues(values);
    
    // Table Styling
    sheet.setHiddenGridlines(true);
    sheet.getRange(1, 1, 1, 4)
         .setFontWeight("bold")
         .setBackground("#134f5c") // Teal Header
         .setFontColor("white")
         .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    if (values.length > 1) {
      const dataRange = sheet.getRange(2, 1, values.length - 1, 4);
      dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.TEAL);
      dataRange.setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
    }

    // Add Sortable Filter
    range.createFilter();
    
    sheet.autoResizeColumns(1, 4);
    // Hide unused space
    const maxCols = sheet.getMaxColumns();
    if (maxCols > 4) sheet.hideColumns(5, maxCols - 4);

    ui.alert(`Success! Refreshed ${values.length-1} employees sorted by Department.`);

  } catch (e) {
    ui.alert("Error: " + e.message);
  }
}

// --- 2. UPDATE SETTINGS DROPDOWNS (FIXED & SAFETY CHECKED) ---
function updateSettingsDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  if (!settings) return;

  // 1. Define Sources (Yellow Table)
  const presenceRange = settings.getRange("A13:A30");
  const breakRange = settings.getRange("C13:C30"); 
  const projectRange = settings.getRange("E13:E30"); 

  // 2. Define Target Columns for DROPDOWNS (Green Table)
  // D = Presence
  // G = Break 1 Type
  // J = Break 2 Type
  // M = Break 3 Type
  // N = Project
  
  const applyDropdown = (sourceRange, targetColLetter) => {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange)
      .setAllowInvalid(true)
      .build();
    // Apply to rows 35-100
    settings.getRange(`${targetColLetter}35:${targetColLetter}100`).setDataValidation(rule);
  };

  // 3. Define Target Columns for TIMES (To Clear Dropdowns)
  // B, C = Shift Start/End
  // E, F = Break 1 Start/End
  // H, I = Break 2 Start/End
  // K, L = Break 3 Start/End
  const timeCols = ["B", "C", "E", "F", "H", "I", "K", "L"];

  // --- EXECUTE FIX ---
  
  // A. Clean the Time Columns (Remove accidental dropdowns & set Time Format)
  timeCols.forEach(col => {
    const range = settings.getRange(`${col}35:${col}100`);
    range.clearDataValidations(); // Remove the dropdown
    range.setNumberFormat("HH:mm"); // Force Time format
  });

  // B. Apply The Correct Dropdowns
  applyDropdown(presenceRange, 'D'); // Presence
  applyDropdown(breakRange, 'G');    // Break 1 Type
  applyDropdown(breakRange, 'J');    // Break 2 Type
  applyDropdown(breakRange, 'M');    // Break 3 Type
  applyDropdown(projectRange, 'N');  // Project

  SpreadsheetApp.getUi().alert("✅ Fixed!\n\n- Time columns are now clean (HH:mm).\n- Dropdowns are only on the 'Type' columns.");
}

// --- 3. SETUP CALENDAR ---
function setupCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let cal = ss.getSheetByName(CALENDAR_SHEET);
  
  if (!cal) cal = ss.insertSheet(CALENDAR_SHEET);
  else { cal.clear(); if(cal.getFilter()) cal.getFilter().remove(); }
  
  cal.setTabColor("#f1c232");

  const startStr = ui.prompt("Start Date (YYYY-MM-01)", "2025-08-01", ui.ButtonSet.OK).getResponseText();
  const endStr = ui.prompt("End Date (YYYY-MM-01)", "2026-07-01", ui.ButtonSet.OK).getResponseText();
  const start = new Date(startStr); 
  const end = new Date(endStr);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) { ui.alert("Invalid Date"); return; }

  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const shiftRange = settings.getRange("A35:A100");
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(shiftRange).setAllowInvalid(true).build();

  const empSheet = ss.getSheetByName(EMPLOYEE_SHEET);
  if (empSheet.getLastRow() < 2) { ui.alert("No employees found."); return; }
  const employees = empSheet.getRange(2, 2, empSheet.getLastRow()-1, 1).getValues().flat();

  const headers = ['Date'].concat(employees);
  const dates = [];
  let ptr = new Date(start);
  while(ptr <= end) {
    dates.push(Utilities.formatDate(ptr, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd"));
    ptr.setDate(ptr.getDate() + 1);
  }
  
  const reqRows = dates.length + 10;
  const reqCols = headers.length + 5;
  if (cal.getMaxRows() < reqRows) cal.insertRowsAfter(cal.getMaxRows(), reqRows - cal.getMaxRows());
  if (cal.getMaxColumns() < reqCols) cal.insertColumnsAfter(cal.getMaxColumns(), reqCols - cal.getMaxColumns());

  cal.setHiddenGridlines(true);

  cal.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1c4587').setFontColor('white')
      .setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  const dateRows = dates.map(d => [d]);
  cal.getRange(2, 1, dateRows.length, 1).setValues(dateRows)
      .setFontWeight('bold').setBackground('#f3f3f3')
      .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);

  const gridRange = cal.getRange(2, 2, dates.length, employees.length);
  gridRange.setDataValidation(rule)
      .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID); 
  
  cal.setFrozenColumns(1);
  cal.setFrozenRows(1);
  cal.autoResizeColumn(1);
  cal.setColumnWidths(2, employees.length, 120);
  
  const maxCols = cal.getMaxColumns();
  if (maxCols > headers.length + 1) cal.hideColumns(headers.length + 2, maxCols - (headers.length + 1));

  ui.alert("Calendar created!\n\nTip: Select the grid > Data Validation > Enable 'Chip' & 'Multi-select'.");
}

// ==========================================
//      SHIFT VALIDATION & HELPERS
// ==========================================

function getValidShifts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  if (!settings) return [];
  const shiftData = settings.getRange("A35:A100").getValues().flat();
  return shiftData.filter(s => s !== "" && s !== null && s !== undefined);
}

function validateCalendarShifts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  
  if (!cal || cal.getLastRow() < 2) { ui.alert('Error', 'No calendar found.', ui.ButtonSet.OK); return; }
  
  const validShifts = getValidShifts();
  if (validShifts.length === 0) { ui.alert('Warning', 'No shifts defined in Settings.', ui.ButtonSet.OK); return; }
  
  const validShiftSet = new Set(validShifts.map(s => s.toString().trim()));
  const data = cal.getDataRange().getValues();
  const employeeHeaders = data[0].slice(1);
  const invalidShifts = new Map();
  let totalInvalidCells = 0;
  
  for (let r = 1; r < data.length; r++) {
    const dateStr = data[r][0];
    for (let c = 1; c < data[r].length; c++) {
      const cell = data[r][c];
      const empName = employeeHeaders[c - 1] || `Column ${c + 1}`;
      if (cell && cell !== "") {
        const shifts = cell.toString().split(',').map(s => s.trim());
        shifts.forEach(shiftName => {
          if (shiftName && !validShiftSet.has(shiftName)) {
            totalInvalidCells++;
            if (!invalidShifts.has(shiftName)) invalidShifts.set(shiftName, []);
            invalidShifts.get(shiftName).push({ date: dateStr, employee: empName, row: r + 1, col: c + 1 });
          }
        });
      }
    }
  }
  
  if (invalidShifts.size === 0) {
    ui.alert('✅ Validation Passed', 'All shifts in the calendar are valid.', ui.ButtonSet.OK);
  } else {
    highlightInvalidShiftCells(cal, invalidShifts);
    ui.alert('⚠️ Validation Failed', `Found ${invalidShifts.size} invalid shift types in ${totalInvalidCells} cells. Cells are highlighted red.`, ui.ButtonSet.OK);
  }
}

function highlightInvalidShiftCells(cal, invalidShifts) {
  invalidShifts.forEach((locations, shiftName) => {
    locations.forEach(loc => {
      try { cal.getRange(loc.row, loc.col).setBackground('#ffcccc').setNote(`Invalid: "${shiftName}"`); } catch (e) {}
    });
  });
}

function clearShiftHighlights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  if (!cal || cal.getLastRow() < 2) return;
  const gridRange = cal.getRange(2, 2, cal.getLastRow() - 1, cal.getLastColumn() - 1);
  gridRange.setBackground(null);
  gridRange.clearNote();
}

// ==========================================
//      BULK SCHEDULING (ABBREVIATED)
// ==========================================
// Note: Keeping logic identical to previous, just ensuring dependencies exist.

function bulkAssignShifts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  if (!cal) return;
  
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const shiftData = settings.getRange("A35:A100").getValues().flat().filter(s => s !== "");
  const headers = cal.getRange(1, 1, 1, cal.getLastColumn()).getValues()[0];
  const employees = headers.slice(1).filter(e => e !== "");

  const htmlTemplate = HtmlService.createHtmlOutput(buildBulkAssignDialog(shiftData, employees))
    .setWidth(500).setHeight(600);
  ui.showModalDialog(htmlTemplate, '📅 Bulk Assign Shifts');
}

// (HTML Builder functions remain same as previously provided, omitted for brevity but required for execution.
// If you need the full HTML string again, I can include it, but the key changes are in the Export/Setup logic below)
function buildBulkAssignDialog(shifts, employees) {
  const shiftOptions = shifts.map(s => `<option value="${s}">${s}</option>`).join('');
  const employeeCheckboxes = employees.map((e, i) => `<label class="employee-item"><input type="checkbox" name="emp" value="${i}" checked> ${e}</label>`).join('');
  return `<!DOCTYPE html><html><head><style>body{font-family:Arial,sans-serif;padding:15px}.section{margin-bottom:20px;padding:15px;border:1px solid #ddd;border-radius:8px}.section-title{font-weight:bold;color:#1c4587;margin-bottom:10px}select,input[type="date"]{width:100%;padding:8px;margin:5px 0}.employee-list{max-height:150px;overflow-y:auto;border:1px solid #ddd;padding:10px}.btn-row{display:flex;gap:10px;margin-top:10px}button{padding:10px 20px;background:#1c4587;color:white;border:none;border-radius:4px;cursor:pointer}</style></head><body><div class="section"><div class="section-title">1. Select Shift</div><select id="shift">${shiftOptions}</select></div><div class="section"><div class="section-title">2. Select Days</div><div style="display:flex;gap:5px"><label><input type="checkbox" name="day" value="1" checked>Mon</label><label><input type="checkbox" name="day" value="2" checked>Tue</label><label><input type="checkbox" name="day" value="3" checked>Wed</label><label><input type="checkbox" name="day" value="4" checked>Thu</label><label><input type="checkbox" name="day" value="5" checked>Fri</label><label><input type="checkbox" name="day" value="6">Sat</label><label><input type="checkbox" name="day" value="0">Sun</label></div></div><div class="section"><div class="section-title">3. Range</div>Start: <input type="date" id="startDate"><br>End: <input type="date" id="endDate"></div><div class="section"><div class="section-title">4. Employees</div><div class="employee-list">${employeeCheckboxes}</div></div><div class="btn-row"><button onclick="apply()">Apply</button></div><script>function apply(){const shift=document.getElementById('shift').value;const days=Array.from(document.querySelectorAll('input[name="day"]:checked')).map(c=>parseInt(c.value));const emps=Array.from(document.querySelectorAll('input[name="emp"]:checked')).map(c=>parseInt(c.value));const sd=document.getElementById('startDate').value;const ed=document.getElementById('endDate').value;google.script.run.withSuccessHandler(()=>google.script.host.close()).applyBulkShifts(shift,days,emps,sd,ed);}</script></body></html>`;
}

function applyBulkShifts(shift, days, empIndices, startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  const dateColumn = cal.getRange(1, 1, cal.getLastRow(), 1).getValues().flat();
  let startRow = 2; let endRow = dateColumn.length;

  if (startDate) { const sd = new Date(startDate); for (let r = 2; r <= dateColumn.length; r++) { if (new Date(dateColumn[r-1]) >= sd) { startRow = r; break; } } }
  if (endDate) { const ed = new Date(endDate); for (let r = dateColumn.length; r >= 2; r--) { if (new Date(dateColumn[r-1]) <= ed) { endRow = r; break; } } }

  let count = 0;
  for (const empIdx of empIndices) {
    const col = empIdx + 2; 
    for (let row = startRow; row <= endRow; row++) {
      const rowDate = new Date(dateColumn[row - 1]);
      if (days.includes(rowDate.getDay())) {
        const cell = cal.getRange(row, col);
        const currentVal = cell.getValue();
        if (currentVal && currentVal !== "") {
          const existing = currentVal.toString().split(',').map(s => s.trim());
          if (!existing.includes(shift)) cell.setValue(existing.concat(shift).join(', '));
        } else { cell.setValue(shift); }
        count++;
      }
    }
  }
  return count;
}

function assignByDepartment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  const empSheet = ss.getSheetByName(EMPLOYEE_SHEET);
  if (!cal) return;

  const empData = empSheet.getDataRange().getValues().slice(1);
  const departments = [...new Set(empData.map(r => r[2]).filter(d => d))];
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const shiftData = settings.getRange("A35:A100").getValues().flat().filter(s => s !== "");

  const html = `<!DOCTYPE html><html><body style="font-family:Arial;padding:10px">
    <b>Dept:</b><br><select id="d">${departments.map(d=>`<option>${d}</option>`).join('')}</select><br><br>
    <b>Shift:</b><br><select id="s">${shiftData.map(s=>`<option>${s}</option>`).join('')}</select><br><br>
    <b>Days:</b><br><label><input type="checkbox" name="day" value="1" checked>Mon-Fri</label> (Logic simplified for UI)<br><br>
    <button onclick="run()">Apply</button>
    <script>function run(){google.script.run.withSuccessHandler(()=>google.script.host.close()).applyDepartmentShifts(document.getElementById('d').value,document.getElementById('s').value,[1,2,3,4,5]);}</script></body></html>`;
  
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setHeight(300), 'Dept Assign');
}

function applyDepartmentShifts(department, shift, days) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  const empSheet = ss.getSheetByName(EMPLOYEE_SHEET);
  const empData = empSheet.getDataRange().getValues().slice(1);
  const deptEmployees = empData.filter(r => r[2] === department).map(r => r[1]);
  const headers = cal.getRange(1, 1, 1, cal.getLastColumn()).getValues()[0];
  const empIndices = [];
  headers.slice(1).forEach((emp, idx) => { if (deptEmployees.includes(emp)) empIndices.push(idx); });
  if (empIndices.length > 0) applyBulkShifts(shift, days, empIndices, '', '');
}

function copyWeekPattern() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  const response = ui.prompt('Date of Monday to Copy (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const sourceMonday = new Date(response.getResponseText());
  
  const dateColumn = cal.getRange(1, 1, cal.getLastRow(), 1).getValues().flat();
  let sourceStartRow = -1;
  const sourceStr = Utilities.formatDate(sourceMonday, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  
  for (let r = 2; r <= dateColumn.length; r++) {
    if (Utilities.formatDate(new Date(dateColumn[r-1]), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') === sourceStr) {
      sourceStartRow = r; break;
    }
  }
  
  if (sourceStartRow === -1) { ui.alert("Date not found"); return; }
  
  const employees = cal.getLastColumn() - 1;
  const sourceData = cal.getRange(sourceStartRow, 2, 7, employees).getValues(); // Copy 7 days
  
  let weeks = 0;
  for (let row = 2; row <= cal.getLastRow(); row++) {
    const d = new Date(dateColumn[row-1]);
    if (d.getDay() === 1 && row !== sourceStartRow) {
      const daysLeft = Math.min(7, cal.getLastRow() - row + 1);
      cal.getRange(row, 2, daysLeft, employees).setValues(sourceData.slice(0, daysLeft));
      weeks++;
    }
  }
  ui.alert(`Copied to ${weeks} weeks.`);
}

function clearCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (SpreadsheetApp.getUi().alert('Clear all shifts?', SpreadsheetApp.getUi().ButtonSet.YES_NO) === SpreadsheetApp.getUi().Button.YES) {
    const cal = ss.getSheetByName(CALENDAR_SHEET);
    cal.getRange(2, 2, cal.getLastRow()-1, cal.getLastColumn()-1).clearContent();
  }
}

// --- 4. PROCESS TO SQL (UPDATED FOR 3 BREAKS + SQL PRE-GENERATION) ---
function processCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  const sqlSheet = ss.getSheetByName(SQL_SHEET);
  
  // 1. Employee Map
  const empMap = new Map(); 
  const empData = ss.getSheetByName(EMPLOYEE_SHEET).getDataRange().getValues();
  empData.slice(1).forEach(r => empMap.set(r[1], {id:r[0], co:r[3]}));

  // 2. ID Lookup (A-F)
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const refData = settings.getRange("A13:F30").getValues(); 
  const idLookup = new Map();
  refData.forEach(r => {
    if(r[0]) idLookup.set(r[0], r[1]); 
    if(r[2]) idLookup.set(r[2], r[3]); 
    if(r[4]) idLookup.set(r[4], r[5]); 
  });

  // 3. Shift Map (UPDATED FOR NEW COLUMNS)
  // Columns A-N (14 cols)
  const shiftRows = settings.getRange("A35:N100").getValues();
  const shiftMap = new Map();
  
  shiftRows.forEach(r => {
    if(r[0]) {
      shiftMap.set(r[0], {
        // Presence
        pStart: r[1], 
        pEnd: r[2], 
        pID: idLookup.get(r[3]) || "", 
        
        // Break 1
        b1Start: r[4], b1End: r[5], b1ID: idLookup.get(r[6]) || "",
        
        // Break 2 (New)
        b2Start: r[7], b2End: r[8], b2ID: idLookup.get(r[9]) || "",

        // Break 3 (New)
        b3Start: r[10], b3End: r[11], b3ID: idLookup.get(r[12]) || "",

        // Project (Moved to Col N / Index 13)
        prID: idLookup.get(r[13]) || ""
      });
    }
  });

  const data = cal.getDataRange().getValues();
  const employeeHeaders = data[0].slice(1);
  const output = [];
  const tz = ss.getSpreadsheetTimeZone();

  // Helper to format timestamps for SQL
  const fmt = (dateStr, timeVal) => {
    if (!timeVal) return null;
    const d = new Date(dateStr);
    let tStr = timeVal;
    // Check if timeVal is Date object or string
    if (timeVal instanceof Date) {
      tStr = Utilities.formatDate(timeVal, tz, 'HH:mm:ss');
    }
    const datePart = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    return `${datePart} ${tStr}+00`;
  };

  // Loop Calendar
  for(let r=1; r<data.length; r++) {
    const dateStr = data[r][0]; 
    if (!dateStr) continue;

    for(let c=1; c<data[r].length; c++) {
      const cell = data[r][c];
      const empName = employeeHeaders[c-1];
      const emp = empMap.get(empName);
      
      if(cell && cell !== "" && emp) {
        const shifts = cell.toString().split(',').map(s => s.trim());
        
        shifts.forEach(shiftName => {
          if(shiftMap.has(shiftName)) {
            const s = shiftMap.get(shiftName);
            
            // --- GENERATE FULL SQL HERE ---
            // 1. Presence (CTE)
            let sql = `with new_time_registration as (
 insert into time_registration_presence (
   time_registration_presence_type_id, start_date, end_date, user_id, company_id
 ) values (
   '${s.pID}', '${fmt(dateStr, s.pStart)}', '${fmt(dateStr, s.pEnd)}', '${emp.id}', '${emp.co}'
 ) returning id
)`;

            // 2. Prepare Entries (Breaks & Projects)
            let values = [];

            // Helper to push break values if valid
            const addBreak = (start, end, typeId) => {
               if(start && end && typeId) {
                 values.push(`('BREAK', '${fmt(dateStr, start)}', '${fmt(dateStr, end)}', '${emp.id}', '${emp.co}', '${typeId}', null, (select id from new_time_registration))`);
               }
            };

            // Check Break 1
            addBreak(s.b1Start, s.b1End, s.b1ID);
            // Check Break 2
            addBreak(s.b2Start, s.b2End, s.b2ID);
            // Check Break 3
            addBreak(s.b3Start, s.b3End, s.b3ID);

            // Check Project
            if(s.prID) {
              values.push(`('PROJECT', '${fmt(dateStr, s.pStart)}', '${fmt(dateStr, s.pEnd)}', '${emp.id}', '${emp.co}', null, '${s.prID}', (select id from new_time_registration))`);
            }

            // 3. Combine
            if (values.length > 0) {
              sql += `
insert into time_registration_entry (
 time_registration_entry_type, start_date, end_date, user_id, company_id, 
 time_registration_break_type_id, time_registration_project_id, time_registration_presence_id
) values 
${values.join(',\n')};`;
            } else {
              sql += `; -- No breaks or projects`;
            }

            // Push to Output Array
            output.push([
              emp.id, empName, Utilities.formatDate(new Date(dateStr), tz, 'yyyy-MM-dd'), shiftName, sql, 'Pending'
            ]);
          }
        });
      }
    }
  }

  // Write to SQL Sheet
  sqlSheet.clear();
  // Simplified Headers for clarity
  const headers = ['UserID', 'Employee Name', 'Date', 'Shift', 'Generated SQL Query', 'Status'];
  
  if(output.length > 0) {
    const range = sqlSheet.getRange(1,1,output.length + 1, headers.length);
    range.setValues([headers, ...output]);
    
    sqlSheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#cc0000').setFontColor('white');
    range.setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
    
    // Set Column Widths: SQL column wide
    sqlSheet.setColumnWidth(5, 500); 
    
    SpreadsheetApp.getUi().alert(`Processed ${output.length} shifts. Check SQL_Output.`);
  } else {
    SpreadsheetApp.getUi().alert("No shifts found in calendar.");
  }
}

// --- 5. EMAIL SQL (SIMPLIFIED) ---
function emailSql() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Email SQL", "Enter recipient email:", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  const email = resp.getResponseText().trim();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SQL_SHEET);
  const data = sheet.getDataRange().getValues(); // Headers + Data
  
  const sqls = [];
  const statusCells = [];

  // Col 5 is SQL (Index 4), Col 6 is Status (Index 5)
  for(let i=1; i<data.length; i++) {
    const row = data[i];
    if(row[5] === 'Pending') {
      sqls.push(`-- Entry for ${row[1]} on ${row[2]} (${row[3]})\n${row[4]}`);
      statusCells.push(sheet.getRange(i+1, 6)); // Store reference to update later
    }
  }

  if(sqls.length > 0) {
    try {
      MailApp.sendEmail(email, "HR-ON SQL Import Batch", sqls.join("\n\n------------------------\n\n"));
      statusCells.forEach(c => c.setValue('Sent'));
      ui.alert(`Sent ${sqls.length} queries to ${email}.`);
    } catch(e) {
      ui.alert("Error sending email: " + e.message);
    }
  } else {
    ui.alert("No 'Pending' rows found to send.");
  }
}

// ==========================================
//      FACTORY RESET (CORRECTED)
// ==========================================
function initializeSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  if (ui.alert('Reset Sheet?', 'This will reformat columns for the 3-Break System. Continue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  // 1. Settings
  let s = ss.getSheetByName(SETTINGS_SHEET) || ss.insertSheet(SETTINGS_SHEET);
  s.clear();
  s.setHiddenGridlines(true); 
  s.setTabColor("#4285f4");
  
  // A. System Config
  s.getRange("A1:B1").merge().setValue("SYSTEM CONFIGURATION").setFontWeight("bold").setBackground("#1c4587").setFontColor("white").setHorizontalAlignment("center");
  const defaults = [
    ["CLIENT_ID", ""], ["CLIENT_SECRET", ""], 
    ["TOKEN_URL", "https://auth.hr-on.com/oauth2/token"],
    ["USERS_API_URL", "https://api.hr-on.com/v1/staff/employees?size=1000"],
    ["COMPANY_API_URL", "https://api.hr-on.com/v1/staff/company"],
    ["Default Company ID", ""]
  ];
  
  // Fix 1: Apply banding last or separately to avoid chaining errors
  const configRange = s.getRange(2, 1, defaults.length, 2);
  configRange.setValues(defaults);
  configRange.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE);
  
  // B. Reference Data
  s.getRange("A11:F11").merge().setValue("REFERENCE DATA (PASTE IDs HERE)").setFontWeight("bold").setBackground("#bf9000").setFontColor("white").setHorizontalAlignment("center");
  s.getRange("A12:F12").setValues([["PRESENCE NAME","ID","BREAK NAME","ID","PROJECT NAME","ID"]])
   .setFontWeight("bold").setBackground("#f1c232");
   
  // Fix 2: Split Range definition from styling to prevent "setBorder is not a function" error
  const refRange = s.getRange("A13:F30");
  refRange.applyRowBanding(SpreadsheetApp.BandingTheme.YELLOW);
  refRange.setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);

  // C. Shift Definitions (EXPANDED FOR 3 BREAKS)
  s.getRange("A33:N33").merge().setValue("SHIFT DEFINITIONS (3 BREAK SUPPORT)").setFontWeight("bold").setBackground("#38761d").setFontColor("white").setHorizontalAlignment("center");
  
  const greenHeaders = [["Shift Name","Start","End","Pres Type","Brk1 Start","Brk1 End","Brk1 Type","Brk2 Start","Brk2 End","Brk2 Type","Brk3 Start","Brk3 End","Brk3 Type","Project"]];
  s.getRange("A34:N34").setValues(greenHeaders)
   .setFontWeight("bold").setBackground("#6aa84f").setFontColor("white")
   .setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  const shiftRange = s.getRange("A35:N100");
  shiftRange.applyRowBanding(SpreadsheetApp.BandingTheme.GREEN);
  shiftRange.setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
  
  // Example Row
  s.getRange("A35:N35").setValues([["Morning A","08:00","16:00","(Select)","12:00","12:30","(Select)","","","","","","","(Select)"]]);

  // Adjust Columns
  s.setColumnWidth(1, 120); // Name
  // Times are small
  [2,3,5,6,8,9,11,12].forEach(c => s.setColumnWidth(c, 60)); 
  // Dropdowns are wider
  [4,7,10,13,14].forEach(c => s.setColumnWidth(c, 100));

  // 2. Clean Up other sheets
  let e = ss.getSheetByName(EMPLOYEE_SHEET) || ss.insertSheet(EMPLOYEE_SHEET); e.clear(); e.setTabColor("#0097a7");
  let q = ss.getSheetByName(SQL_SHEET) || ss.insertSheet(SQL_SHEET); q.clear(); q.setTabColor("#ea4335");
  
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  ui.alert("Reset Complete! Columns expanded for 3 Breaks.");
}

// --- UTILS ---
function getConfiguration() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
  const data = sheet.getRange("A2:B9").getValues();
  const config = {};
  data.forEach(r => { if(r[0]) config[String(r[0]).trim().toUpperCase().replace(" ", "_")] = r[1]; });
  config['DEFAULT_COMPANY_ID'] = config['DEFAULT_COMPANY_ID'] || "";
  return config;
}

function getAccessToken(config) {
  if (!config.CLIENT_ID) throw new Error("🛑 CLIENT_ID missing.");
  if (!config.CLIENT_SECRET) throw new Error("🛑 CLIENT_SECRET missing.");
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('TOKEN');
  if (saved) { const t = JSON.parse(saved); if (t.exp > Date.now()) return t.val; }
  const auth = Utilities.base64Encode(`${config.CLIENT_ID}:${config.CLIENT_SECRET}`);
  const resp = UrlFetchApp.fetch(config.TOKEN_URL, { method: 'post', payload: { grant_type: 'client_credentials' }, headers: { 'Authorization': 'Basic ' + auth } });
  const json = JSON.parse(resp.getContentText());
  props.setProperty('TOKEN', JSON.stringify({ val: json.access_token, exp: Date.now() + (json.expires_in-300)*1000 }));
  return json.access_token;
}
