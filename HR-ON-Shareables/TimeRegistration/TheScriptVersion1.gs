// ==========================================
//       HR-ON AUTOMATION (DASHBOARD EDITION)
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
      .addItem('🗑️ Clear Calendar', 'clearCalendar'))
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
        // Grab the first department name found
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
    // Remove existing filter if any
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

// --- 2. UPDATE SETTINGS DROPDOWNS (COMPACT) ---
function updateSettingsDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  if (!settings) return;

  // Reference Data is now in Columns A-F (no gaps)
  // Presence: A
  // Break: C
  // Project: E
  const presenceRange = settings.getRange("A13:A30");
  const breakRange = settings.getRange("C13:C30"); 
  const projectRange = settings.getRange("E13:E30"); 

  // Shift Definitions start at Row 35
  const apply = (sourceRange, targetCol) => {
    const rule = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRange).setAllowInvalid(true).build();
    settings.getRange(`${targetCol}35:${targetCol}100`).setDataValidation(rule);
  };

  apply(presenceRange, 'F'); // Presence -> Col F
  apply(breakRange, 'G');    // Break -> Col G
  apply(projectRange, 'H');  // Project -> Col H
  
  SpreadsheetApp.getUi().alert("Dropdowns updated! Check your Shift Definitions table.");
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

  const headers = ['Employee'];
  const dates = [];
  let ptr = new Date(start);
  while(ptr <= end) {
    headers.push(Utilities.formatDate(ptr, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd"));
    dates.push(new Date(ptr));
    ptr.setDate(ptr.getDate() + 1);
  }
  
  // Resize Sheet
  const reqRows = employees.length + 10;
  const reqCols = headers.length + 5;
  if (cal.getMaxRows() < reqRows) cal.insertRowsAfter(cal.getMaxRows(), reqRows - cal.getMaxRows());
  if (cal.getMaxColumns() < reqCols) cal.insertColumnsAfter(cal.getMaxColumns(), reqCols - cal.getMaxColumns());

  cal.setHiddenGridlines(true);

  // Headers
  cal.getRange(1, 1, 1, headers.length).setValues([headers])
     .setFontWeight('bold').setBackground('#1c4587').setFontColor('white')
     .setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Rows
  const empRows = employees.map(e => [e]);
  cal.getRange(2, 1, empRows.length, 1).setValues(empRows)
     .setFontWeight('bold').setBackground('#f3f3f3')
     .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);

  // Grid
  const gridRange = cal.getRange(2, 2, empRows.length, dates.length);
  gridRange.setDataValidation(rule)
     .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID); 
  
  cal.setFrozenColumns(1);
  cal.setFrozenRows(1);
  cal.autoResizeColumn(1);
  cal.setColumnWidths(2, dates.length, 100);
  
  // Hide unused columns
  const maxCols = cal.getMaxColumns();
  if (maxCols > headers.length + 1) cal.hideColumns(headers.length + 2, maxCols - (headers.length + 1));

  ui.alert("Calendar created!\n\nTip: Select the grid > Data Validation > Enable 'Chip' & 'Multi-select'.");
}

// ==========================================
//      BULK SCHEDULING FUNCTIONS
// ==========================================

/**
 * Bulk assign shifts using a pattern-based dialog.
 * Users can specify which days of the week to fill and the shift to assign.
 */
function bulkAssignShifts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  
  if (!cal || cal.getLastRow() < 2) {
    ui.alert('Error', 'Please create a calendar first using "Create/Reset Calendar".', ui.ButtonSet.OK);
    return;
  }
  
  // Get available shifts from Settings
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const shiftData = settings.getRange("A35:A100").getValues().flat().filter(s => s !== "");
  if (shiftData.length === 0) {
    ui.alert('Error', 'No shifts defined in Settings. Please define shifts first.', ui.ButtonSet.OK);
    return;
  }
  
  // Get employees from the calendar
  const employees = cal.getRange(2, 1, cal.getLastRow() - 1, 1).getValues().flat().filter(e => e !== "");
  if (employees.length === 0) {
    ui.alert('Error', 'No employees found in the calendar.', ui.ButtonSet.OK);
    return;
  }
  
  // Build HTML dialog for user-friendly selection
  const htmlTemplate = HtmlService.createHtmlOutput(buildBulkAssignDialog(shiftData, employees))
    .setWidth(500)
    .setHeight(600);
  ui.showModalDialog(htmlTemplate, '📅 Bulk Assign Shifts');
}

/**
 * Builds the HTML content for the bulk assign dialog
 */
function buildBulkAssignDialog(shifts, employees) {
  const shiftOptions = shifts.map(s => `<option value="${s}">${s}</option>`).join('');
  const employeeCheckboxes = employees.map((e, i) => 
    `<label class="employee-item"><input type="checkbox" name="emp" value="${i}" checked> ${e}</label>`
  ).join('');
  
  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; padding: 15px; }
    .section { margin-bottom: 20px; padding: 15px; border: 1px solid #ddd; border-radius: 8px; }
    .section-title { font-weight: bold; color: #1c4587; margin-bottom: 10px; }
    select, input[type="date"] { width: 100%; padding: 8px; margin: 5px 0; border: 1px solid #ccc; border-radius: 4px; }
    .day-checkboxes { display: flex; flex-wrap: wrap; gap: 10px; }
    .day-checkboxes label { 
      padding: 8px 12px; background: #e8f0fe; border-radius: 4px; cursor: pointer;
      border: 2px solid transparent; transition: all 0.2s;
    }
    .day-checkboxes label:has(input:checked) { background: #1c4587; color: white; }
    .day-checkboxes input { display: none; }
    .employee-list { max-height: 150px; overflow-y: auto; border: 1px solid #ddd; padding: 10px; border-radius: 4px; }
    .employee-item { display: block; padding: 3px 0; }
    .btn-row { display: flex; gap: 10px; margin-top: 10px; }
    button { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; }
    .btn-primary { background: #1c4587; color: white; }
    .btn-secondary { background: #f1f3f4; color: #333; }
    .btn-select { background: #38761d; color: white; font-size: 12px; padding: 5px 10px; }
  </style>
</head>
<body>
  <div class="section">
    <div class="section-title">1. Select Shift</div>
    <select id="shift">${shiftOptions}</select>
  </div>
  
  <div class="section">
    <div class="section-title">2. Select Days of Week</div>
    <div class="day-checkboxes">
      <label><input type="checkbox" name="day" value="1" checked> Mon</label>
      <label><input type="checkbox" name="day" value="2" checked> Tue</label>
      <label><input type="checkbox" name="day" value="3" checked> Wed</label>
      <label><input type="checkbox" name="day" value="4" checked> Thu</label>
      <label><input type="checkbox" name="day" value="5" checked> Fri</label>
      <label><input type="checkbox" name="day" value="6"> Sat</label>
      <label><input type="checkbox" name="day" value="0"> Sun</label>
    </div>
    <div class="btn-row" style="margin-top:10px;">
      <button type="button" class="btn-select" onclick="selectWeekdays()">Weekdays Only</button>
      <button type="button" class="btn-select" onclick="selectWeekends()">Weekends Only</button>
      <button type="button" class="btn-select" onclick="selectAll('day')">All Days</button>
    </div>
  </div>
  
  <div class="section">
    <div class="section-title">3. Date Range (Optional)</div>
    <label>Start: <input type="date" id="startDate"></label>
    <label>End: <input type="date" id="endDate"></label>
    <p style="font-size: 12px; color: #666;">Leave blank to use entire calendar range.</p>
  </div>
  
  <div class="section">
    <div class="section-title">4. Select Employees</div>
    <div class="btn-row">
      <button type="button" class="btn-select" onclick="selectAll('emp')">Select All</button>
      <button type="button" class="btn-select" onclick="deselectAll('emp')">Deselect All</button>
    </div>
    <div class="employee-list">${employeeCheckboxes}</div>
  </div>
  
  <div class="btn-row">
    <button class="btn-primary" onclick="apply()">Apply Shifts</button>
    <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
  </div>
  
  <script>
    function selectWeekdays() {
      document.querySelectorAll('input[name="day"]').forEach(cb => {
        cb.checked = ['1','2','3','4','5'].includes(cb.value);
      });
    }
    function selectWeekends() {
      document.querySelectorAll('input[name="day"]').forEach(cb => {
        cb.checked = ['0','6'].includes(cb.value);
      });
    }
    function selectAll(name) {
      document.querySelectorAll('input[name="' + name + '"]').forEach(cb => cb.checked = true);
    }
    function deselectAll(name) {
      document.querySelectorAll('input[name="' + name + '"]').forEach(cb => cb.checked = false);
    }
    function apply() {
      const shift = document.getElementById('shift').value;
      const days = Array.from(document.querySelectorAll('input[name="day"]:checked')).map(cb => parseInt(cb.value));
      const empIndices = Array.from(document.querySelectorAll('input[name="emp"]:checked')).map(cb => parseInt(cb.value));
      const startDate = document.getElementById('startDate').value;
      const endDate = document.getElementById('endDate').value;
      
      if (days.length === 0) { alert('Please select at least one day.'); return; }
      if (empIndices.length === 0) { alert('Please select at least one employee.'); return; }
      
      google.script.run
        .withSuccessHandler(() => {
          alert('Shifts assigned successfully!');
          google.script.host.close();
        })
        .withFailureHandler((err) => alert('Error: ' + err.message))
        .applyBulkShifts(shift, days, empIndices, startDate, endDate);
    }
  </script>
</body>
</html>`;
}

/**
 * Server-side function to apply bulk shifts based on user selection.
 */
function applyBulkShifts(shift, days, empIndices, startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  const headers = cal.getRange(1, 1, 1, cal.getLastColumn()).getValues()[0];
  const tz = ss.getSpreadsheetTimeZone();
  
  // Determine date range
  let startCol = 2;
  let endCol = headers.length;
  
  if (startDate) {
    const sd = new Date(startDate);
    for (let c = 2; c <= headers.length; c++) {
      const headerDate = new Date(headers[c - 1]);
      if (headerDate >= sd) { startCol = c; break; }
    }
  }
  if (endDate) {
    const ed = new Date(endDate);
    for (let c = headers.length; c >= 2; c--) {
      const headerDate = new Date(headers[c - 1]);
      if (headerDate <= ed) { endCol = c; break; }
    }
  }
  
  // Apply shifts
  let count = 0;
  for (const empIdx of empIndices) {
    const row = empIdx + 2; // +2 because empIdx is 0-based and row 1 is header
    
    for (let col = startCol; col <= endCol; col++) {
      const headerDate = new Date(headers[col - 1]);
      const dayOfWeek = headerDate.getDay();
      
      if (days.includes(dayOfWeek)) {
        const cell = cal.getRange(row, col);
        const currentVal = cell.getValue();
        // Append shift if cell already has a value, otherwise set it
        if (currentVal && currentVal !== "") {
          const existingShifts = currentVal.toString().split(',').map(s => s.trim());
          if (!existingShifts.includes(shift)) {
            cell.setValue(existingShifts.concat(shift).join(', '));
            count++;
          }
        } else {
          cell.setValue(shift);
          count++;
        }
      }
    }
  }
  
  return count;
}

/**
 * Assign shifts to employees by department.
 */
function assignByDepartment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  const empSheet = ss.getSheetByName(EMPLOYEE_SHEET);
  
  if (!cal || cal.getLastRow() < 2) {
    ui.alert('Error', 'Please create a calendar first.', ui.ButtonSet.OK);
    return;
  }
  
  // Get departments from employee data
  const empData = empSheet.getDataRange().getValues().slice(1);
  const departments = [...new Set(empData.map(r => r[2]).filter(d => d))];
  
  if (departments.length === 0) {
    ui.alert('Error', 'No departments found. Please refresh employee data.', ui.ButtonSet.OK);
    return;
  }
  
  // Get shifts
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const shiftData = settings.getRange("A35:A100").getValues().flat().filter(s => s !== "");
  
  if (shiftData.length === 0) {
    ui.alert('Error', 'No shifts defined.', ui.ButtonSet.OK);
    return;
  }
  
  const htmlTemplate = HtmlService.createHtmlOutput(buildDepartmentDialog(shiftData, departments))
    .setWidth(450)
    .setHeight(400);
  ui.showModalDialog(htmlTemplate, '🏢 Assign Shifts by Department');
}

/**
 * Builds the HTML content for the department assignment dialog
 */
function buildDepartmentDialog(shifts, departments) {
  const shiftOptions = shifts.map(s => `<option value="${s}">${s}</option>`).join('');
  const deptOptions = departments.map(d => `<option value="${d}">${d}</option>`).join('');
  
  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; padding: 15px; }
    .section { margin-bottom: 20px; }
    .section-title { font-weight: bold; color: #1c4587; margin-bottom: 8px; }
    select { width: 100%; padding: 10px; margin: 5px 0; border: 1px solid #ccc; border-radius: 4px; }
    .day-checkboxes { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 10px; }
    .day-checkboxes label { 
      padding: 8px 12px; background: #e8f0fe; border-radius: 4px; cursor: pointer;
      border: 2px solid transparent;
    }
    .day-checkboxes label:has(input:checked) { background: #1c4587; color: white; }
    .day-checkboxes input { display: none; }
    .btn-row { display: flex; gap: 10px; margin-top: 20px; }
    button { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; }
    .btn-primary { background: #1c4587; color: white; }
    .btn-secondary { background: #f1f3f4; color: #333; }
  </style>
</head>
<body>
  <div class="section">
    <div class="section-title">1. Select Department</div>
    <select id="dept">${deptOptions}</select>
  </div>
  
  <div class="section">
    <div class="section-title">2. Select Shift</div>
    <select id="shift">${shiftOptions}</select>
  </div>
  
  <div class="section">
    <div class="section-title">3. Select Days of Week</div>
    <div class="day-checkboxes">
      <label><input type="checkbox" name="day" value="1" checked> Mon</label>
      <label><input type="checkbox" name="day" value="2" checked> Tue</label>
      <label><input type="checkbox" name="day" value="3" checked> Wed</label>
      <label><input type="checkbox" name="day" value="4" checked> Thu</label>
      <label><input type="checkbox" name="day" value="5" checked> Fri</label>
      <label><input type="checkbox" name="day" value="6"> Sat</label>
      <label><input type="checkbox" name="day" value="0"> Sun</label>
    </div>
  </div>
  
  <div class="btn-row">
    <button class="btn-primary" onclick="apply()">Apply to Department</button>
    <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
  </div>
  
  <script>
    function apply() {
      const dept = document.getElementById('dept').value;
      const shift = document.getElementById('shift').value;
      const days = Array.from(document.querySelectorAll('input[name="day"]:checked')).map(cb => parseInt(cb.value));
      
      if (days.length === 0) { alert('Please select at least one day.'); return; }
      
      google.script.run
        .withSuccessHandler((count) => {
          alert('Assigned ' + count + ' shifts to department: ' + dept);
          google.script.host.close();
        })
        .withFailureHandler((err) => alert('Error: ' + err.message))
        .applyDepartmentShifts(dept, shift, days);
    }
  </script>
</body>
</html>`;
}

/**
 * Server-side function to apply shifts to a department.
 */
function applyDepartmentShifts(department, shift, days) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  const empSheet = ss.getSheetByName(EMPLOYEE_SHEET);
  
  // Get employees in the department
  const empData = empSheet.getDataRange().getValues().slice(1);
  const deptEmployees = empData.filter(r => r[2] === department).map(r => r[1]);
  
  // Get calendar employee list
  const calEmployees = cal.getRange(2, 1, cal.getLastRow() - 1, 1).getValues().flat();
  const empIndices = [];
  
  calEmployees.forEach((emp, idx) => {
    if (deptEmployees.includes(emp)) {
      empIndices.push(idx);
    }
  });
  
  if (empIndices.length === 0) {
    throw new Error('No employees from this department found in the calendar.');
  }
  
  return applyBulkShifts(shift, days, empIndices, '', '');
}

/**
 * Copy a week's pattern and repeat it across the calendar.
 */
function copyWeekPattern() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  
  if (!cal || cal.getLastRow() < 2) {
    ui.alert('Error', 'Please create a calendar first.', ui.ButtonSet.OK);
    return;
  }
  
  // Get the source week start date
  const response = ui.prompt(
    'Copy Week Pattern',
    'Enter the Monday date of the week to copy (YYYY-MM-DD):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const sourceMonday = new Date(response.getResponseText());
  if (isNaN(sourceMonday.getTime()) || sourceMonday.getDay() !== 1) {
    ui.alert('Error', 'Please enter a valid Monday date (YYYY-MM-DD).', ui.ButtonSet.OK);
    return;
  }
  
  const headers = cal.getRange(1, 1, 1, cal.getLastColumn()).getValues()[0];
  const employees = cal.getRange(2, 1, cal.getLastRow() - 1, 1).getValues().flat();
  const tz = ss.getSpreadsheetTimeZone();
  
  // Find source week columns (Monday to Sunday)
  let sourceStartCol = -1;
  for (let c = 1; c < headers.length; c++) {
    const headerDate = new Date(headers[c]);
    const headerStr = Utilities.formatDate(headerDate, tz, 'yyyy-MM-dd');
    const sourceStr = Utilities.formatDate(sourceMonday, tz, 'yyyy-MM-dd');
    if (headerStr === sourceStr) {
      sourceStartCol = c + 1; // +1 because column index is 1-based
      break;
    }
  }
  
  if (sourceStartCol === -1) {
    ui.alert('Error', 'Source week not found in the calendar.', ui.ButtonSet.OK);
    return;
  }
  
  // Read source week data (7 days or less if near end of calendar)
  const lastCol = cal.getLastColumn();
  const sourceDays = Math.min(7, lastCol - sourceStartCol + 1);
  const sourceData = cal.getRange(2, sourceStartCol, employees.length, sourceDays).getValues();
  
  // Apply to all other weeks
  let weeksUpdated = 0;
  
  for (let col = 2; col <= lastCol; col++) {
    const headerDate = new Date(headers[col - 1]);
    if (headerDate.getDay() === 1 && col !== sourceStartCol) {
      // This is a Monday - copy the week pattern
      const targetDays = Math.min(sourceDays, lastCol - col + 1);
      const targetRange = cal.getRange(2, col, employees.length, targetDays);
      const targetData = [];
      
      for (let row = 0; row < employees.length; row++) {
        const weekRow = [];
        for (let day = 0; day < targetDays; day++) {
          weekRow.push(sourceData[row][day] || '');
        }
        targetData.push(weekRow);
      }
      
      targetRange.setValues(targetData);
      weeksUpdated++;
    }
  }
  
  ui.alert('Success', `Copied week pattern to ${weeksUpdated} other weeks.`, ui.ButtonSet.OK);
}

/**
 * Clear all shift assignments from the calendar.
 */
function clearCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  
  if (!cal || cal.getLastRow() < 2) {
    ui.alert('Error', 'No calendar found.', ui.ButtonSet.OK);
    return;
  }
  
  const confirm = ui.alert(
    'Clear Calendar',
    'This will remove all shift assignments from the calendar. Are you sure?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  const lastRow = cal.getLastRow();
  const lastCol = cal.getLastColumn();
  
  if (lastRow > 1 && lastCol > 1) {
    cal.getRange(2, 2, lastRow - 1, lastCol - 1).clearContent();
    ui.alert('Success', 'Calendar cleared.', ui.ButtonSet.OK);
  }
}

// --- 4. PROCESS TO SQL ---
function processCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cal = ss.getSheetByName(CALENDAR_SHEET);
  const sqlSheet = ss.getSheetByName(SQL_SHEET);
  
  const empMap = new Map(); 
  const empData = ss.getSheetByName(EMPLOYEE_SHEET).getDataRange().getValues();
  empData.slice(1).forEach(r => empMap.set(r[1], {id:r[0], co:r[3]}));

  // Lookup IDs (A-F)
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const refData = settings.getRange("A13:F30").getValues(); 
  const idLookup = new Map();
  refData.forEach(r => {
    if(r[0]) idLookup.set(r[0], r[1]); 
    if(r[2]) idLookup.set(r[2], r[3]); 
    if(r[4]) idLookup.set(r[4], r[5]); 
  });

  const shiftMap = new Map();
  const shiftRows = settings.getRange("A35:H100").getValues();
  shiftRows.forEach(r => {
    if(r[0]) {
      shiftMap.set(r[0], {
        pStart: r[1], pEnd: r[2], bStart: r[3], bEnd: r[4],
        pID: idLookup.get(r[5]) || "", 
        bID: idLookup.get(r[6]) || "",
        prID: idLookup.get(r[7]) || ""
      });
    }
  });

  const data = cal.getDataRange().getValues();
  const dateHeaders = data[0];
  const output = [];

  for(let r=1; r<data.length; r++) {
    const empName = data[r][0];
    const emp = empMap.get(empName);
    if(!emp) continue;

    for(let c=1; c<data[r].length; c++) {
      const cell = data[r][c];
      if(cell && cell !== "") {
        const shifts = cell.toString().split(',').map(s => s.trim());
        shifts.forEach(shiftName => {
          if(shiftMap.has(shiftName)) {
            const shift = shiftMap.get(shiftName);
            const dateStr = dateHeaders[c]; 
            output.push([
              emp.id, emp.co, dateStr,
              shift.pID, shift.pStart, shift.pEnd,
              shift.bID, shift.bStart, shift.bEnd,
              shift.prID, shift.pStart, shift.pEnd,
              'Pending'
            ]);
          }
        });
      }
    }
  }

  sqlSheet.clear();
  const headers = ['UserID', 'CompanyID', 'Schedule_Date', 'PresenceTypeID', 'Presence_Start_Time', 'Presence_End_Time', 'BreakTypeID', 'Break_Start_Time', 'Break_End_Time', 'ProjectID', 'Project_Start_Time', 'Project_End_Time', 'Status'];
  
  if(output.length > 0) {
    const range = sqlSheet.getRange(1,1,output.length + 1, output[0].length);
    range.setValues([headers, ...output]);
    
    sqlSheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#cc0000').setFontColor('white');
    range.setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
    range.applyRowBanding(SpreadsheetApp.BandingTheme.PINK);
    
    SpreadsheetApp.getUi().alert(`Processed ${output.length} shifts.`);
  } else {
    SpreadsheetApp.getUi().alert("No shifts found in calendar.");
  }
}

// --- 5. EMAIL SQL ---
function emailSql() {
  const ui = SpreadsheetApp.getUi();
  
  const resp = ui.prompt("Email SQL", "Enter recipient email address:", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  let email = resp.getResponseText().trim();
  if (!email) { ui.alert("No email provided."); return; }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SQL_SHEET);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) { ui.alert("No data to send."); return; }

  const sqls = [];
  const updates = [];
  const fmt = (dStr, tVal) => {
    const d = new Date(dStr);
    let tStr = tVal;
    if (tVal instanceof Date) tStr = Utilities.formatDate(tVal, Session.getScriptTimeZone(), 'HH:mm:ss');
    const datePart = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return `${datePart} ${tStr}+00`;
  };

  for(let i=1; i<data.length; i++) {
    const row = data[i]; 
    if(row[12] === 'Pending') {
      let q = `
with new_time_registration as (
 insert into time_registration_presence (
   time_registration_presence_type_id, start_date, end_date, user_id, company_id
 ) values (
   '${row[3]}', '${fmt(row[2],row[4])}', '${fmt(row[2],row[5])}', '${row[0]}', '${row[1]}'
 ) returning id
)
insert into time_registration_entry (
 time_registration_entry_type, start_date, end_date, user_id, company_id, 
 time_registration_break_type_id, time_registration_project_id, time_registration_presence_id
) values (
 'BREAK', '${fmt(row[2],row[7])}', '${fmt(row[2],row[8])}', '${row[0]}', '${row[1]}', '${row[6]}', null, (select id from new_time_registration)
), (
 'PROJECT', '${fmt(row[2],row[10])}', '${fmt(row[2],row[11])}', '${row[0]}', '${row[1]}', null, '${row[9]}', (select id from new_time_registration)
);`;
      sqls.push(q);
      updates.push(sheet.getRange(i+1, 13));
    }
  }

  if(sqls.length > 0) {
    MailApp.sendEmail(email, "HR-ON SQL Import", sqls.join("\n\n-- NEXT ENTRY --\n\n"));
    updates.forEach(c => c.setValue('Sent'));
    ui.alert(`Sent ${sqls.length} queries to ${email}.`);
  } else {
    ui.alert("No Pending rows found.");
  }
}

// ==========================================
//      HELPER: FACTORY RESET (STYLED & CLEAN)
// ==========================================
function initializeSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Reset Sheet?', 'This will delete content and apply Professional Styles. Continue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  // 1. Settings (The Master Sheet)
  let s = ss.getSheetByName(SETTINGS_SHEET) || ss.insertSheet(SETTINGS_SHEET);
  s.clear();
  s.setHiddenGridlines(true); 
  s.setTabColor("#4285f4");
  
  // Block A: System Config (Blue Table)
  s.getRange("A1:B1").merge().setValue("SYSTEM CONFIGURATION").setFontWeight("bold").setBackground("#1c4587").setFontColor("white").setHorizontalAlignment("center");
  const defaults = [
    ["CLIENT_ID", ""], ["CLIENT_SECRET", ""], 
    ["TOKEN_URL", "https://auth.hr-on.com/oauth2/token"],
    ["USERS_API_URL", "https://api.hr-on.com/v1/staff/employees?size=1000"],
    ["COMPANY_API_URL", "https://api.hr-on.com/v1/staff/company"],
    ["DEPARTMENTS_API_URL", "https://api.hr-on.com/v1/staff/departments"], 
    ["Default Company ID", ""]
  ];
  const configRange = s.getRange(2, 1, defaults.length, 2);
  configRange.setValues(defaults);
  configRange.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE);
  configRange.setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  
  // Block B: Reference Data (Yellow Table - NO GAPS)
  s.getRange("A11:F11").merge().setValue("REFERENCE DATA (PASTE IDs HERE)").setFontWeight("bold").setBackground("#bf9000").setFontColor("white").setHorizontalAlignment("center");
  const refHeaders = [["PRESENCE NAME","ID","BREAK NAME","ID","PROJECT NAME","ID"]];
  s.getRange("A12:F12").setValues(refHeaders).setFontWeight("bold").setBackground("#f1c232")
   .setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Style the Reference Data Area
  const refRange = s.getRange("A13:F30");
  refRange.applyRowBanding(SpreadsheetApp.BandingTheme.YELLOW);
  refRange.setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
  // Thick Dividers between sections (After B and D)
  s.getRange("B12:B30").setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  s.getRange("D12:D30").setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Block C: Shift Definitions (Green Table)
  s.getRange("A33:H33").merge().setValue("SHIFT DEFINITIONS (CONFIGURE SCHEDULES)").setFontWeight("bold").setBackground("#38761d").setFontColor("white").setHorizontalAlignment("center");
  s.getRange("A34:H34").setValues([["Shift Name","Start","End","Brk Start","Brk End","Presence","Break","Project"]])
   .setFontWeight("bold").setBackground("#6aa84f").setFontColor("white")
   .setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  const shiftRange = s.getRange("A35:H100");
  shiftRange.applyRowBanding(SpreadsheetApp.BandingTheme.GREEN);
  shiftRange.setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
  
  s.getRange("A35:H35").setValues([["Morning A","08:00","16:00","12:00","12:30","(Run 'Update Dropdowns')","(Select)","(Select)"]]);

  // Sizing
  s.setColumnWidth(1, 160); s.setColumnWidth(2, 200); 
  s.setColumnWidth(3, 160); s.setColumnWidth(4, 200); 
  s.setColumnWidth(5, 160); s.setColumnWidth(6, 200);
  s.setColumnWidth(7, 160); s.setColumnWidth(8, 200);
  
  // Hide Unused Columns (I to Z)
  const maxCols = s.getMaxColumns();
  if (maxCols > 8) s.hideColumns(9, maxCols - 8);

  // 2. Clean Up other sheets
  let e = ss.getSheetByName(EMPLOYEE_SHEET) || ss.insertSheet(EMPLOYEE_SHEET); e.clear(); e.setTabColor("#0097a7");
  let q = ss.getSheetByName(SQL_SHEET) || ss.insertSheet(SQL_SHEET); q.clear(); q.setTabColor("#ea4335");
  
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  ui.alert("Reset Complete! Tables are solid and high-contrast.");
}

// --- UTILS ---
function getConfiguration() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
  const data = sheet.getRange("A2:B9").getValues();
  const config = {};
  data.forEach(r => { 
    if(r[0]) config[String(r[0]).trim().toUpperCase().replace(" ", "_")] = r[1]; 
  });
  config['DEFAULT_COMPANY_ID'] = config['DEFAULT_COMPANY_ID'] || "";
  return config;
}

function getAccessToken(config) {
  if (!config.CLIENT_ID || config.CLIENT_ID === "") throw new Error("🛑 CLIENT_ID is missing in Settings.");
  if (!config.CLIENT_SECRET || config.CLIENT_SECRET === "") throw new Error("🛑 CLIENT_SECRET is missing in Settings.");

  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('TOKEN');
  if (saved) { const t = JSON.parse(saved); if (t.exp > Date.now()) return t.val; }
  
  const auth = Utilities.base64Encode(`${config.CLIENT_ID}:${config.CLIENT_SECRET}`);
  const resp = UrlFetchApp.fetch(config.TOKEN_URL, {
    method: 'post', payload: { grant_type: 'client_credentials' },
    headers: { 'Authorization': 'Basic ' + auth }
  });
  const json = JSON.parse(resp.getContentText());
  props.setProperty('TOKEN', JSON.stringify({ val: json.access_token, exp: Date.now() + (json.expires_in-300)*1000 }));
  return json.access_token;
}
