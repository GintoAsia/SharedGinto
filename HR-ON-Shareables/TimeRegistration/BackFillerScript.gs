// ==========================================
//      HR-ON BACKFILL TOOL (AUTO COMPANY ID)
// ==========================================

const SETTINGS_SHEET = 'Settings_Backfill';
const EMP_SHEET = 'Employees_Backfill';
const RECORD_SHEET = 'Record';
const SQL_SHEET = 'SQL_Export';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('HR-ON Backfill')
    .addItem('⚙️ 1. Initialize Settings', 'initSettings')
    .addItem('🔄 2. Refresh Employee List & Company ID', 'refreshEmployeeList')
    .addSeparator()
    .addItem('🚀 3. Run Backfill Process', 'runBackfill')
    .addToUi();
}

// ==========================================
//      STEP 1: INITIALIZATION
// ==========================================
function initSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  if (ss.getSheetByName(SETTINGS_SHEET)) {
    if (ui.alert("Reset Settings?", "This will overwrite the Settings tab. Continue?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    ss.deleteSheet(ss.getSheetByName(SETTINGS_SHEET));
  }

  const s = ss.insertSheet(SETTINGS_SHEET);
  s.setTabColor("#4285f4"); 
  s.setHiddenGridlines(true);

  // --- BLOCK A: SYSTEM CONFIG (BLUE) ---
  s.getRange("A1:B1").merge().setValue("SYSTEM CONFIGURATION").setFontWeight("bold").setBackground("#1c4587").setFontColor("white").setHorizontalAlignment("center");
  const defaults = [
    ["CLIENT_ID", ""], ["CLIENT_SECRET", ""], 
    ["TOKEN_URL", "https://auth.hr-on.com/oauth2/token"],
    ["USERS_API_URL", "https://api.hr-on.com/v1/staff/employees?size=1000"],
    ["COMPANY_API_URL", "https://api.hr-on.com/v1/staff/company"],
    ["Default Company ID", "(Auto-filled after Step 2)"] // Placeholder
  ];
  const configRange = s.getRange(2, 1, defaults.length, 2);
  configRange.setValues(defaults);
  configRange.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE);
  configRange.setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  
  // --- BLOCK B: REFERENCE DATA (YELLOW) ---
  s.getRange("A11:F11").merge().setValue("REFERENCE DATA (PASTE FROM DASHBOARD TOOL)").setFontWeight("bold").setBackground("#bf9000").setFontColor("white").setHorizontalAlignment("center");
  s.getRange("A12:F12").setValues([["PRESENCE NAME","ID","BREAK NAME","ID","PROJECT NAME","ID"]])
   .setFontWeight("bold").setBackground("#f1c232");
  
  const refRange = s.getRange("A13:F30");
  refRange.applyRowBanding(SpreadsheetApp.BandingTheme.YELLOW);
  refRange.setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);

  // --- BLOCK C: BACKFILL MAPPING (PURPLE) ---
  s.getRange("A33:F33").merge().setValue("BACKFILL MAPPING & SELECTION").setFontWeight("bold").setBackground("#674ea7").setFontColor("white").setHorizontalAlignment("center");
  
  s.getRange("A34:F34").setValues([["'Type of Day' (Record Tab)", "Presence ID", "Break ID", "Project ID", "", "Selected Employee:"]])
   .setFontWeight("bold").setBackground("#8e7cc3").setFontColor("white");

  const sampleMap = [
    ["Workday", "(Paste ID)", "(Paste ID)", "", "", "(Run 'Refresh Employee List')"],
    ["Workday: Full Day", "(Paste ID)", "(Paste ID)", "", "", ""],
    ["Workday: Flex", "(Paste ID)", "(Paste ID)", "", "", ""],
    ["Leave: Sick", "(Paste ID)", "", "", "", ""],
    ["Leave: Personal", "(Paste ID)", "", "", "", ""]
  ];
  
  const mapRange = s.getRange(35, 1, sampleMap.length, 6);
  mapRange.setValues(sampleMap);
  s.getRange("A35:D50").applyRowBanding(SpreadsheetApp.BandingTheme.PINK); 
  
  // Dropdowns
  const pRule = SpreadsheetApp.newDataValidation().requireValueInRange(s.getRange("B13:B30")).setAllowInvalid(true).build();
  const bRule = SpreadsheetApp.newDataValidation().requireValueInRange(s.getRange("D13:D30")).setAllowInvalid(true).build();
  const prRule = SpreadsheetApp.newDataValidation().requireValueInRange(s.getRange("F13:F30")).setAllowInvalid(true).build(); 

  s.getRange("B35:B50").setDataValidation(pRule);
  s.getRange("C35:C50").setDataValidation(bRule);
  s.getRange("D35:D50").setDataValidation(prRule);

  // Formatting
  s.setColumnWidth(1, 200); s.setColumnWidth(2, 250); s.setColumnWidth(3, 250); s.setColumnWidth(4, 250); s.setColumnWidth(6, 250); 

  ui.alert("Settings Initialized!\n\n1. Copy Blue/Yellow tables from Dashboard.\n2. Run 'Refresh Employee List'.\n3. Select Employee in Cell F35.");
}

// ==========================================
//      STEP 2: REFRESH EMPLOYEES & COMPANY ID
// ==========================================
function refreshEmployeeList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  
  if (!settings) { ui.alert("Run Initialize Settings first."); return; }

  const config = getConfiguration(settings);
  if (!config.CLIENT_ID) { ui.alert("Missing CLIENT_ID in Settings."); return; }
  
  try {
    const token = getAccessToken(config);
    
    // --- PART A: FETCH COMPANY ID ---
    let companyId = "";
    if (config.COMPANY_API_URL) {
      try {
        const cResp = UrlFetchApp.fetch(config.COMPANY_API_URL, { 
          headers: { 'Authorization': `Bearer ${token}` },
          muteHttpExceptions: true
        });
        const cJson = JSON.parse(cResp.getContentText());
        // Handle array or object response
        const cData = Array.isArray(cJson) ? cJson[0] : (cJson.data ? (Array.isArray(cJson.data) ? cJson.data[0] : cJson.data) : cJson);
        
        if (cData && cData.id) {
          companyId = cData.id;
          // UPDATE SETTINGS SHEET WITH COMPANY ID (Cell B7)
          settings.getRange("B7").setValue(companyId);
        }
      } catch (err) {
        console.warn("Could not fetch Company ID: " + err.message);
      }
    }

    // --- PART B: FETCH EMPLOYEES ---
    const resp = UrlFetchApp.fetch(config.USERS_API_URL, { 
      headers: { 'Authorization': `Bearer ${token}` },
      muteHttpExceptions: true 
    });
    
    if (resp.getResponseCode() !== 200) throw new Error(`API Error ${resp.getResponseCode()}`);

    const json = JSON.parse(resp.getContentText());
    let rawList = [];
    if (Array.isArray(json)) rawList = json;
    else if (json.data && Array.isArray(json.data)) rawList = json.data;
    else if (json.items && Array.isArray(json.items)) rawList = json.items;
    else if (json.employees && Array.isArray(json.employees)) rawList = json.employees;
    
    if (rawList.length === 0) {
      ui.alert(`Connected, but found 0 employees.`);
      return;
    }

    const output = rawList.map(u => {
      let first = u.firstName || (u.systemFields && u.systemFields.firstName) || "";
      let last = u.lastName || (u.systemFields && u.systemFields.lastName) || "";
      let name = `${first} ${last}`.trim();
      if (!name) name = u.name || u.email || "Unknown";
      
      // Use fetched company ID if user doesn't have one specific
      let uCo = u.companyId || companyId || config['DEFAULT_COMPANY_ID'];
      
      return [name, u.id, uCo];
    }).sort((a,b) => a[0].localeCompare(b[0]));

    output.unshift(["Employee Name", "Employee ID", "Company ID"]); 

    let eSheet = ss.getSheetByName(EMP_SHEET);
    if (!eSheet) { eSheet = ss.insertSheet(EMP_SHEET); eSheet.setTabColor("#0097a7"); }
    eSheet.clear();
    
    eSheet.getRange(1, 1, output.length, 3).setValues(output);
    eSheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#134f5c").setFontColor("white");
    eSheet.autoResizeColumns(1, 3);
    eSheet.setHiddenGridlines(true);

    if (output.length > 1) {
      const empNameRange = eSheet.getRange(2, 1, output.length - 1, 1);
      const rule = SpreadsheetApp.newDataValidation().requireValueInRange(empNameRange).setAllowInvalid(false).build();
      settings.getRange("F35").setDataValidation(rule).setValue("(Select Employee)");
      
      let msg = `Success! Fetched ${output.length - 1} employees.`;
      if (companyId) msg += `\n\n✅ Company ID found and saved to Settings: ${companyId}`;
      ui.alert(msg);
    }

  } catch (e) {
    ui.alert("Error: " + e.message);
  }
}

// ==========================================
//      STEP 3: RUN BACKFILL
// ==========================================
function runBackfill() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const record = ss.getSheetByName(RECORD_SHEET);
  const empSheet = ss.getSheetByName(EMP_SHEET);

  if (!record) { ui.alert("Error: 'Record' tab not found."); return; }
  
  // 1. Get Selected Employee
  const empName = settings.getRange("F35").getValue();
  if (!empName || empName.includes("Select") || empName === "") {
    ui.alert("Please select an employee in Settings (Cell F35).");
    return;
  }

  // 2. Lookup ID & Company ID
  const empData = empSheet.getDataRange().getValues();
  let empId = "";
  let companyId = "";
  for (let i = 1; i < empData.length; i++) {
    if (empData[i][0] === empName) {
      empId = empData[i][1];
      companyId = empData[i][2]; // Preference: Employee specific Company ID
      break;
    }
  }
  
  // Fallback to Settings if Employee sheet didn't have Company ID
  if (!companyId) {
    companyId = settings.getRange("B7").getValue(); 
  }

  if (!empId) { ui.alert("Employee ID not found."); return; }
  if (!companyId) { ui.alert("Company ID missing. Please Run Step 2 again to fetch it."); return; }

  // 3. Build Mapping
  const mapRange = settings.getRange("A35:D50").getValues();
  const typeMap = new Map();
  mapRange.forEach(r => {
    if(r[0]) typeMap.set(r[0].toString().trim(), { pID: r[1], bID: r[2], prID: r[3] });
  });

  // 4. Process Record Data
  const lastRow = record.getLastRow();
  const data = record.getRange(2, 1, lastRow - 1, 5).getValues(); 
  const output = [];
  const tz = ss.getSpreadsheetTimeZone();
  
  const fmt = (dateObj, timeVal) => {
    if (!timeVal || !dateObj) return null;
    const d = new Date(dateObj);
    const t = new Date(timeVal);
    if (timeVal instanceof Date) { d.setHours(t.getHours(), t.getMinutes(), t.getSeconds()); }
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm:ss') + "+00";
  };

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const dateObj = row[0];
    const typeStr = row[1] ? row[1].toString().trim() : "";
    const startTime = row[2];
    const endTime = row[3];
    const breakDur = row[4];

    if (!dateObj || !typeStr) continue;

    const mapping = typeMap.get(typeStr);
    
    if (mapping && mapping.pID && startTime && endTime) {
      
      const startSql = fmt(dateObj, startTime);
      const endSql = fmt(dateObj, endTime);
      const projSql = mapping.prID ? `'${mapping.prID}'` : 'NULL';

      let sql = `with new_time_registration as (
 insert into time_registration_presence (
   time_registration_presence_type_id, start_date, end_date, user_id, company_id, project_id
 ) values (
   '${mapping.pID}', '${startSql}', '${endSql}', '${empId}', '${companyId}', ${projSql}
 ) returning id
)`;

      let entries = [];
      if (mapping.bID && breakDur) {
        const bStartObj = new Date(dateObj);
        const tStart = new Date(startTime);
        const tDur = new Date(breakDur);
        
        bStartObj.setHours(tStart.getHours() + 4, tStart.getMinutes(), 0);
        
        const bEndObj = new Date(bStartObj);
        bEndObj.setHours(bEndObj.getHours() + tDur.getHours(), bEndObj.getMinutes() + tDur.getMinutes(), 0);

        const bsStr = Utilities.formatDate(bStartObj, tz, 'yyyy-MM-dd HH:mm:ss') + "+00";
        const beStr = Utilities.formatDate(bEndObj, tz, 'yyyy-MM-dd HH:mm:ss') + "+00";

        entries.push(`('BREAK', '${bsStr}', '${beStr}', '${empId}', '${companyId}', '${mapping.bID}', null, (select id from new_time_registration))`);
      }

      if (entries.length > 0) {
        sql += `
insert into time_registration_entry (
 time_registration_entry_type, start_date, end_date, user_id, company_id, 
 time_registration_break_type_id, time_registration_project_id, time_registration_presence_id
) values 
${entries.join(',\n')};`;
      } else {
        sql += `;`;
      }

      output.push([Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd'), typeStr, sql]);
    }
  }

  // 5. Output
  let sqlSheet = ss.getSheetByName(SQL_SHEET);
  if (!sqlSheet) sqlSheet = ss.insertSheet(SQL_SHEET);
  sqlSheet.clear();
  const headers = ['Date', 'Type', 'Generated SQL'];
  sqlSheet.getRange(1, 1, 1, 3).setValues([headers]).setFontWeight("bold").setBackground("#cc0000").setFontColor("white");
  
  if (output.length > 0) {
    sqlSheet.getRange(2, 1, output.length, 3).setValues(output);
    sqlSheet.setColumnWidth(3, 500);
    ui.alert(`✅ Complete! Processed ${output.length} rows.`);
  } else {
    ui.alert("⚠️ No valid rows found. Check your Purple Mapping table.");
  }
}

// ==========================================
//      API HELPERS
// ==========================================
function getConfiguration(sheet) {
  const data = sheet.getRange("A2:B9").getValues();
  const config = {};
  data.forEach(r => { if(r[0]) config[r[0]] = r[1]; });
  config['DEFAULT_COMPANY_ID'] = config['DEFAULT_COMPANY_ID'] || "";
  return config;
}

function getAccessToken(config) {
  const auth = Utilities.base64Encode(`${config.CLIENT_ID}:${config.CLIENT_SECRET}`);
  const resp = UrlFetchApp.fetch(config.TOKEN_URL, {
    method: 'post', payload: { grant_type: 'client_credentials' },
    headers: { 'Authorization': 'Basic ' + auth }
  });
  return JSON.parse(resp.getContentText()).access_token;
}
