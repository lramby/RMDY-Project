/**
 * Service_Manage.gs
 */
function getManageData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manage");
  const data = sheet.getDataRange().getDisplayValues();
  const props = PropertiesService.getUserProperties();
  
  // Use PID for the visual check
  const activePid = (props.getProperty('ACTIVE_PID') || "").trim();

  return data.slice(1).map((row, i) => {
    const currentPid = row[0].trim();
    return {
      pid: currentPid,
      type: row[1],
      client: row[2],
      status: row[5],
      policy: row[6],
      claim: row[7],
      isChecked: (currentPid === activePid && activePid !== ""), 
      rowIndex: i + 2
    };
  });
}

function webSelectProject(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manage");
  
  // Get the PID from the sheet
  const selectedPid = sheet.getRange(rowIndex, 1).getDisplayValue().trim();
  
  const props = PropertiesService.getUserProperties();
  // Save both so the rest of your app stays in sync
  props.setProperty('ACTIVE_PID', selectedPid);
  props.setProperty('ACTIVE_PROJECT_ROW', rowIndex.toString());
  
  // Force Google to finalize the write immediately
  SpreadsheetApp.flush();
  
  return selectedPid;
}

function updateManageRow(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manage");
  const row = parseInt(payload.rowIndex);

  sheet.getRange(row, 6).setValue(payload.status);
  sheet.getRange(row, 7).setValue(payload.policy);
  sheet.getRange(row, 8).setValue(payload.claim);
  
  const pid = sheet.getRange(row, 1).getDisplayValue().trim();
  const props = PropertiesService.getUserProperties();
  props.setProperty('ACTIVE_PROJECT_ROW', row.toString());
  props.setProperty('ACTIVE_PID', pid);

  SpreadsheetApp.flush(); 
  return getManageData();
}

function getContactCompanies() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contacts");
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const companies = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) companies.push(values[i][0]);
  }
  return [...new Set(companies)].sort();
}

function createNewProject(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Settings");
  const settingsData = settingsSheet.getDataRange().getValues();
  
  const sheetsToUpdate = settingsData
    .filter(row => row[0] === "LinkedSheet")
    .map(row => row[1]);

  const contactsSheet = ss.getSheetByName("Contacts");
  const contactData = contactsSheet.getDataRange().getValues();
  let clientCode = "UNK";
  
  for (let i = 1; i < contactData.length; i++) {
    if (contactData[i][0] === formData.client) {
      clientCode = contactData[i][1]; 
      break;
    }
  }

  const now = new Date();
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMddyyHHmmss");
  const typeSuffix = (formData.type === "Water") ? "WD" : "FD";
  const randomTail = Math.floor(1000 + Math.random() * 9000);
  //const newPid = `${clientCode}-${typeSuffix}-${timeStr}${randomTail}`;
  const newPid = `${clientCode}-${timeStr}${randomTail}`;

  const manageSheet = ss.getSheetByName("Manage");
  const manageRow = [newPid, formData.type, formData.client, "New", now, "Active", formData.policy, formData.claim];
  manageSheet.appendRow(manageRow);
  
  sheetsToUpdate.forEach(sheetName => {
    const targetSheet = ss.getSheetByName(sheetName);
    if (targetSheet) {
      if (sheetName === "Details") {
        targetSheet.appendRow([newPid, "", "", "", "", "", formData.client]);
      } else {
        targetSheet.appendRow([newPid]);
      }
    }
  });

  webSelectProject(manageSheet.getLastRow());
  return getManageData();
}