/**
 * Service_Manage.gs
 */
function getManageData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manage");
  const data = sheet.getDataRange().getDisplayValues();
  const props = PropertiesService.getUserProperties();
  
  // 1. Get the PID currently saved in memory
  // Use the key 'ACTIVE_PID' consistently
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
      // 2. This check now uses 'activePid' which IS defined above
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


/*============================================
* Edit a Project
*============================================*/

function updateManageRow(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manage");
  const row = parseInt(payload.rowIndex);

  // Only update the Status column
  sheet.getRange(row, 6).setValue(payload.status);

  SpreadsheetApp.flush(); 
  return getManageData();
}


/*============================================
* Create a Project
*============================================*/

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

	// Create PiD
  const now = new Date();
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMddyyHHmmss");
  const randomTail = Math.floor(1000 + Math.random() * 9000);
  const newPid = `${clientCode}-${timeStr}${randomTail}`;

  const manageSheet = ss.getSheetByName("Manage");
  const manageRow = [newPid, formData.type, formData.client, "New", now, "Active", formData.policy, formData.claim];
  manageSheet.appendRow(manageRow);
  
  sheetsToUpdate.forEach(sheetName => {
    const targetSheet = ss.getSheetByName(sheetName);
    if (targetSheet) {
			targetSheet.appendRow([newPid]);
    }
  });

  webSelectProject(manageSheet.getLastRow());
  return getManageData();
}


/*============================================
* Delete a Project
*============================================*/

function deleteProject(pid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const manageSheet = ss.getSheetByName("Manage");
  const data = manageSheet.getDataRange().getValues();
  
  // Find the row index based on PID
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === pid.toString()) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex !== -1) {
    manageSheet.deleteRow(rowIndex);
    
    // Cleanup: If the deleted project was the active one, clear memory
    const props = PropertiesService.getUserProperties();
    if (props.getProperty('ACTIVE_PID') === pid) {
      props.deleteProperty('ACTIVE_PID');
      props.deleteProperty('ACTIVE_PROJECT_ROW');
    }
    
    // Optional: Delete from linked sheets if you want total cleanup
    const settingsSheet = ss.getSheetByName("Settings");
    const linkedSheets = settingsSheet.getDataRange().getValues()
      .filter(row => row[0] === "LinkedSheet")
      .map(row => row[1]);

    linkedSheets.forEach(sName => {
      const s = ss.getSheetByName(sName);
      if (s) {
        const sData = s.getDataRange().getValues();
        for (let j = sData.length - 1; j >= 0; j--) {
          if (sData[j][0].toString() === pid.toString()) s.deleteRow(j + 1);
        }
      }
    });
  }
  return getManageData();
}