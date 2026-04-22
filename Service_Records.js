function getRecordsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Records");
    const manageSheet = ss.getSheetByName("Manage");
    const rowIndex = PropertiesService.getUserProperties().getProperty('ACTIVE_PROJECT_ROW');
    
    if (!rowIndex || !manageSheet || !sheet) return [];
    const selectedPid = manageSheet.getRange(Number(rowIndex), 1).getValue();

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    return data.slice(1)
      .map((row, index) => {
        return {
          pid: row[0] ? String(row[0]) : "",
          recordType: row[1] ? String(row[1]) : "",
          readingType: row[2] ? String(row[2]) : "",
          day: row[3] || "",
          value: row[4] || "",
          taskName: row[5] ? String(row[5]) : "",
          taskId: row[6] ? String(row[6]) : "",
          roomName: row[7] ? String(row[7]) : "",
          roomId: row[8] ? String(row[8]) : "",
          note: row[9] ? String(row[9]) : "",
          sheetRow: index + 2
        };
      })
      .filter(r => String(r.pid) === String(selectedPid));
  } catch (e) {
    console.error("GS Error (getRecordsData): " + e.toString());
    return [];
  }
}

function saveRecordData(recObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Records");
  
  const rowValues = [[
    String(recObj.recordType), 
    String(recObj.readingType), 
    recObj.day,
    String(recObj.value),
    String(recObj.taskName), 
    String(recObj.taskId),
    String(recObj.roomName),
    String(recObj.roomId),
    String(recObj.note)
  ]];

  if (recObj.sheetRow && Number(recObj.sheetRow) > 0) {
    sheet.getRange(Number(recObj.sheetRow), 2, 1, 9).setValues(rowValues);
  } else {
    const manageRow = PropertiesService.getUserProperties().getProperty('ACTIVE_PROJECT_ROW');
    const pid = ss.getSheetByName("Manage").getRange(Number(manageRow), 1).getValue();
    sheet.appendRow([pid, ...rowValues[0]]);
  }
  return getRecordsData();
}

function deleteRecord(rowIdx) {
  if (typeof safeDeleteRow === "function") {
    safeDeleteRow("Records", rowIdx);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Records").deleteRow(rowIdx);
  }
  return { success: true };
}