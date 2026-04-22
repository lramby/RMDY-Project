  /**
   * Fetches required fields for a specific form from the Settings sheet.
   */
  function getRequiredFields(formId) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Settings");
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    
    // Filter for Name='RequiredField' and Condition=[Your Form ID]
    return data
      .filter(row => row[0] === "RequiredField" && row[2] === formId)
      .map(row => row[1]); 
  }