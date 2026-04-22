
/**
 * Service_Dropdowns.gs
 * Fetches and parses dropdown data. 
 * Expects Column B to be a standard JSON array like ["Value 1", "Value 2"]
 */
function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dropdowns");
  if (!sheet) return {};

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {}; 
  
  const data = sheet.getRange("A2:B" + lastRow).getValues();
  
  const dropdowns = data.reduce((acc, [type, value]) => {
    if (type && value) {
      let rawValue = value.toString().trim();
      
      try {
        // Standard JSON.parse handles double quotes and internal apostrophes perfectly.
        acc[type] = JSON.parse(rawValue);
      } catch (e) {
        console.warn("Parsing failed for " + type + ": Check for standard double quotes.");
        // Fallback: split by comma if JSON parsing fails
        acc[type] = rawValue.replace(/[\[\]"]/g, '').split(',').map(s => s.trim());
      }
    }
    return acc;
  }, {});
  
  return dropdowns;
}