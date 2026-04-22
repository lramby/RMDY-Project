/**
 * GET PAYMENTS
 * Fetches all payments for the active project.
 */
function getPaymentsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Payments");
  const manageSheet = ss.getSheetByName("Manage");
  const rowIndex = parseInt(PropertiesService.getUserProperties().getProperty('ACTIVE_PROJECT_ROW'), 10);
  
  if (!sheet || isNaN(rowIndex) || rowIndex < 1) return [];

  const selectedPid = manageSheet.getRange(rowIndex, 1).getDisplayValue(); 
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return [];

  // Expanded to 8 columns to capture Column H (InvoiceID)
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
  
  return data
    .filter(row => row[0] === selectedPid) 
    .map((row, index) => ({
      sheetRow:  index + 2, 
      pid:       row[0], // A
      invoice:   row[1], // B
      amount:    row[2], // C
      date:      row[3], // D
      note:      row[4], // E
      method:    row[5], // F
      paymentId: row[6], // G
      invoiceId: row[7]  // H: The missing link
    }));
}

/**
 * ADD PAYMENT
 * Generates a unique PaymentID and appends the row.
 */
function addPayment(payObj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Payments");
    const manageSheet = ss.getSheetByName("Manage");
    const rowIndex = parseInt(PropertiesService.getUserProperties().getProperty('ACTIVE_PROJECT_ROW'), 10);
    
    const pid = manageSheet.getRange(rowIndex, 1).getDisplayValue();
    const paymentId = "PAY-" + new Date().getTime().toString().slice(-6);

    sheet.appendRow([
      pid, 
      payObj.invoice, 
      payObj.amount, 
      payObj.date || new Date(), 
      payObj.note, 
      payObj.method, 
      paymentId,
      payObj.invoiceId 
    ]);

    // --- SYNC START ---
    syncInvoiceBalance(payObj.invoiceId);
    // --- SYNC END ---
    
    return getPaymentsData(); 
  } catch (e) {
    throw new Error(e.message);
  }
}


/**
 * Fetches only the invoice numbers for the active project
 * to populate the dropdown in the Payment Modal.
 */
function getInvoiceNumbersForActiveProject() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName("Invoices");
  const manageSheet = ss.getSheetByName("Manage");
  const rowIndex = PropertiesService.getUserProperties().getProperty('ACTIVE_PROJECT_ROW');
  
  if (!invSheet || !rowIndex) return [];
  
  const selectedPid = manageSheet.getRange(rowIndex, 1).getDisplayValue().trim();
  const lastRow = invSheet.getLastRow();
  if (lastRow < 2) return [];

  // Get PiD (A), Invoice Num (B), and InvoiceID (L is column 12)
  const data = invSheet.getRange(2, 1, lastRow - 1, 12).getValues(); 
  
  return data
    .filter(row => String(row[0]).trim() === selectedPid)
    .map(row => ({
      number: row[1],
      id: row[11] // Column L
    }));
}



/**
 * Updates an existing payment record
 */
function updatePayment(obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Payments");
  const row = parseInt(obj.sheetRow, 10);
  
  if (!sheet || isNaN(row) || row < 2) return false;

  sheet.getRange(row, 2).setValue(obj.invoice);
  sheet.getRange(row, 3).setValue(obj.amount);
  sheet.getRange(row, 4).setValue(obj.date);
  sheet.getRange(row, 5).setValue(obj.note);
  sheet.getRange(row, 6).setValue(obj.method);
  
  // --- SYNC START ---
  // Assuming the invoiceId is stored in Column H (Index 8)
  const invId = sheet.getRange(row, 8).getValue();
  syncInvoiceBalance(invId);
  // --- SYNC END ---
  
  return true;
}


/**
 * DELETE PAYMENT
 * Removes the row and returns the updated list.
 */
function deletePayment(rowIdx) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Payments");
    const rowToDelete = parseInt(rowIdx, 10);

    if (isNaN(rowToDelete) || rowToDelete < 2) throw new Error("Invalid row");

    // --- SYNC PREP: Get ID before deleting row ---
    const invId = sheet.getRange(rowToDelete, 8).getValue();
    
    sheet.deleteRow(rowToDelete);
    
    // --- SYNC START ---
    syncInvoiceBalance(invId);
    // --- SYNC END ---
    
    return getPaymentsData();
  } catch (e) {
    throw new Error("Delete failed: " + e.toString());
  }
}


function syncInvoiceBalance(invoiceId) {
  if (!invoiceId) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName("Invoices");
  const data = invSheet.getDataRange().getValues();
  
  // Find the row in Invoices where Col L (Index 11) matches the InvoiceID
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][11]).trim() === String(invoiceId).trim()) {
      const total = parseFloat(data[i][7]) || 0;
      const newDue = calculateAmountDue(invoiceId, total);
      invSheet.getRange(i + 1, 9).setValue(newDue); // Update Col I
      break;
    }
  }
}