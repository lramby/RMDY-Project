/**
 * SERVICE_INVOICES.gs
 */
function getInvoicesData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Invoices");
    const manageSheet = ss.getSheetByName("Manage");
    const activeRow = PropertiesService.getUserProperties().getProperty('ACTIVE_PROJECT_ROW');
    
    if (!activeRow || !sheet) return [];

    const selectedPid = String(manageSheet.getRange(Number(activeRow), 1).getValue()).trim();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const timezone = ss.getSpreadsheetTimeZone();

    return data.slice(1).map(function(row, index) {
      // Date handling to prevent the "Illegal Value" error
      let d = row[2];
      let formattedDate = d instanceof Date ? Utilities.formatDate(d, timezone, "MM/dd/yyyy") : String(d || "");

      return {
        pid: row[0] ? String(row[0]).trim() : "",
        number: String(row[1] || ""),
        date: formattedDate,
        total: row[7] || 0,
        amountDue: row[8] || 0,
        invoiceId: row[11] || "", // Column L
        note: row[12] || "", // Column L
        sheetRow: index + 2 
      };
    }).filter(inv => inv.pid === selectedPid);
  } catch (e) {
    return [];
  }
}

function saveInvoiceData(obj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Invoices");
    const manageRow = PropertiesService.getUserProperties().getProperty('ACTIVE_PROJECT_ROW');
    const pid = String(ss.getSheetByName("Manage").getRange(Number(manageRow), 1).getValue()).trim();

    const totalNum = parseFloat(String(obj.total).replace(/[$,]/g, '')) || 0;
    const targetRow = obj.row ? Number(obj.row) : 0;
    
    const freshAmountDue = calculateAmountDue(obj.invoiceId, totalNum);

    if (targetRow > 1) {
      // EDIT MODE
      sheet.getRange(targetRow, 2).setValue(obj.number);
      sheet.getRange(targetRow, 3).setValue(obj.date); 
      sheet.getRange(targetRow, 8).setValue(totalNum);
      sheet.getRange(targetRow, 9).setValue(freshAmountDue);
      sheet.getRange(targetRow, 12).setValue(obj.invoiceId);
      sheet.getRange(targetRow, 13).setValue(obj.note);

      // --- NEW CASCADE CALL ---
      // This keeps the Payments sheet display number in sync
      cascadeInvoiceNumberUpdate(obj.invoiceId, obj.number);
      // -------------------------

    } else {
      // NEW MODE
      sheet.appendRow([pid, obj.number, obj.date || new Date(), "", "", "", "", totalNum, freshAmountDue, "", "", obj.invoiceId, obj.note]);
    }
    return getInvoicesData();
  } catch (e) {
    throw new Error("Save Error: " + e.toString());
  }
}

/**
 * Ensure this exact function name exists in Service_Invoices.gs
 */
function deleteInvoiceRow(rowIdx, invId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Invoices");
    
    const rowToDelete = parseInt(rowIdx, 10);
    
    if (isNaN(rowToDelete) || rowToDelete <= 1) {
      throw new Error("Invalid row: " + rowIdx);
    }

    // Safety check for payments
    const paySheet = ss.getSheetByName("Payments");
    if (invId && paySheet) {
      const payData = paySheet.getDataRange().getValues();
      const hasPayment = payData.some(r => String(r[7]).trim() === String(invId).trim());
      if (hasPayment) throw new Error("Cannot delete: payments are attached.");
    }

    sheet.deleteRow(rowToDelete);
    return getInvoicesData();
  } catch (e) {
    throw new Error(e.message);
  }
}


/**
 * Calculates the balance due for a specific Invoice ID
 * Formula: Invoice Total - Sum(Payments with matching Invoice ID)
 */
function calculateAmountDue(invoiceId, invoiceTotal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paySheet = ss.getSheetByName("Payments");
  if (!paySheet) return invoiceTotal;

  const payData = paySheet.getDataRange().getValues();
  if (payData.length < 2) return invoiceTotal;

  // Column H (Index 7) is InvoiceID, Column C (Index 2) is Amount
  const totalPaid = payData.slice(1).reduce((sum, row) => {
    const rowInvId = String(row[7]).trim();
    const amount = parseFloat(row[2]) || 0;
    return rowInvId === String(invoiceId).trim() ? sum + amount : sum;
  }, 0);

  return invoiceTotal - totalPaid;
}

/**
 * Updates the display Invoice Number in the Payments sheet 
 * for all rows matching the given InvoiceID.
 */
function cascadeInvoiceNumberUpdate(invoiceId, newInvoiceNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paySheet = ss.getSheetByName("Payments");
  if (!paySheet) return;

  const data = paySheet.getDataRange().getValues();
  if (data.length < 2) return;

  // Payments layout: Invoice Number is Col B (Index 1), InvoiceID is Col H (Index 7)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][7]).trim() === String(invoiceId).trim()) {
      // Update Column B with the new human-friendly number
      paySheet.getRange(i + 1, 2).setValue(newInvoiceNumber);
    }
  }
}