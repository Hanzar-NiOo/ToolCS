function setTheValue(sheet, headerRow, row, col, rowName, colName, value)
{
  const cellRowName = sheet.getRange(row, 1).getValue().toString().trim();
  let cellColName = sheet.getRange(headerRow, col).getValue().toString().trim();

  if (cellColName === colName)
  {
    if (rowName == null)
    {
      const existingValue = sheet.getRange(row, col).getValue();
      if (!existingValue && existingValue > value)
      {
        throw new Error('Failed to update in "Data Summary for Libeia"');
      }
      sheet.getRange(row, col).setValue(value).setFontFamily("Arial");
    }
    else if (cellRowName === rowName)
    {
      sheet.getRange(row, col).setValue(value).setFontFamily("Calibri");
    }
  }
  else
  {
    throw new Error(`Please check the ${rowName} ${colName}`);
  }
}

function  setValuesInSummaryTable(sheet, headerRow, values, LRD_to_USD)
{
  // Set the values
  setTheValue(sheet, headerRow + 1, headerRow + 2, 2, "P2P", "LRD Transactions", formatCount(P2PLRDTransactionCount));
  setTheValue(sheet, headerRow + 1, headerRow + 2, 3, "P2P", "USD Transactions", formatCount(P2PUSDTransactionCount));
  setTheValue(sheet, headerRow + 1, headerRow + 3, 2, "G2P", "LRD Transactions", formatCount(G2PLRDTransactionCount));
  setTheValue(sheet, headerRow + 1, headerRow + 3, 3, "G2P", "USD Transactions", formatCount(G2PUSDTransactionCount));
  setTheValue(sheet, headerRow + 1, headerRow + 2, 4, "P2P", "LRD Value Processed", formatNumber(P2PLRDTransactionAmount));
  setTheValue(sheet, headerRow + 1, headerRow + 2, 5, "P2P", "LRD Value equilvalent USD(184LRD/USD)", formatNumber(P2PUSDTransactionAmount / LRD_to_USD));
  setTheValue(sheet, headerRow + 1, headerRow + 2, 6, "P2P", "USD Value Processed", formatNumber(P2PUSDTransactionAmount));
  setTheValue(sheet, headerRow + 1, headerRow + 3, 4, "G2P", "LRD Value Processed", formatNumber(G2PLRDTransactionAmount));
  setTheValue(sheet, headerRow + 1, headerRow + 3, 5, "G2P", "LRD Value equilvalent USD(184LRD/USD)", formatNumber(G2PLRDTransactionAmount / LRD_to_USD));
  setTheValue(sheet, headerRow + 1, headerRow + 3, 6, "G2P", "USD Value Processed", formatNumber(G2PUSDTransactionAmount));

  // Set the total values
  setTheValue(sheet, headerRow + 1, headerRow + 4, 2, "Total", "LRD Transactions", formatCount(totalLRDTransactionCount));
  setTheValue(sheet, headerRow + 1, headerRow + 4, 3, "Total", "USD Transactions", formatCount(totalUSDTransactionCount));
  setTheValue(sheet, headerRow, headerRow + 5, 2, "Grand Total", "Transaction Counts", formatCount(totalTransactionCount));
  setTheValue(sheet, headerRow + 1, headerRow + 4, 4, "Total", "LRD Value Processed", formatNumber(totalLRDTransactionAmount));
  setTheValue(sheet, headerRow + 1, headerRow + 4, 5, "Total", "LRD Value equilvalent USD(184LRD/USD)", formatNumber(totalLRDTransactionAmount / LRD_to_USD));
  setTheValue(sheet, headerRow + 1, headerRow + 4, 6, "Total", "USD Value Processed", formatNumber(totalUSDTransactionAmount));
  setTheValue(sheet, headerRow, headerRow + 5, 4, "Grand Total", "Amount", formatNumber(totalTransactionAmount));
}

function toMonthYear(isoDateString) {
    const date = new Date(isoDateString + "T00:00:00Z");
    const month = date.toLocaleString('en-US', { month: 'short', timeZone: 'UTC' });
    const year = String(date.getUTCFullYear()).slice(-2);
    return month + '-' + year;
}

function toMonthYear2(dateInput) {
  var date = new Date(dateInput);
  return Utilities.formatDate(date, "GMT+7", "MMM-yy");
}

function  initiateTheSheet(fileId, sheetId, fromDate)
{
  const spreadsheet = SpreadsheetApp.openById(fileId);
  const sheet   = spreadsheet.getSheetById(sheetId);
  const lastRow = sheet.getLastRow();
  const latestMonth  = sheet.getRange(lastRow, 1).getValue();

  let row = lastRow;
  const formatMonth = toMonthYear2(latestMonth);
  const formatReportMonth = toMonthYear(fromDate);
  if (formatMonth != formatReportMonth)
  {
    row = lastRow + 1;
    sheet.getRange(row, 1).setValue(formatReportMonth).setFontFamily('Arial');
  }

  return [sheet, row, formatReportMonth];
}

function  updateMonthlySheet(fileId, monthlySheetId, values, currentMonth)
{
  [sheet, row] = initiateTheSheet(fileId, monthlySheetId, currentMonth);

  // Set the values in Monthly Table
  setTheValue(sheet, 1, row, 2, null, "P2P Total Transaction(LRD)", formatCount(P2PLRDTransactionCount));
  setTheValue(sheet, 1, row, 3, null, "P2P Total Transaction(USD)", formatCount(P2PUSDTransactionCount));
  setTheValue(sheet, 1, row, 4, null, "P2P Total Amount(LRD)", formatNumber(P2PLRDTransactionAmount));
  setTheValue(sheet, 1, row, 5, null, "P2P Total Amount(USD)", formatNumber(P2PUSDTransactionAmount));
  setTheValue(sheet, 1, row, 6, null, "G2P Total Transaction(LRD)", formatCount(G2PLRDTransactionCount));
  setTheValue(sheet, 1, row, 7, null, "G2P Total Transaction(USD)", formatCount(G2PUSDTransactionCount));
  setTheValue(sheet, 1, row, 8, null, "G2P Total Amount(LRD)", formatNumber(G2PLRDTransactionAmount));
  setTheValue(sheet, 1, row, 9, null, "G2P Total Amount(USD)", formatNumber(G2PUSDTransactionAmount));
  setTheValue(sheet, 1, row, 10, null, "Total LRD", formatNumber(totalLRDTransactionAmount));
  setTheValue(sheet, 1, row, 11, null, "Total USD", formatNumber(totalUSDTransactionAmount));
}

function  updateDFSPsSheet(fileId, DFSPsSheetId, values, currentMonth)
{
  [sheet, row, currentMonth] = initiateTheSheet(fileId, DFSPsSheetId, currentMonth);

  // Set the values in DFSP Table
  setTheValue(sheet, 1, row, 2, null, "OM to MTN(LRD Total Count)", formatCount(OMtoMTN_LRD_Count));
  setTheValue(sheet, 1, row, 3, null, "OM to MTN (LRD Total Value)", formatNumber(OMtoMTN_LRD_Amount));
  setTheValue(sheet, 1, row, 4, null, "OM to MTN(USD Total Count)", formatCount(OMtoMTN_USD_Count));
  setTheValue(sheet, 1, row, 5, null, "OM to MTN (USD Total Value)", formatNumber(OMtoMTN_USD_Amount));
  setTheValue(sheet, 1, row, 6, null, "MTN to OM(LRD Total Count)", formatCount(MTNtoOM_LRD_Count));
  setTheValue(sheet, 1, row, 7, null, "MTN to OM(LRD TotalValue)", formatNumber(MTNtoOM_LRD_Amount));
  setTheValue(sheet, 1, row, 8, null, "MTN to OM(USD Total Count)", formatCount(MTNtoOM_USD_Count));
  setTheValue(sheet, 1, row, 9, null, "MTN to OM(USD Total Value)", formatNumber(MTNtoOM_USD_Amount));
  setTheValue(sheet, 1, row, 10, null, "MFDP to MTN(LRD Total Count)", formatCount(GovToMTN_LRD_Count));
  setTheValue(sheet, 1, row, 11, null, "MFDP to MTN(LRD Total Value)", formatNumber(GovToMTN_LRD_Amount));
  setTheValue(sheet, 1, row, 12, null, "MFDP to MTN(USD Total Count)", formatCount(GovToMTN_USD_Count));
  setTheValue(sheet, 1, row, 13, null, "MFDP to MTN(USD Total Value)", formatNumber(GovToMTN_USD_Amount));
  setTheValue(sheet, 1, row, 14, null, "MFDP to OM(LRD Total Count)", formatCount(GovToOM_LRD_Count));
  setTheValue(sheet, 1, row, 15, null, "MFDP to OM(LRD Total Value)", formatNumber(GovToOM_LRD_Amount));
  setTheValue(sheet, 1, row, 16, null, "MFDP to OM(USD Total Count)", formatCount(GovToOM_USD_Count));
  setTheValue(sheet, 1, row, 17, null, "MFDP to OM(USD Total Value)", formatNumber(GovToOM_USD_Amount));
}
