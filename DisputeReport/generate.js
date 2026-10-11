/**
 * Main server-side function for generating Dispute Reports.
 * Called from index.html via google.script.run.generateDisputeReport(params)
 *
 * @param {Object} params - Dispute form parameters and file payload.
 * @return {Object} Result object containing filename, base64 data, size, and timestamp.
 */
function generateDisputeReport(params) {
  let tempImportFileId = null;
  let reportSpreadsheetId = null;

  try {
    // 1. Destructure incoming parameters
    const {
      fileBase64,
      fileName,
      fileType,
      reportDate,
      serialNumber,
      failedCount,
      investigatorName,
      investigatorPosition,
      reportType,
      updateDatabase
    } = params;

    if (!fileBase64) {
      throw new Error("No source file was provided for report generation.");
    }

    // 2. Decode the uploaded file blob
    const decodedBytes = Utilities.base64Decode(fileBase64);
    const uploadedBlob = Utilities.newBlob(decodedBytes, fileType, fileName);

    // 3. Extract transaction rows from CSV or XLSX
    let sourceData = [];
    const isCsv = fileName.toLowerCase().endsWith('.csv') || (fileType && fileType.includes('csv'));

    if (isCsv) {
      const csvString = uploadedBlob.getDataAsString();
      sourceData = Utilities.parseCsv(csvString);
    } else {
      // XLSX: Use Drive v2 Advanced Service to convert to temporary Google Sheets
      const tempImportFile = Drive.Files.insert(
        {
          title: 'temp_dispute_import_' + new Date().getTime(),
          mimeType: MimeType.GOOGLE_SHEETS
        },
        uploadedBlob,
        { convert: true }
      );
      tempImportFileId = tempImportFile.id;

      const tempSpreadsheet = SpreadsheetApp.openById(tempImportFileId);
      const tempSheet = tempSpreadsheet.getSheets()[0];
      sourceData = tempSheet.getDataRange().getValues();
    }

    // 4. Create the Dispute Report Spreadsheet
    const cleanSerial = (serialNumber || '001').replace(/[^a-zA-Z0-9_-]/g, '_');
    const typeLabel = reportType || 'P2P';
    const outputFileName = `Dispute_Report_${typeLabel}_${cleanSerial}.xlsx`;

    const reportSpreadsheet = SpreadsheetApp.create(`Dispute_Report_${typeLabel}_${cleanSerial}`);
    reportSpreadsheetId = reportSpreadsheet.getId();
    const sheet = reportSpreadsheet.getActiveSheet();
    sheet.setName("Dispute Summary");

    // 5. Build Header & Metadata Section
    sheet.getRange("A1:G1").merge()
      .setValue("DISPUTE INVESTIGATION & SETTLEMENT REPORT")
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground("#131b2e")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sheet.setRowHeight(1, 40);

    const metadata = [
      ["Report Date:", reportDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"), "Serial Number:", serialNumber || "N/A"],
      ["Report Type:", typeLabel, "Failed Count (Synced):", failedCount || 0],
      ["Investigator:", `${investigatorName || 'N/A'} (${investigatorPosition || 'Staff'})`, "Protocol:", "ISO-8583 Compliant / MTN-MFDP to OML"],
      ["Generated At:", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"), "Audit Log Status:", updateDatabase ? "Synced" : "Manual"]
    ];

    const metaRange = sheet.getRange(3, 1, metadata.length, 4);
    metaRange.setValues(metadata);
    sheet.getRange(3, 1, metadata.length, 1).setFontWeight("bold").setFontColor("#45464d");
    sheet.getRange(3, 3, metadata.length, 1).setFontWeight("bold").setFontColor("#45464d");

    // 6. Populate Reconciled Source Transactions
    const startRow = 8;
    if (sourceData && sourceData.length > 0) {
      const numRows = sourceData.length;
      const numCols = sourceData[0].length;
      const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
      dataRange.setValues(sourceData);

      // Header row styling
      const headerRowRange = sheet.getRange(startRow, 1, 1, numCols);
      headerRowRange.setFontWeight("bold")
        .setBackground("#0051d5")
        .setFontColor("#ffffff")
        .setHorizontalAlignment("center");
      sheet.setRowHeight(startRow, 28);

      // Border and grid formatting
      dataRange.setBorder(true, true, true, true, true, true, "#c6c6cd", SpreadsheetApp.BorderStyle.SOLID);
    } else {
      sheet.getRange(startRow, 1).setValue("No transaction data found in source file.");
    }

    // Auto-fit columns
    const maxCols = Math.max(7, sourceData && sourceData[0] ? sourceData[0].length : 4);
    for (let c = 1; c <= maxCols; c++) {
      sheet.autoResizeColumn(c);
    }

    // Flush changes to ensure content is saved
    SpreadsheetApp.flush();

    // 7. Export Spreadsheet as XLSX binary
    const exportUrl = "https://docs.google.com/spreadsheets/d/" + reportSpreadsheetId + "/export?format=xlsx";
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error("Failed to export report as XLSX. Error code: " + response.getResponseCode());
    }

    const xlsxBlob = response.getBlob().setName(outputFileName);
    const xlsxBytes = xlsxBlob.getBytes();
    const resultBase64 = Utilities.base64Encode(xlsxBytes);

    // 8. Return result object to UI
    return {
      name: outputFileName,
      base64: resultBase64,
      sizeText: formatBytes_(xlsxBytes.length),
      generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "hh:mm a")
    };

  } catch (error) {
    throw new Error("Report Generation Failed: " + error.message);
  } finally {
    // 9. Clean up temporary Drive files
    if (tempImportFileId) {
      try { Drive.Files.remove(tempImportFileId); } catch (e) {}
    }
    if (reportSpreadsheetId) {
      try { Drive.Files.remove(reportSpreadsheetId); } catch (e) {}
    }
  }
}

/**
 * Helper to format byte sizes into KB / MB.
 */
function formatBytes_(bytes) {
  if (bytes < 1024 * 1024) {
    return (bytes / 1024).toFixed(1) + " KB";
  }
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}
