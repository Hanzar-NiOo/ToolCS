// Replace this with your Google Sheet Template File ID
const DISPUTE_TEMPLATE_FILE_ID = "1prQ7Y9-hxmfcKB8bfHHC80Em32O_8Yqi";

/**
 * Main server-side handler called from UI.
 * Populates the duplicated template and provides the downloadable XLSX.
 */
function generateDisputeReport(params) {
  let tempImportFileId = null;
  let copiedReportId = null;

  try {
    const {
      fileBase64,
      fileName,
      fileType,
      reportDate,
      serialNumber,
      failedCount,
      investigatorName,
      investigatorPosition,
      reportType
    } = params;

    if (!fileBase64) {
      throw new Error("No source transaction file provided.");
    }

    // 1. Construct final report filename
    const cleanSerial = (serialNumber || '001').replace(/[^a-zA-Z0-9_-]/g, '_');
    const category = reportType || 'P2P';
    const outputFileName = `Dispute_Report_${category}_${cleanSerial}.xlsx`;

    // 2. Duplicate the Template file
    const duplicated = createReportFromTemplate(DISPUTE_TEMPLATE_FILE_ID, outputFileName);
    copiedReportId = duplicated.id;
    const spreadsheet = duplicated.spreadsheet;
    const sheet = spreadsheet.getSheets()[0]; // Active report sheet in template

    // 3. Fill in Header & Metadata into the duplicated template
    // (Adjust cell coordinates if your template uses different cells)
    sheet.getRange("B3").setValue(reportDate);
    sheet.getRange("D3").setValue(serialNumber);
    sheet.getRange("B4").setValue(category);
    sheet.getRange("D4").setValue(failedCount || 0);
    sheet.getRange("B5").setValue(`${investigatorName} (${investigatorPosition || 'Staff'})`);
    sheet.getRange("D5").setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"));

    // 4. Parse the uploaded source transactions (CSV or XLSX)
    const decodedBytes = Utilities.base64Decode(fileBase64);
    const uploadedBlob = Utilities.newBlob(decodedBytes, fileType, fileName);
    let rowsToInsert = [];

    const isCsv = fileName.toLowerCase().endsWith('.csv') || (fileType && fileType.includes('csv'));
    if (isCsv) {
      rowsToInsert = Utilities.parseCsv(uploadedBlob.getDataAsString());
    } else {
      // XLSX: convert to temporary Google Sheet using Drive v2
      const tempImport = Drive.Files.insert(
        { title: 'temp_import_' + new Date().getTime(), mimeType: MimeType.GOOGLE_SHEETS },
        uploadedBlob,
        { convert: true }
      );
      tempImportFileId = tempImport.id;
      const tempSheet = SpreadsheetApp.openById(tempImportFileId).getSheets()[0];
      rowsToInsert = tempSheet.getDataRange().getValues();
    }

    // 5. Append transaction rows into the duplicated template
    // Starts at row 9 (after the template headers). Adjust if your template starts on another row.
    const startRow = 9;
    if (rowsToInsert && rowsToInsert.length > 0) {
      sheet.getRange(startRow, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
    }

    // Save and commit all sheet changes
    SpreadsheetApp.flush();

    // 6. Export the duplicated spreadsheet as XLSX
    const exportUrl = `https://docs.google.com/spreadsheets/d/${copiedReportId}/export?format=xlsx`;
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error("Export failed with status code " + response.getResponseCode());
    }

    const xlsxBlob = response.getBlob().setName(outputFileName);
    const xlsxBytes = xlsxBlob.getBytes();

    // 7. Return payload to UI for immediate client-side download
    return {
      name: outputFileName,
      base64: Utilities.base64Encode(xlsxBytes),
      sizeText: (xlsxBytes.length / (1024 * 1024)).toFixed(2) + " MB",
      generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "hh:mm a")
    };

  } catch (err) {
    throw new Error("Report Generation Failed: " + err.message);
  } finally {
    // 8. Clean up temporary files
    if (tempImportFileId) {
      try { Drive.Files.remove(tempImportFileId); } catch (e) {}
    }
    // Clean up duplicated working copy once exported (or remove this line to keep a copy in Drive)
    if (copiedReportId) {
      try { Drive.Files.remove(copiedReportId); } catch (e) {}
    }
  }
}
