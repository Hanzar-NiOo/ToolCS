const fileId = "1gsSy1pJWEoUM52JoSvapimWZQVmHozHs";
const monthlySheetId = "1058338125";
const DFSPsSheetId = "2037740635";
const backupFolderID = "1gtsp5Su8uf-jfuEZ-_U_yO_ZTJPrRFVy";

function mgnSummaryReport(base64Data, fileName, mimeType, fromDate, toDate, dataUpdate, abortedTransactions, exchangeRate)
{
  // Initiate the sheet
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const tempConvertedFile = Drive.Files.insert(
    {
      title: "Temp_" + fileName,
      mimeType: MimeType.GOOGLE_SHEETS
    },
    blob
  );
  const spreadsheet = SpreadsheetApp.openById(tempConvertedFile.id);
  const sheet = spreadsheet.getSheets()[0];
  const theLatestRow = sheet.getLastRow();
  headerRow = theLatestRow + 2;

  values = getTheValues(sheet, theLatestRow, exchangeRate);
  createSummaryTableStructure(sheet, headerRow);
  addAbortedTransactionsBox(sheet, headerRow + 7, abortedTransactions);
  setValuesInSummaryTable(sheet, headerRow, values, exchangeRate);
  
  if (dataUpdate)
  {
    uploadBackupFile(fileId, backupFolderID);
    renameOriginalFile(fileId, fromDate, toDate)
    updateDFSPsSheet(fileId, DFSPsSheetId, values, fromDate);
    updateMonthlySheet(fileId, monthlySheetId, values, fromDate, abortedTransactions);
  }

  SpreadsheetApp.flush();

  try
  {
    var xlsxBlob = exportSheetAsXlsx(spreadsheet.getId());
    const formattedfromDate = extractDayMonth(fromDate);
    const formattedtoDate = extractDayMonth(toDate);

    var reportName = "ManagementSummaryReport-" + formattedfromDate + "to" + formattedtoDate + ".xlsx";
    xlsxBlob.setName(reportName);
    var xlsxBytes = xlsxBlob.getBytes();
    var base64Out = Utilities.base64Encode(xlsxBytes);
    var sizeKB = Math.round((xlsxBytes.length / 1024) * 10) / 10;
    var generatedAt = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "MMM d, yyyy HH:mm"
    );

    // if (fileUpload)
    // {
    //   uploadToMonthFolder(xlsxBlob, reportName, fromDate);
    // }
    
    // Clean up the temp Google Sheet from Drive
    Drive.Files.remove(tempConvertedFile.id);

    return {
      base64: base64Out,
      name: reportName,
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      sizeText: sizeKB + " KB",
      generatedAt: generatedAt
    };
  }
  catch
  {
    throw new Error("Failed to generate the report")
  }
}

function getCsvFile()
{
  try
  {
    const url = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + token }
    });
    const blob = response.getBlob();
    const bytes = blob.getBytes();
    const base64 = Utilities.base64Encode(bytes);

    const file = DriveApp.getFileById(fileId);

    return {
      base64: base64,
      name: file.getName() + ".xlsx",
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    };
  }
  catch (e)
  {
    throw new Error("Failed to download file: " + e.message);
  }
}