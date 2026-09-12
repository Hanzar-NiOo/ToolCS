const fileId = "1gsSy1pJWEoUM52JoSvapimWZQVmHozHs";
const monthlySheetId = "1058338125";
const DFSPsSheetId = "2037740635";
const LRD_to_USD = 184;

// Test
abortedCount = "";

function mgnSummaryReport(base64Data, fileName, mimeType, fromDate, toDate, dataUpdate)
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

  values = getTheValues(sheet, theLatestRow, LRD_to_USD);
  createSummaryTableStructure(sheet, headerRow);
  addAbortedTransactionsBox(sheet, headerRow + 7, abortedCount);
  setValuesInSummaryTable(sheet, headerRow, values, LRD_to_USD);
  
  if (dataUpdate)
  {
    updateDFSPsSheet(fileId, DFSPsSheetId, values, fromDate);
    updateMonthlySheet(fileId, monthlySheetId, values, fromDate);
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

function getCsvFile() {
  try
  {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var bytes = blob.getBytes();
    var base64 = Utilities.base64Encode(bytes);

    return {
      base64: base64,
      name: file.getName(),
      mimeType: "text/csv"
    };
  } catch (e) {
    throw new Error("Failed to download");
  }
}