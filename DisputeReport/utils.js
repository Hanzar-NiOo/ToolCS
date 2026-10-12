/**
 * Duplicates a template Google Sheet file and returns the new copy.
 *
 * @param {string} templateFileId - The Drive File ID of the master template spreadsheet.
 * @param {string} newFileName - The name for the newly created file.
 * @return {Object} { id: string, spreadsheet: SpreadsheetApp.Spreadsheet, file: DriveApp.File }
 */
function createReportFromTemplate(templateFileId, newFileName) {
  if (!templateFileId) {
    throw new Error("Template File ID was not provided.");
  }

  // 1. Locate the master template in Drive
  const templateFile = DriveApp.getFileById(templateFileId);

  // 2. Duplicate the template with the new name
  const copiedFile = templateFile.makeCopy(newFileName);
  const copiedFileId = copiedFile.getId();

  // 3. Open the newly created spreadsheet
  const newSpreadsheet = SpreadsheetApp.openById(copiedFileId);

  return {
    id: copiedFileId,
    spreadsheet: newSpreadsheet,
    file: copiedFile
  };
}
