// var MAIN_FOLDER_ID = "1MobTXre3RodB-C4rb00R1mMcGga4njld";

// function uploadToMonthFolder(xlsxBlob, reportName, startDate) {
//   try
//   {
//     var monthFolderName = Utilities.formatDate(
//       new Date(startDate),
//       Session.getScriptTimeZone(),
//       "MMMM yyyy"
//     );

//     var mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
//     var existingFolders = mainFolder.getFoldersByName(monthFolderName);
//     var monthFolder;

//     if (existingFolders.hasNext()) {
//       monthFolder = existingFolders.next();
//     } else {
//       monthFolder = mainFolder.createFolder(monthFolderName);
//     }

//     var existingFiles = monthFolder.getFilesByName(reportName);
//     while (existingFiles.hasNext()) {
//       existingFiles.next().setTrashed(true);
//     }

//     var uploadedFile = monthFolder.createFile(xlsxBlob);

//     return uploadedFile;
//   }
//   catch
//   {
//     throw new Error("Failed to upload the file");
//   }
// }

function formatMonthDay(dateString)
{
  const date = new Date(dateString);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM dd");
}

function uploadBackupFile(fileId, folderId)
{
  try
  {
    const originalFile = Drive.Files.get(fileId);

    const backupFile = Drive.Files.copy(
      {
        title: "Backup_for_" + originalFile.title,
        parents: [{ id: folderId }]
      },
      fileId
    );

    return backupFile;
  }
  catch (e)
  {
    throw new Error("Failed to backup 'Data Summary for Liberia': " + e.message);
  }
}

function renameOriginalFile(fileId, fromDate, toDate)
{
  try
  {
    const formattedFrom = formatMonthDay(fromDate);
    const formattedTo = formatMonthDay(toDate);
    const newTitle = "Data Summary for Liberia (" + formattedTo + " Update)";

    const renamedFile = Drive.Files.update(
      { title: newTitle },
      fileId
    );

    return renamedFile;
  }
  catch (e)
  {
    throw new Error("Failed to rename 'Data Summary for Liberia': " + e.message);
  }
}
