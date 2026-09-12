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