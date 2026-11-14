// ----------------------
// Backup Zone -->
// ----------------------
function backupSheet() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = sheet.getName();
    var backupFolderId = PropertiesService.getScriptProperties().getProperty("backupFolderId") // Folder ID
    var folder = DriveApp.getFolderById(backupFolderId);
    var file = DriveApp.getFileById(sheet.getId());
    
    // Get current date and time in Asia/Kolkata timezone
    var now = new Date();
    var timeZone = 'Asia/Kolkata';
    var formattedDate = Utilities.formatDate(now, timeZone, 'yyyyMMdd-HHmmss');
    
    // Find and delete previous backups
    var files = folder.getFiles();
    while (files.hasNext()) {
      var existingFile = files.next();
      var existingFileName = existingFile.getName();
      
      // Check if the file name starts with 'Master Sheet Backup'
      if (existingFileName.indexOf('Master Sheet Backup') === 0) {
        existingFile.setTrashed(true); // Move the file to trash
      }
    }
    
    // Create the new backup
    file.makeCopy('Master Sheet Backup ' + formattedDate, folder);
}