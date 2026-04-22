/**
 * Service_Attachments.gs
 * Fetches data, provisions folders, and crawls subfolders for files.
 */
function getAttachmentsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attachments");
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
  
  if (!rowIndex) return null;
  const targetRow = parseInt(rowIndex);
  
  // Guard against range errors
  if (targetRow < 1 || targetRow > sheet.getLastRow()) return null;

  const displayValues = sheet.getRange(targetRow, 1, 1, 5).getDisplayValues()[0];
  const pid = displayValues[0];
  let folderId = displayValues[4];

  // Provision folders if Column E is empty
  if (!folderId || folderId === "") {
    folderId = provisionProjectFolders(pid, targetRow);
  }

  const fileList = [];
  try {
    const parentFolder = DriveApp.getFolderById(folderId);
    const subFolders = parentFolder.getFolders();
    
    // Crawl subfolders (Documents/Images) for files
    while (subFolders.hasNext()) {
      const sub = subFolders.next();
      const folderName = sub.getName();
      const files = sub.getFiles();
      
      while (files.hasNext()) {
        const file = files.next();
        fileList.push({
          name: file.getName(),
          type: folderName,
          url: file.getUrl()
        });
      }
    }
  } catch (e) {
    console.log("Drive Error: " + e.message);
  }

  return {
    pid: pid,
    contract: displayValues[1],
    photos: displayValues[2],
    measurements: displayValues[3],
    folderLink: `https://drive.google.com/drive/folders/${folderId}`,
    files: fileList
  };
}

function provisionProjectFolders(pid, targetRow) {
  const parentFolderId = "19xcSknfOjf-9OD6EP3lKXXq2no-hWAEw";
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const folderSearch = parentFolder.getFoldersByName(pid);
  
  let projectFolder;
  if (folderSearch.hasNext()) {
    projectFolder = folderSearch.next();
  } else {
    projectFolder = parentFolder.createFolder(pid);
    projectFolder.createFolder("Documents");
    projectFolder.createFolder("Images");
  }
  
  const newId = projectFolder.getId();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attachments");
  sheet.getRange(targetRow, 5).setValue(newId);
  
  return newId;
}