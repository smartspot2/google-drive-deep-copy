/**
 * Source folder ID.
 *
 * To find this ID, open the folder in Google Drive; the URL should be of the form
 *
 *      https://drive.google.com/drive/folders/<ID>
 *
 * Copy the <ID> portion of the URL here.
 */
const SOURCE_FOLDER_ID = "12WSpralBgW44ys0oiV6_VbvJY9xakL_Z";

/**
 * Destination folder name.
 *
 * The given source folder will be copied to the root of the Google Drive,
 * renamed with the following name.
 */
const DEST_FOLDER_NAME = "TEST";

/**
 * Temporary state file. Used to keep track of the file IDs that have already been copied over,
 * in case of a script timeout.
 *
 * This file is written to the drive root, and will be deleted after a successful copy.
 */
const TEMP_STATE_FILENAME = "_temp_clone_state.json";

/**
 * Maximum script runtime in ms.
 */
const MAX_RUNTIME_MS = 5 * 60 * 1000;

/**
 * Retry delay after a timeout occurs, in ms.
 */
const RETRY_DELAY_MS = 1000;

/**
 * Global variable; start time of the job.
 */
let jobTimeoutTime;

/**
 * Main function; initial checks and starts the copy process.
 *
 * TODO: forms seem to be duplicated INSIDE the source folder?
 */
function main() {
  jobTimeoutTime = Date.now() + MAX_RUNTIME_MS;

  const rootFolder = DriveApp.getRootFolder();
  const sourceFolder = DriveApp.getFolderById(SOURCE_FOLDER_ID);

  // check the temporary state file
  const copiedFileIds = _readState();

  const destFolderIterator = rootFolder.getFoldersByName(DEST_FOLDER_NAME);

  let destFolder;
  if (copiedFileIds.size == 0) {
    // if not a retry, check that the destination folder doesn't exist already
    if (destFolderIterator.hasNext()) {
      // exists already, fail with error
      throw Error(
        `Destination folder ${DEST_FOLDER_NAME} already exists at the root of the Google Drive!`,
      );
    }

    // create the destination folder
    destFolder = rootFolder.createFolder(DEST_FOLDER_NAME);
  } else {
    // folder should exist already
    if (!destFolderIterator.hasNext()) {
      throw Error(
        `Destination folder ${DEST_FOLDER_NAME} does not exist, but the temporary state file exists!`,
      );
    }

    // use existing destination folder
    destFolder = destFolderIterator.next();
  }

  // create the destination folder
  sourceFolder.getFiles();

  // recursively copy files
  const copySuccessful = copyFiles(copiedFileIds, sourceFolder, destFolder);

  if (!copySuccessful) {
    Logger.log(
      `TIMED OUT! Retrying in ${RETRY_DELAY_MS}ms through a new trigger...`,
    );
    // if unsuccessful, it's because of a timeout; save state to a file
    _saveState(copiedFileIds);

    // set a new trigger to run the next retry
    ScriptApp.newTrigger("main").timeBased().after(RETRY_DELAY_MS).create();
  } else {
    // successful copy, delete temporary state if it exists
    _deleteState();

    // clean up disabled triggers
    // (if this execution was run by a trigger, it'll be left over, though it should be deleted on the next successful execution)
    for (const trigger of ScriptApp.getProjectTriggers()) {
      if (trigger.isDisabled()) {
        ScriptApp.deleteTrigger(trigger);
      }
    }
  }
}

/**
 * Recursively copy files from `sourceFolder` to `destFolder`.
 *
 * TODO: keep track of the exact folder structure and the file IDs;
 * right now, if there are multiple folders with the same name, they'll get merged together
 * (which honestly shouldn't be an issue, but it means this isn't really an exact clone of the source folder)
 *
 * @param {GoogleAppsScript.Drive.Folder} sourceFolder - Source folder to copy files from
 * @param {GoogleAppsScript.Drive.Folder} destFolder - Destination folder to copy files to
 */
function copyFiles(copiedFileIds, sourceFolder, destFolder, history = "") {
  if (_checkTimeout()) {
    return false;
  }

  Logger.log(
    `Copying files in ${history}/${sourceFolder.getName()} [${copiedFileIds.size} items so far]`,
  );

  const sourceFileIterator = sourceFolder.getFiles();
  while (sourceFileIterator.hasNext()) {
    const file = sourceFileIterator.next();
    const fileId = file.getId();
    const fileName = file.getName();

    if (_checkTimeout()) {
      return false;
    }

    if (copiedFileIds.has(fileId)) {
      // already copied; skip
      continue;
    }

    // check mimetype
    const fileMimeType = file.getMimeType();

    if (fileMimeType == MimeType.GOOGLE_SHEETS) {
      // handle spreadsheets separately
      handleCopySpreadsheet(fileId, destFolder);
    } else {
      // make a copy of the file in the destination folder with the same name
      file.makeCopy(fileName, destFolder);
    }

    // add file ID to the current state
    copiedFileIds.add(fileId);
  }

  // return value (accumulated throughout recursive calls)
  let successful = true;

  const sourceFolderIterator = sourceFolder.getFolders();
  while (sourceFolderIterator.hasNext()) {
    const innerSourceFolder = sourceFolderIterator.next();
    const innerSourceFolderId = innerSourceFolder.getId();

    if (_checkTimeout()) {
      return false;
    }

    if (copiedFileIds.has(innerSourceFolderId)) {
      // already copied; skip
      continue;
    }

    const innerDestFolderIterator = destFolder.getFoldersByName(
      innerSourceFolder.getName(),
    );
    let innerDestFolder;
    if (innerDestFolderIterator.hasNext()) {
      // use existing folder
      innerDestFolder = innerDestFolderIterator.next();
    } else {
      // create a folder with the same name in the destination directory
      innerDestFolder = destFolder.createFolder(innerSourceFolder.getName());
    }

    // recurse
    successful = copyFiles(
      copiedFileIds,
      innerSourceFolder,
      innerDestFolder,
      `${history}/${sourceFolder.getName()}`,
    );

    if (!successful) {
      break;
    } else {
      // successful, so add to copied files
      copiedFileIds.add(innerSourceFolderId);
    }
  }

  return successful;
}

/**
 * Special handling for copying a spreadsheet.
 *
 * If a spreadsheet has a linked form, copying the spreadsheet will make a new copy of the form
 * in the ORIGINAL folder. This clutters the source folder, which is undesired.
 *
 * This function will perform the copy (which also copies the form),
 * and then follows the linked form(s) in the duplicate spreadsheet to delete them.
 *
 * @param {number} fileId - ID of the spreadsheet file to copy; assumes that this is actually a spreadsheet
 * @param {DriveApp.Folder} destFolder - Destination folder to copy to
 */
function handleCopySpreadsheet(fileId, destFolder) {
  const spreadsheetFile = DriveApp.getFileById(fileId);
  const spreadsheet = SpreadsheetApp.open(spreadsheetFile);
  const sheets = spreadsheet.getSheets();

  let hasLinkedForm = false;
  for (const sheet of sheets) {
    const assocFormUrl = sheet.getFormUrl();
    if (assocFormUrl !== null) {
      // has linked form
      hasLinkedForm = true;
      break;
    }
  }

  const destSpreadsheetFile = spreadsheetFile.makeCopy(
    spreadsheetFile.getName(),
    destFolder,
  );

  // special handling if there are any linked forms
  if (hasLinkedForm) {
    const destSpreadsheet = SpreadsheetApp.open(destSpreadsheetFile);
    for (const sheet of destSpreadsheet.getSheets()) {
      const assocFormUrl = sheet.getFormUrl();
      if (assocFormUrl !== null) {
        // get the associated form
        const assocForm = FormApp.openByUrl(assocFormUrl);
        // unlink the form
        assocForm.removeDestination();
        // delete the form
        DriveApp.getFileById(assocForm.getId()).setTrashed(true);
      }
    }
  }
}

/**
 * Save the current state of copied file IDs to a file.
 */
function _saveState(copiedFileIds) {
  const jsonIds = JSON.stringify(Array.from(copiedFileIds));
  const rootFolder = DriveApp.getRootFolder();
  const stateFileIterator = rootFolder.getFilesByName(TEMP_STATE_FILENAME);
  if (stateFileIterator.hasNext()) {
    // state file exists; overwrite it
    const stateFile = stateFileIterator.next();
    stateFile.setContent(jsonIds);
  } else {
    // state file does not exist; create it
    rootFolder.createFile(TEMP_STATE_FILENAME, jsonIds);
  }
}

/**
 * Read the current state of copied file IDs.
 */
function _readState() {
  let copiedFileIds = new Set();
  const rootFolder = DriveApp.getRootFolder();
  const stateFileIterator = rootFolder.getFilesByName(TEMP_STATE_FILENAME);
  if (stateFileIterator.hasNext()) {
    // prior state exists; load it
    const stateFile = stateFileIterator.next();
    const stateFileContent = stateFile.getBlob().getDataAsString();
    copiedFileIds = new Set(JSON.parse(stateFileContent));
  }
  return copiedFileIds;
}

/**
 * Delete the saved state.
 */
function _deleteState() {
  const rootFolder = DriveApp.getRootFolder();
  const stateFileIterator = rootFolder.getFilesByName(TEMP_STATE_FILENAME);
  if (stateFileIterator.hasNext()) {
    // prior state exists; delete it
    const stateFile = stateFileIterator.next();
    stateFile.setTrashed(true);
  }
}

function _checkTimeout() {
  return Date.now() >= jobTimeoutTime;
}

