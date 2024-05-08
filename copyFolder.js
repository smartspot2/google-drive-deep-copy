/**
 * Source folder ID.
 *
 * To find this ID, open the folder in Google Drive; the URL should be of the form
 *
 *      https://drive.google.com/drive/folders/<ID>
 *
 * Copy the <ID> portion of the URL here.
 */
const SOURCE_FOLDER_ID = "1UO5MWPn4ylPaZHdZLTgxqw55kFUmaFMT";

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
 */
function main() {
  jobTimeoutTime = Date.now() + MAX_RUNTIME_MS;

  const rootFolder = DriveApp.getRootFolder();
  const sourceFolder = DriveApp.getFolderById(SOURCE_FOLDER_ID);

  // check the temporary state file
  let fileTree = _readState();

  const destFolderIterator = rootFolder.getFoldersByName(DEST_FOLDER_NAME);

  let destFolder;
  if (fileTree == null) {
    // if not a retry, check that the destination folder doesn't exist already
    if (destFolderIterator.hasNext()) {
      // exists already, fail with error
      throw new Error(
        `Destination folder ${DEST_FOLDER_NAME} already exists at the root of the Google Drive!`,
      );
    }

    // create the destination folder
    destFolder = rootFolder.createFolder(DEST_FOLDER_NAME);

    // generate the file tree
    Logger.log("Exploring file tree to generate structure");
    fileTree = exploreFileTree(sourceFolder);
    fileTree.destId = destFolder.getId();
  } else {
    // folder should exist already
    if (!destFolderIterator.hasNext()) {
      throw new Error(
        `Destination folder ${DEST_FOLDER_NAME} does not exist, but the temporary state file exists!`,
      );
    }

    // use existing destination folder
    destFolder = destFolderIterator.next();
  }

  // create folders
  let copySuccessful = createFolders(fileTree);

  if (copySuccessful) {
    // copy files
    copySuccessful = copyFiles(fileTree);
  }

  if (!copySuccessful) {
    Logger.log(
      `TIMED OUT! Retrying in ${RETRY_DELAY_MS}ms through a new trigger...`,
    );
    // if unsuccessful, it's because of a timeout; save state to a file
    _saveState(fileTree);

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
 * FolderState type used in future functions.
 *
 * Contains information about a Google Drive folder,
 * used to keep track of state when copying folders and files.
 *
 * `foldersCopied`: whether all child folders have been successfully copied (including recursive children)
 * `filesCopied`: whether all child files have been successfully copied (including recursive children)
 *
 * @typedef {{
 *    id: string,
 *    name: string,
 *    filesCopied: boolean,
 *    foldersCopied: boolean,
 *    files: FileState[],
 *    folders: FolderState[],
 *    destId: null | string,
 * }} FolderState
 */

/**
 * FileSstate type used in future functions.
 *
 * Contains information about a Google Drive file,
 * used to keep track of state when copying.
 *
 * `copied`: whether this file has been successfully copied
 *
 * @typedef {{
 *    id: string,
 *    name: string,
 *    destId: null | string,
 * }} FileState
 */

/**
 * Explore the folder structure of a given folder.
 *
 * @param {GoogleAppsScript.Drive.Folder} folder - Folder to explore
 *
 * @returns {FolderState}
 */
function exploreFileTree(folder) {
  const childrenFiles = [];
  const childrenFileIterator = folder.getFiles();
  while (childrenFileIterator.hasNext()) {
    const childFile = childrenFileIterator.next();
    childrenFiles.push({
      id: childFile.getId(),
      name: childFile.getName(),
      destId: null,
    });
  }

  const childrenFolders = [];
  const childrenFolderIterator = folder.getFolders();
  while (childrenFolderIterator.hasNext()) {
    const childFolder = childrenFolderIterator.next();
    childrenFolders.push(exploreFileTree(childFolder));
  }

  return {
    id: folder.getId(),
    name: folder.getName(),
    // flag for whether all children files (recursively) have been copied
    filesCopied: false,
    // flag for whether all children folders (recusrively) have been copied
    foldersCopied: false,
    // destination folder; null if not copied,
    // non-null if this folder (not necessarily its contents) has been copied already
    destId: null,
    // files are stored as FileState objects
    files: childrenFiles,
    // recursive structure for folders
    folders: childrenFolders,
  };
}

/**
 * Copy all folders in a file tree in preparation for file copying.
 *
 * Assumes that the root folder has already been copied over.
 * The `destFolder` corresponds to this copied root folder.
 *
 * @param {FolderState} fileTree - File tree to copy
 * @param {GoogleAppsScript.Drive.Folder} destFolder - Destination folder; corresopnds to the `fileTree` root folder.
 *
 * @returns {boolean} - Whether the folders have all been successully copied. If false, a timeout occurred.
 */
function createFolders(fileTree, history = "") {
  const newHistory = `${history}/${fileTree.name}`;
  if (fileTree.foldersCopied) {
    // if already done, skip
    return true;
  }

  if (_checkTimeout()) {
    return false;
  }

  const destFolderId = fileTree.destId;
  const destFolder = DriveApp.getFolderById(destFolderId);

  let creationSuccessful = true;

  for (const folderState of fileTree.folders) {
    if (folderState.destId != null) {
      // skip if folder already copied
      continue;
    }

    // create folder with the same name
    Logger.log(`Creating folder ${newHistory}/${folderState.name}`);
    const folderName = folderState.name;
    const newFolder = destFolder.createFolder(folderName);
    folderState.destId = newFolder.getId();

    // recurse on the child folder
    creationSuccessful = createFolders(folderState, newHistory);

    if (!creationSuccessful) break;
  }

  if (creationSuccessful) {
    // set flag if all child folder creation was successful
    fileTree.foldersCopied = true;
  }

  return creationSuccessful;
}

/**
 * Recursively copy files from `sourceFolder` to `destFolder`.
 *
 * TODO: keep track of the exact folder structure and the file IDs;
 * right now, if there are multiple folders with the same name, they'll get merged together
 * (which honestly shouldn't be an issue, but it means this isn't really an exact clone of the source folder)
 *
 * @param {FolderState} fileTree - File tree to copy over
 *
 * @returns {boolean} - Whether or not files were copied successfully. If false, a timeout occurred.
 */
function copyFiles(fileTree, history = "") {
  const newHistory = `${history}/${fileTree.name}`;

  if (fileTree.filesCopied) {
    return true;
  }

  if (_checkTimeout()) {
    return false;
  }

  const parentFolder = DriveApp.getFolderById(fileTree.destId);

  for (const fileObj of fileTree.files) {
    if (fileObj.destId != null) {
      // skip if already copied
      continue;
    }

    if (_checkTimeout()) {
      return false;
    }

    Logger.log(`Copying ${newHistory}/${fileObj.name}`);

    const file = DriveApp.getFileById(fileObj.id);
    const fileId = fileObj.id;
    const fileName = fileObj.name;
    const fileMimeType = file.getMimeType();

    // copy file
    let copiedFile;
    if (fileMimeType == MimeType.GOOGLE_SHEETS) {
      // handle spreadsheets separately
      copiedFile = handleCopySpreadsheet(fileId, parentFolder);
    } else {
      // make a copy of the file in the destination folder with the same name
      copiedFile = file.makeCopy(fileName, parentFolder);
    }

    fileObj.destId = copiedFile.getId();
  }

  // recurse on all child folders
  let copySuccessful = true;
  for (const childFileTree of fileTree.folders) {
    copySuccessful = copyFiles(childFileTree, newHistory);
    if (!copySuccessful) break;
  }

  if (copySuccessful) {
    // set flag if all child files were successfully copied
    fileTree.filesCopied = true;
  }

  return copySuccessful;
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
 *
 * @returns {DriveApp.File}
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

  return destSpreadsheetFile;
}

/**
 * Save the current state of copied file IDs to a file.
 *
 * @param {FolderState} fileTree - file tree to store
 */
function _saveState(fileTree) {
  const jsonIds = JSON.stringify(fileTree);
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
 *
 * @returns {FolderState | null}
 */
function _readState() {
  const rootFolder = DriveApp.getRootFolder();
  const stateFileIterator = rootFolder.getFilesByName(TEMP_STATE_FILENAME);
  let fileTree = null;
  if (stateFileIterator.hasNext()) {
    // prior state exists; load it
    const stateFile = stateFileIterator.next();
    const stateFileContent = stateFile.getBlob().getDataAsString();
    fileTree = JSON.parse(stateFileContent);
  }
  return fileTree;
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
