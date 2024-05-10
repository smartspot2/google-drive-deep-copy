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
 * Whether to convert all docs/slides/sheets to native Google Drive formats.
 */
const CONVERT_NATIVE_FORMAT = true;

/**
 * Maximum time allowed for exponential backoff.
 *
 * Exponential backoff is used when a Drive API request fails due to a rate limit.
 */
const MAX_BACKOFF_TIME_MS = 32 * 1000;

/**
 * Maximum number of attempts in case of retries during exponential backoff.
 *
 * Exponential backoff is used when a Drive API request fails due to a rate limit.
 */
const MAX_BACKOFF_ATTEMPTS = 20;

/**
 * Debug level for print statements.
 *
 * 0: DEBUG; Print all logs
 * 1: INFO; Print only stage logs (and above)
 * 2: WARN; Print only warnings (and above)
 * 3: ERROR; Print only errors
 */
const DEBUG_LEVEL = 0;

/******
 * GLOBAL VARIABLES
 ******/

/**
 * Start time of the job.
 */
let jobTimeoutTime;

/**
 * Debug level mapping
 */
const DEBUG_LEVELS = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3,
};

/**
 * Conversion mapping.
 *
 * @type {Map<string, string>}
 */
const FORMAT_CONVERSION_MAPPING = new Map(
  Object.entries({
    [MimeType.MICROSOFT_EXCEL]: MimeType.GOOGLE_SHEETS,
    [MimeType.MICROSOFT_EXCEL_LEGACY]: MimeType.GOOGLE_SHEETS,
    [MimeType.MICROSOFT_POWERPOINT]: MimeType.GOOGLE_SLIDES,
    [MimeType.MICROSOFT_POWERPOINT_LEGACY]: MimeType.GOOGLE_SLIDES,
    [MimeType.MICROSOFT_WORD]: MimeType.GOOGLE_DOCS,
    [MimeType.MICROSOFT_WORD_LEGACY]: MimeType.GOOGLE_DOCS,
  }),
);

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

    // generate the file tree
    logWhen(DEBUG_LEVELS.INFO, () =>
      console.info("Exploring file tree to generate structure"),
    );
    fileTree = exploreFileTree(sourceFolder);

    // create the destination folder
    destFolder = rootFolder.createFolder(DEST_FOLDER_NAME);
    fileTree.destId = destFolder.getId();
  } else {
    // folder should exist already
    if (!destFolderIterator.hasNext()) {
      throw new Error(
        `Destination folder ${DEST_FOLDER_NAME} does not exist, but the temporary state file exists!`,
      );
    }

    // use existing destination folder
    logWhen(DEBUG_LEVELS.INFO, () =>
      console.info("Resuming from prior run; using last file tree"),
    );
    destFolder = destFolderIterator.next();
  }

  logWhen(DEBUG_LEVELS.INFO, () => printFileTreeSummary(fileTree));

  logWhen(DEBUG_LEVELS.INFO, () => {
    console.info("Copying folder structure");
    console.time("Copying folders");
  });

  // create folders
  let copySuccessful = createFolders(fileTree);

  logWhen(DEBUG_LEVELS.INFO, () => {
    console.timeEnd("Copying folders");
  });

  if (copySuccessful) {
    if (DEBUG_LEVEL <= DEBUG_LEVELS.INFO) {
      console.info("Copying files");
      console.time("Copying files");
    }

    // copy files
    copySuccessful = copyFiles(fileTree);

    if (DEBUG_LEVEL <= DEBUG_LEVELS.INFO) {
      console.timeEnd("Copying files");
    }
  }

  if (!copySuccessful) {
    logWhen(DEBUG_LEVELS.WARN, () =>
      console.warn(
        `TIMED OUT! Retrying in ${RETRY_DELAY_MS}ms through a new trigger...`,
      ),
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
    logWhen(DEBUG_LEVELS.INFO, () => {
      console.info("Done copying!");
      printFileTreeSummary(fileTree);
    });
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
 * FileState type used in future functions.
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
    logWhen(DEBUG_LEVELS.DEBUG, () =>
      console.log(`Creating folder ${newHistory}/${folderState.name}`),
    );
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

    const file = DriveApp.getFileById(fileObj.id);
    const fileId = fileObj.id;
    const fileName = fileObj.name;
    const fileMimeType = file.getMimeType();

    // copy file
    let copiedFile;
    if (fileMimeType == MimeType.GOOGLE_SHEETS) {
      // handle spreadsheets separately
      logWhen(DEBUG_LEVELS.DEBUG, () =>
        console.log(
          `Copying ${newHistory}/${fileObj.name} (Google Sheets handling)`,
        ),
      );
      copiedFile = handleCopySpreadsheet(fileId, parentFolder);
    } else if (FORMAT_CONVERSION_MAPPING.has(fileMimeType)) {
      // handle conversion between formats
      logWhen(DEBUG_LEVELS.DEBUG, () =>
        console.log(
          `Copying ${newHistory}/${fileObj.name} (Format conversion handling)`,
        ),
      );
      copiedFile = handleFormatConversion(fileId, parentFolder);
    } else {
      // make a copy of the file in the destination folder with the same name
      logWhen(DEBUG_LEVELS.DEBUG, () =>
        console.log(`Copying ${newHistory}/${fileObj.name}`),
      );
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
 * @param {GoogleAppsScript.Drive.Folder} destFolder - Destination folder to copy to
 *
 * @returns {GoogleAppsScript.Drive.File}
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
 * Special handling for copying with format conversion.
 *
 * @param {string} fileId - ID of the file to copy
 * @param {GoogleAppsScript.Drive.Folder} destFolder - Destination folder to copy to
 *
 * @returns {GoogleAppsScript.Drive.File}
 */
function handleFormatConversion(fileId, destFolder) {
  const file = DriveApp.getFileById(fileId);
  const fileBlob = file.getBlob();

  const fileMimeType = file.getMimeType();

  if (FORMAT_CONVERSION_MAPPING.has(fileMimeType)) {
    // get converted MIME type
    const convertedMimeType = FORMAT_CONVERSION_MAPPING.get(fileMimeType);
    // create the new file in the destination folder
    return withBackoff(() =>
      Drive.Files.create(
        {
          name: file.getName(),
          parents: [destFolder.getId()],
          mimeType: convertedMimeType,
        },
        fileBlob,
      ),
    );
  } else {
    return defaultCopy(file, destFolder);
  }
}

/**
 * Default copy operation.
 *
 * @param {GoogleAppsScript.Drive.File} file - File to copy
 * @param {GoogleAppsScript.Drive.Folder} destFolder - Destination folder to copy to
 *
 * @returns {GoogleAppsScript.Drive.File} Copied file
 */
function defaultCopy(file, destFolder) {
  return file.makeCopy(file.getName(), destFolder);
}

/**
 * Print a summary of the file tree.
 *
 * @param {FolderState} fileTree - File tree to print summary details for
 */
function printFileTreeSummary(fileTree) {
  let numFilesCopied = 0;
  let numFoldersCopied = 0;
  let totalFileCount = 0;
  let totalFolderCount = 0;

  /**
   *  Helper function for recursively traversing the file tree.
   *
   *  @param {FolderState} tree
   */
  const _traverse = (tree) => {
    // stats for child files
    for (const file of tree.files) {
      numFilesCopied += file.destId != null;
      totalFileCount++;
    }

    // stats for this folder
    numFoldersCopied += tree.destId != null;
    totalFolderCount++;

    // stats for child folders
    for (const folder of tree.folders) {
      _traverse(folder);
    }
  };

  // collect data
  _traverse(fileTree);

  // print data
  logWhen(DEBUG_LEVELS.INFO, () => {
    console.info(
      `Total copied: ${numFilesCopied + numFoldersCopied} / ${totalFileCount + totalFolderCount}\n` +
        `Copied files: ${numFilesCopied} / ${totalFileCount}\n` +
        `Copied folders: ${numFoldersCopied} / ${totalFolderCount}`,
    );
  });
}

/**
 * Exponential backoff, as described in the Google API docs.
 *
 * Calls the callback function repeatedly until no errors occur.
 *
 * @param {() => void} callback
 */
function withBackoff(callback) {
  let attempts = 0;

  while (attempts <= MAX_BACKOFF_ATTEMPTS) {
    try {
      return callback();
    } catch (e) {
      console.error(e);

      const delayTime = Math.min(
        MAX_BACKOFF_TIME_MS,
        2 ** attempts + Math.random() * 1000,
      );

      Utilities.sleep(delayTime);
    }

    attempts++;
  }
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

/**
 * Call the callback if the current log level meets the given log level.
 *
 * @param {number} logLevel - Log level condition
 * @param {() => void} callback - Function to call if the log level is met
 */
function logWhen(logLevel, callback) {
  if (DEBUG_LEVEL <= logLevel) {
    callback();
  }
}
