// sheets-automation.js
import fs from "fs";
import path from "path";
import { google } from "googleapis";
import { createAuthClient } from "./auth.js";
import { parse } from "csv-parse/sync";

/**
 * Main function to create a new spreadsheet based on a template and perform data operations
 * @param {Object} config - Configuration object containing all parameters
 * @returns {Promise<string>} - URL of the newly created spreadsheet
 */
async function createAndUpdateSpreadsheet(config) {
  try {
    // Validate required configuration
    validateConfig(config);

    // Create authentication client (using service account)
    const auth = createAuthClient();
    const sheets = google.sheets({ version: "v4", auth });
    const drive = google.drive({ version: "v3", auth });

    console.log("Authentication successful. Creating new spreadsheet...");

    // 1. Copy the template spreadsheet
    const sourceSpreadsheetId = extractSpreadsheetId(config.templateUrl);
    const newFile = await copySpreadsheet(
      drive,
      sourceSpreadsheetId,
      config.newSpreadsheetName
    );

    // 2. Move the new spreadsheet to the specified folder in the shared drive
    await moveFileToFolder(
      drive,
      newFile.id,
      config.folderId,
      config.sharedDriveId
    );

    // 3. Update sharing permissions to be viewable by everyone in the organization
    await shareWithOrganization(drive, newFile.id, config.orgDomain);

    // 4. Process all data operations
    await processDataOperations(sheets, newFile.id, config.operations);

    // 5. Generate the URL for the new spreadsheet
    const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${newFile.id}`;
    console.log(`New spreadsheet created successfully: ${spreadsheetUrl}`);

    return spreadsheetUrl;
  } catch (error) {
    console.error("Error in createAndUpdateSpreadsheet:", error.message);
    throw error;
  }
}

/**
 * Validates the configuration object
 * @param {Object} config - The configuration object
 */
function validateConfig(config) {
  const requiredFields = [
    "templateUrl",
    "newSpreadsheetName",
    "folderId",
    "sharedDriveId",
    "orgDomain",
    "operations",
  ];

  for (const field of requiredFields) {
    if (!config[field]) {
      throw new Error(`Missing required configuration: ${field}`);
    }
  }

  if (!Array.isArray(config.operations) || config.operations.length === 0) {
    throw new Error("Operations must be a non-empty array");
  }

  // Validate each operation
  config.operations.forEach((operation, index) => {
    if (!operation.sheetName) {
      throw new Error(`Operation #${index + 1} is missing sheetName`);
    }
    if (!operation.dataPath) {
      throw new Error(`Operation #${index + 1} is missing dataPath`);
    }
    if (!operation.operationType) {
      throw new Error(`Operation #${index + 1} is missing operationType`);
    }
    if (operation.operationType === "replaceAtCell" && !operation.cellId) {
      throw new Error(
        `Operation #${index + 1} with type 'replaceAtCell' is missing cellId`
      );
    }
  });
}

/**
 * Extracts the spreadsheet ID from a Google Sheets URL
 * @param {string} url - The Google Sheets URL
 * @returns {string} - The spreadsheet ID
 */
function extractSpreadsheetId(url) {
  const idMatch = url.match(/[-\w]{25,}/);
  if (!idMatch) {
    throw new Error(
      "Invalid Google Sheets URL. Could not extract spreadsheet ID."
    );
  }
  return idMatch[0];
}

/**
 * Creates a copy of a spreadsheet
 * @param {Object} drive - Google Drive API client
 * @param {string} sourceId - Source spreadsheet ID
 * @param {string} name - Name for the new spreadsheet
 * @returns {Object} - The newly created file
 */
async function copySpreadsheet(drive, sourceId, name) {
  try {
    const response = await drive.files.copy({
      fileId: sourceId,
      requestBody: {
        name: name,
      },
      supportsAllDrives: true,
    });

    console.log(`Template spreadsheet copied with ID: ${response.data.id}`);
    return response.data;
  } catch (error) {
    console.error("Error copying spreadsheet:", error.message);
    throw error;
  }
}

/**
 * Moves a file to a specified folder in a shared drive
 * @param {Object} drive - Google Drive API client
 * @param {string} fileId - ID of the file to move
 * @param {string} folderId - ID of the destination folder
 * @param {string} sharedDriveId - ID of the shared drive
 */
async function moveFileToFolder(drive, fileId, folderId, sharedDriveId) {
  try {
    // Remove the file from its parent folders
    const file = await drive.files.get({
      fileId: fileId,
      fields: "parents",
      supportsAllDrives: true,
    });

    const previousParents = file.data.parents.join(",");

    // Move the file to the new folder
    await drive.files.update({
      fileId: fileId,
      addParents: folderId,
      removeParents: previousParents,
      supportsAllDrives: true,
      driveId: sharedDriveId,
      fields: "id, parents",
    });

    console.log(
      `File moved to folder with ID: ${folderId} in shared drive: ${sharedDriveId}`
    );
  } catch (error) {
    console.error("Error moving file to folder:", error.message);
    throw error;
  }
}

/**
 * Sets file sharing permissions to be accessible by everyone in the organization
 * @param {Object} drive - Google Drive API client
 * @param {string} fileId - ID of the file to share
 * @param {string} orgDomain - Your organization's domain (e.g., example.com)
 */
async function shareWithOrganization(drive, fileId, orgDomain) {
  try {
    await drive.permissions.create({
      fileId: fileId,
      supportsAllDrives: true,
      requestBody: {
        type: "domain",
        role: "reader",
        domain: orgDomain,
      },
    });

    console.log(`File shared with everyone in the organization (${orgDomain})`);
  } catch (error) {
    console.error("Error sharing file with organization:", error.message);
    throw error;
  }
}

/**
 * Processes all data operations on the spreadsheet
 * @param {Object} sheets - Google Sheets API client
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @param {Array} operations - Array of operations to perform
 */
async function processDataOperations(sheets, spreadsheetId, operations) {
  try {
    // Get all sheet names and IDs from the spreadsheet
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId: spreadsheetId,
    });

    const sheetsMap = {};
    spreadsheet.data.sheets.forEach((sheet) => {
      sheetsMap[sheet.properties.title] = sheet.properties.sheetId;
    });

    // Process each operation
    for (const operation of operations) {
      console.log(`Processing operation for sheet: ${operation.sheetName}`);

      // Check if the sheet exists
      if (!sheetsMap[operation.sheetName]) {
        console.warn(
          `Sheet "${operation.sheetName}" not found, skipping operation`
        );
        continue;
      }

      // Read CSV data
      const csvData = await readCsvFile(operation.dataPath);

      // Perform the operation based on its type
      switch (operation.operationType) {
        case "replaceEntireSheet":
          await replaceEntireSheet(
            sheets,
            spreadsheetId,
            operation.sheetName,
            csvData
          );
          break;
        case "replaceAtCell":
          await replaceAtCell(
            sheets,
            spreadsheetId,
            operation.sheetName,
            operation.cellId,
            csvData
          );
          break;
        default:
          console.warn(
            `Unknown operation type: ${operation.operationType}, skipping`
          );
      }
    }

    console.log("All operations completed successfully");
  } catch (error) {
    console.error("Error processing data operations:", error.message);
    throw error;
  }
}

/**
 * Reads and parses a CSV file
 * @param {string} filePath - Path to the CSV file
 * @returns {Array} - 2D array of CSV data
 */
async function readCsvFile(filePath) {
  try {
    const absolutePath = path.resolve(filePath);
    const fileContent = fs.readFileSync(absolutePath, "utf8");

    // Parse CSV content
    const records = parse(fileContent, {
      skip_empty_lines: true,
      columns: false, // Return data as arrays, not objects
    });

    console.log(`CSV file read successfully: ${filePath}`);
    return records;
  } catch (error) {
    console.error(`Error reading CSV file (${filePath}):`, error.message);
    throw error;
  }
}

/**
 * Replaces all data in a sheet with new data
 * @param {Object} sheets - Google Sheets API client
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @param {string} sheetName - Name of the sheet to update
 * @param {Array} data - 2D array of data to insert
 */
async function replaceEntireSheet(sheets, spreadsheetId, sheetName, data) {
  try {
    // Clear the existing content first
    await sheets.spreadsheets.values.clear({
      spreadsheetId: spreadsheetId,
      range: sheetName,
    });

    // Insert the new data
    await sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetId,
      range: sheetName,
      valueInputOption: "USER_ENTERED", // This preserves formulas, dates, etc.
      requestBody: {
        values: data,
      },
    });

    console.log(`Replaced all data in sheet: ${sheetName}`);
  } catch (error) {
    console.error(`Error replacing sheet data (${sheetName}):`, error.message);
    throw error;
  }
}

/**
 * Replaces data starting at a specific cell
 * @param {Object} sheets - Google Sheets API client
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @param {string} sheetName - Name of the sheet to update
 * @param {string} cellId - Starting cell ID (e.g., "A2")
 * @param {Array} data - 2D array of data to insert
 */
async function replaceAtCell(sheets, spreadsheetId, sheetName, cellId, data) {
  try {
    // Determine the range based on the starting cell and data dimensions
    const range = `${sheetName}!${cellId}:${getEndCell(cellId, data)}`;

    // Insert the data at the specified cell
    await sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetId,
      range: range,
      valueInputOption: "USER_ENTERED", // This preserves formulas, dates, etc.
      requestBody: {
        values: data,
      },
    });

    console.log(`Replaced data at cell ${cellId} in sheet: ${sheetName}`);
  } catch (error) {
    console.error(
      `Error replacing data at cell (${sheetName}, ${cellId}):`,
      error.message
    );
    throw error;
  }
}

/**
 * Calculates the end cell reference based on the starting cell and data dimensions
 * @param {string} startCell - Starting cell reference (e.g., "A2")
 * @param {Array} data - 2D array of data
 * @returns {string} - End cell reference
 */
function getEndCell(startCell, data) {
  // Extract the column letter and row number
  const colMatch = startCell.match(/[A-Z]+/);
  const rowMatch = startCell.match(/\d+/);

  if (!colMatch || !rowMatch) {
    throw new Error(`Invalid cell reference: ${startCell}`);
  }

  const startCol = colMatch[0];
  const startRow = parseInt(rowMatch[0]);

  // Calculate the end row and column
  const endRow = startRow + data.length - 1;
  const endCol = getColumnLetter(
    columnToNumber(startCol) + getMaxRowLength(data) - 1
  );

  return `${endCol}${endRow}`;
}

/**
 * Converts a column letter to its corresponding number
 * @param {string} column - Column letter (e.g., "A", "Z", "AA")
 * @returns {number} - Column number (0-based)
 */
function columnToNumber(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Converts a column number to its corresponding letter
 * @param {number} column - Column number (1-based)
 * @returns {string} - Column letter
 */
function getColumnLetter(column) {
  let temp,
    letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Gets the maximum row length in a 2D array
 * @param {Array} data - 2D array
 * @returns {number} - Maximum row length
 */
function getMaxRowLength(data) {
  let maxLength = 0;
  for (const row of data) {
    maxLength = Math.max(maxLength, row.length);
  }
  return maxLength;
}

// Export the main function
export { createAndUpdateSpreadsheet };
