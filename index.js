// index.js
import { createAndUpdateSpreadsheet } from "./sheets-automation.js";
import readline from "readline";

// Create readline interface
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

// Prompt helper function
function prompt(question) {
  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      resolve(answer);
    });
  });
}

/**
 * Main application function
 */
async function main() {
  try {
    console.log("Google Sheets Automation Tool");
    console.log("============================");
    console.log("Current working directory:", process.cwd());

    // Get user input for configuration
    const templateUrl = await prompt(
      "Enter the URL of the template spreadsheet: "
    );
    const newSpreadsheetName = await prompt(
      "Enter a name for the new spreadsheet: "
    );
    const folderId = await prompt(
      "Enter the folder ID where the spreadsheet should be created: "
    );
    const sharedDriveId = await prompt("Enter the shared drive ID: ");
    const orgDomain = await prompt(
      "Enter your organization domain (e.g., example.com): "
    );

    // Get number of operations
    const numOperations = parseInt(
      await prompt("How many sheet operations do you want to perform? "),
      10
    );

    const operations = [];

    // Collect details for each operation
    for (let i = 0; i < numOperations; i++) {
      console.log(`\nOperation #${i + 1}:`);

      const sheetName = await prompt(
        "Enter the name of the sub-sheet to modify: "
      );
      const dataPath = await prompt(
        "Enter the path to the CSV file with data: "
      );

      console.log("Operation types:");
      console.log("1. Replace entire sheet");
      console.log("2. Replace data starting at a specific cell");

      const operationTypeChoice = await prompt(
        "Choose operation type (1 or 2): "
      );

      let operationType, cellId;

      if (operationTypeChoice === "1") {
        operationType = "replaceEntireSheet";
      } else if (operationTypeChoice === "2") {
        operationType = "replaceAtCell";
        cellId = await prompt("Enter the starting cell ID (e.g., A2): ");
      } else {
        console.error(
          "Invalid operation type. Defaulting to replace entire sheet."
        );
        operationType = "replaceEntireSheet";
      }

      // Add operation to the list
      operations.push({
        sheetName,
        dataPath,
        operationType,
        ...(cellId && { cellId }),
      });
    }

    // Construct the configuration object
    const config = {
      templateUrl,
      newSpreadsheetName,
      folderId,
      sharedDriveId,
      orgDomain,
      operations,
    };

    console.log("\nStarting spreadsheet creation and data operations...");

    // Execute the main function
    const spreadsheetUrl = await createAndUpdateSpreadsheet(config);

    console.log("\nProcess completed successfully!");
    console.log(`New spreadsheet URL: ${spreadsheetUrl}`);
  } catch (error) {
    console.error("An error occurred:", error.message);
    console.error("Stack trace:", error.stack);
  } finally {
    rl.close();
  }
}

// Run the application
main();
