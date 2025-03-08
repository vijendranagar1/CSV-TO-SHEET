// simplified-auth.js
import { google } from "googleapis";
import fs from "fs";
import path from "path";

/**
 * Creates a simple authentication client using a service account
 * This avoids the need for OAuth tokens and browser authentication
 * @returns {google.auth.JWT} - Authenticated Google API client
 */
function createAuthClient() {
  try {
    // Path to your service account credentials file (downloaded from Google Cloud Console)
    const CREDENTIALS_PATH = path.join(process.cwd(), "credentials.json");

    // Read the credentials file
    const content = fs.readFileSync(CREDENTIALS_PATH, "utf8");
    const credentials = JSON.parse(content);

    // Create a JWT client using the service account credentials
    const client = new google.auth.JWT(
      credentials.client_email,
      null,
      credentials.private_key,
      [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
      ]
    );

    console.log("Authentication client created successfully");
    return client;
  } catch (err) {
    console.error("Error creating authentication client:", err.message);
    throw new Error(
      "Failed to create authentication client. Make sure service-account.json exists."
    );
  }
}

export { createAuthClient };
