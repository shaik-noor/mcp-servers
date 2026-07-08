import { OneNoteClient } from "./onenote-client.js";
import * as dotenv from "dotenv";

// Load local .env
dotenv.config();

async function main() {
  console.log("Initializing MSAL Node application...");
  
  let client: OneNoteClient;
  try {
    client = new OneNoteClient();
  } catch (error: any) {
    console.error("\n[ERROR] Initialization failed:", error.message);
    console.log("\nPlease ensure your .env file is configured correctly in the project root.");
    process.exit(1);
  }

  const pca = client.getPca();
  const scopes = client.getScopes();

  const deviceCodeRequest = {
    scopes,
    deviceCodeCallback: (response: any) => {
      console.log("\n======================================================================");
      console.log("                  MICROSOFT ONENOTE AUTHENTICATION                    ");
      console.log("======================================================================");
      console.log(`\n${response.message}`);
      console.log("\n======================================================================");
    },
  };

  try {
    console.log("Requesting device code from Microsoft Identity platform...");
    const response = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
    console.log("\n======================================================================");
    console.log("               AUTHENTICATION COMPLETED SUCCESSFULLY!                 ");
    console.log("======================================================================");
    if (response && response.account) {
      console.log(`Account Username: ${response.account.username}`);
      console.log(`Environment:      ${response.account.environment}`);
    } else {
      console.log("Authenticated, but no account info was returned.");
    }
    console.log("\nToken has been cached successfully in 'token-cache.json'.");
    console.log("You can now start and use the MCP server.");
    console.log("======================================================================");
    process.exit(0);
  } catch (error: any) {
    console.error("\n[ERROR] Authentication failed:", error.message);
    process.exit(1);
  }
}

main();
