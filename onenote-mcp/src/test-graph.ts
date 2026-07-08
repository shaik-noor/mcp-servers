import { OneNoteClient } from "./onenote-client.js";
import * as dotenv from "dotenv";
import * as path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.join(__dirname, "../.env") });

async function runTest() {
  console.log("Initializing OneNoteClient...");
  const client = new OneNoteClient();

  console.log("\n--- Testing listPages (should fallback to local cache on API failure) ---");
  try {
    const pages = await client.listPages();
    console.log(`Successfully retrieved ${pages.length} pages.`);
    console.log("Sample pages (top 5):");
    pages.slice(0, 5).forEach((p: any, idx: number) => {
      console.log(`  ${idx + 1}. [${p.sectionName}] ${p.title} (ID: ${p.id})`);
    });
  } catch (err: any) {
    console.error("listPages failed:", err.message);
  }

  console.log("\n--- Testing search('templates') ---");
  try {
    const results = await client.search("templates");
    console.log(`Found ${results.length} search results.`);
    results.forEach((r: any, idx: number) => {
      console.log(`  ${idx + 1}. [${r.sectionName}] ${r.title} (ID: ${r.id})`);
    });
  } catch (err: any) {
    console.error("search failed:", err.message);
  }

  console.log("\n--- Testing getPageContent (local ID) ---");
  try {
    const pages = await client.listPages();
    const localPage = pages.find((p: any) => p.id.startsWith("local_"));
    if (localPage) {
      console.log(`Fetching content for local page: ${localPage.title} (ID: ${localPage.id})`);
      const content = await client.getPageContent(localPage.id);
      console.log("Extracted Content Snippet:");
      console.log(content.slice(0, 300) + "\n...");
    } else {
      console.log("No local pages found in listPages results.");
    }
  } catch (err: any) {
    console.error("getPageContent failed:", err.message);
  }
}

runTest().catch(console.error);
