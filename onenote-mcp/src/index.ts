import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { OneNoteClient } from "./onenote-client.js";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import * as dotenv from "dotenv";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Load environment variables from .env file relative to this script
dotenv.config({ path: join(__dirname, "../.env") });

// Initialize clients conditionally so server doesn't crash if config is missing
let onenoteClient: OneNoteClient | null = null;

try {
  if (process.env.ONENOTE_CLIENT_ID) {
    onenoteClient = new OneNoteClient();
  } else {
    console.error("[WARN] ONENOTE_CLIENT_ID is missing. OneNote tools will not function.");
  }
} catch (error: any) {
  console.error(`[ERROR] Failed to initialize OneNote client: ${error.message}`);
}

// Create the MCP server
const server = new Server(
  {
    name: "onenote-mcp-server",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// Register tools list handler
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      // OneNote Tools
      {
        name: "onenote_search",
        description: "Search for pages in your OneNote notebooks in real-time.",
        inputSchema: {
          type: "object",
          properties: {
            query: {
              type: "string",
              description: "Search query string",
            },
          },
          required: ["query"],
        },
      },
      {
        name: "onenote_list_notebooks",
        description: "List all OneNote notebooks in your Microsoft account.",
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "onenote_list_sections",
        description: "List sections in a specific OneNote notebook or all sections.",
        inputSchema: {
          type: "object",
          properties: {
            notebookId: {
              type: "string",
              description: "Optional Notebook ID to filter sections",
            },
          },
        },
      },
      {
        name: "onenote_list_pages",
        description: "List pages in a specific section, notebook, or all pages.",
        inputSchema: {
          type: "object",
          properties: {
            sectionId: {
              type: "string",
              description: "Optional Section ID to filter pages",
            },
            notebookId: {
              type: "string",
              description: "Optional Notebook ID to filter pages",
            },
          },
        },
      },
      {
        name: "onenote_get_page_content",
        description: "Retrieve the HTML content of a OneNote page parsed into Markdown.",
        inputSchema: {
          type: "object",
          properties: {
            pageId: {
              type: "string",
              description: "The OneNote page ID",
            },
          },
          required: ["pageId"],
        },
      },
      {
        name: "onenote_create_page",
        description: "Create a new OneNote page inside a specific section.",
        inputSchema: {
          type: "object",
          properties: {
            sectionId: {
              type: "string",
              description: "The section ID where the page should be created",
            },
            title: {
              type: "string",
              description: "The title of the page",
            },
            contentMarkdown: {
              type: "string",
              description: "The page content in Markdown",
            },
          },
          required: ["sectionId", "title", "contentMarkdown"],
        },
      },
      {
        name: "onenote_update_page",
        description: "Append new content in Markdown to an existing OneNote page.",
        inputSchema: {
          type: "object",
          properties: {
            pageId: {
              type: "string",
              description: "The OneNote page ID to update",
            },
            contentMarkdown: {
              type: "string",
              description: "The Markdown content to append to the page",
            },
          },
          required: ["pageId", "contentMarkdown"],
        },
      },
      {
        name: "onenote_delete_page",
        description: "Delete a page in OneNote permanently.",
        inputSchema: {
          type: "object",
          properties: {
            pageId: {
              type: "string",
              description: "The OneNote page ID to delete",
            },
          },
          required: ["pageId"],
        },
      },
      {
        name: "onenote_get_auth_status",
        description: "Check the authentication status with Microsoft Graph for OneNote.",
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
    ],
  };
});

// Helper for returning errors
const errorResponse = (msg: string) => ({
  content: [{ type: "text", text: `Error: ${msg}` }],
  isError: true,
});

// Helper for returning text
const textResponse = (text: string) => ({
  content: [{ type: "text", text }],
});

// Register tool execution handler
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {


    // ----------------------------------------------------
    // OneNote Tools Handlers
    // ----------------------------------------------------
    if (name.startsWith("onenote_")) {
      if (!onenoteClient) {
        return errorResponse(
          "OneNote client is not initialized. Please ensure ONENOTE_CLIENT_ID is defined in environment variables."
        );
      }

      // Check auth status for all tools except checking auth status itself
      if (name !== "onenote_get_auth_status") {
        const authenticated = await onenoteClient.isAuthenticated();
        if (!authenticated) {
          return errorResponse(
            "You are not authenticated with Microsoft Graph. Please run 'npm run login' in the server directory to sign in via Device Code Flow."
          );
        }
      }

      switch (name) {
        case "onenote_get_auth_status": {
          const authenticated = await onenoteClient.isAuthenticated();
          if (authenticated) {
            return textResponse("OneNote Status: Authenticated successfully!");
          } else {
            return textResponse(
              "OneNote Status: NOT Authenticated.\n\nTo authenticate, please run the following command in the server directory:\n  npm run login\n\nThis will generate a verification code and prompt you to log in via your web browser."
            );
          }
        }

        case "onenote_search": {
          const { query } = args as { query: string };
          const results = await onenoteClient.search(query);
          return textResponse(JSON.stringify(results, null, 2));
        }

        case "onenote_list_notebooks": {
          const notebooks = await onenoteClient.listNotebooks();
          return textResponse(JSON.stringify(notebooks, null, 2));
        }

        case "onenote_list_sections": {
          const { notebookId } = args as { notebookId?: string };
          const sections = await onenoteClient.listSections(notebookId);
          return textResponse(JSON.stringify(sections, null, 2));
        }

        case "onenote_list_pages": {
          const { sectionId, notebookId } = args as {
            sectionId?: string;
            notebookId?: string;
          };
          const pages = await onenoteClient.listPages(sectionId, notebookId);
          return textResponse(JSON.stringify(pages, null, 2));
        }

        case "onenote_get_page_content": {
          const { pageId } = args as { pageId: string };
          const content = await onenoteClient.getPageContent(pageId);
          return textResponse(content);
        }

        case "onenote_create_page": {
          const { sectionId, title, contentMarkdown } = args as {
            sectionId: string;
            title: string;
            contentMarkdown: string;
          };
          const page = await onenoteClient.createPage(
            sectionId,
            title,
            contentMarkdown
          );
          return textResponse(
            `Successfully created OneNote page!\nID: ${page.id}\nTitle: ${page.title}\nURL: ${page.url}`
          );
        }

        case "onenote_update_page": {
          const { pageId, contentMarkdown } = args as {
            pageId: string;
            contentMarkdown: string;
          };
          await onenoteClient.updatePage(pageId, contentMarkdown);
          return textResponse(`Successfully appended content to OneNote page ${pageId}`);
        }

        case "onenote_delete_page": {
          const { pageId } = args as { pageId: string };
          await onenoteClient.deletePage(pageId);
          return textResponse(`Successfully deleted OneNote page ${pageId}`);
        }



        default:
          return errorResponse(`Unknown OneNote tool: ${name}`);
      }
    }

    return errorResponse(`Tool not found: ${name}`);
  } catch (error: any) {
    console.error(`Error executing tool ${name}:`, error);
    return errorResponse(error.message || String(error));
  }
});

// Run the stdio transport
async function run() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("OneNote MCP Server running on stdio");
}

run().catch((error) => {
  console.error("Fatal error running MCP Server:", error);
  process.exit(1);
});
