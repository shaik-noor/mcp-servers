# Microsoft OneNote MCP Server

This is a Model Context Protocol (MCP) server that connects Claude Desktop and Claude Code with Microsoft OneNote using the Microsoft Graph API. It allows Claude to search, read, create, update, and delete OneNote pages in real-time.

## Features

- **Microsoft OneNote Integration**:
  - `onenote_search`: Search pages across all notebooks in real-time.
  - `onenote_list_notebooks`: List all notebooks in the account.
  - `onenote_list_sections`: List sections within a notebook.
  - `onenote_list_pages`: List pages in a specific section, notebook, or globally (handles pagination for large notebooks).
  - `onenote_get_page_content`: Retrieve page content parsed to clean Markdown.
  - `onenote_create_page`: Create a new page inside a specific section.
  - `onenote_update_page`: Append Markdown content to an existing page.
  - `onenote_delete_page`: Delete a page permanently.
  - `onenote_get_auth_status`: Check Microsoft Graph login status.

---

## Prerequisites & Installation

### Prerequisites

To run and build this MCP server, the following runtimes are required/recommended:
1. **Node.js (v18+)**: **Required** for the main MCP server execution.
2. **Python 3**: **Optional** (highly recommended) for extracting and parsing local offline `.one` and `.onepkg` notebook files. If Python is not on your `PATH`, local offline file scanning will be skipped, but the online cloud functionalities via Microsoft Graph will function perfectly.

### Installation

We provide automated setup scripts that verify your environment, notify you if Node.js/npm or Python is missing with instructions on how to install them, copy the environment configuration template, and install npm dependencies.

To run the automated setup:
- **Windows**: Double-click `scripts\setup.bat` or run:
  ```cmd
  scripts\setup.bat
  ```
- **macOS / Linux**: Run:
  ```bash
  chmod +x scripts/setup.sh && ./scripts/setup.sh
  ```

*(If Node.js is already installed, you can also run `npm run setup` directly).*

Alternatively, you can perform these steps manually:

#### 1. Install dependencies
Ensure dependencies are installed:
```bash
npm install
```

#### 2. Configure environment variables
Copy `.env.example` to `.env` in the root of the project:
```bash
cp .env.example .env
```
Open `.env` and fill in the values described below.

---

## 3. Microsoft OneNote (Microsoft Graph) Setup

To connect to OneNote, you need to register a free application in the Azure Portal to act as a client.

1. Go to the [Microsoft Azure Portal - App Registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade).
2. Click **+ New registration**.
3. Fill out the form:
   - **Name**: `OneNote-MCP-Server` (or any name you prefer).
   - **Supported account types**: Select **Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.
   - **Redirect URI**: Select **Public client/native (mobile & desktop)** from the dropdown and enter:
     `https://login.microsoftonline.com/common/oauth2/nativeclient`
4. Click **Register**.
5. Copy the **Application (client) ID** from the overview page and paste it into your `.env` file:
   ```env
   ONENOTE_CLIENT_ID=your_copied_client_id_here
   ONENOTE_TENANT_ID=common
   ```
6. In the Azure portal, click **Authentication** in the left sidebar, scroll down to **Advanced settings**, and set **Allow public client flows** (Enable the next mobile and desktop flows) to **Yes**. Click **Save**.

### Authenticating OneNote
Once `.env` has your `ONENOTE_CLIENT_ID`, run the CLI login command:
```bash
npm run login
```
This will print a code and a URL:
1. Open the URL: `https://microsoft.com/devicelogin`
2. Enter the code shown in your terminal.
3. Sign in with your Microsoft account (personal, work, or school).
4. Upon successful login, the script will cache the token in `token-cache.json`.

---

## 4. Integrating with Claude

### Claude Desktop Configuration
Open your Claude Desktop config file. It is located at:
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json` (or the equivalent local packages path if using the Microsoft Store version).

Add the server to the `mcpServers` object:

```json
{
  "mcpServers": {
    "onenote-mcp": {
      "command": "node",
      "args": ["c:/noor/workspace/mcp-servers/onenote-mcp/dist/index.js"]
    }
  }
}
```
*(Make sure to use forward slashes `/` in windows file paths inside the JSON configuration).*

### Running in Development
To run in development mode or check logs:
1. Build the server:
   ```bash
   npm run build
   ```
2. Start the server on stdio:
   ```bash
   npm run start
   ```
