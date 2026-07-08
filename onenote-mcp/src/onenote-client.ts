import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";
import { PublicClientApplication, Configuration } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { marked } from "marked";
import { NodeHtmlMarkdown } from "node-html-markdown";
import * as child_process from "child_process";

// Simple local file-based token cache implementation
export class FileTokenCache {
  private cachePath: string;

  constructor(cachePath: string) {
    this.cachePath = cachePath;
  }

  public getPlugin() {
    return {
      beforeCacheAccess: async (cacheContext: any) => {
        if (fs.existsSync(this.cachePath)) {
          try {
            const cacheData = fs.readFileSync(this.cachePath, "utf-8");
            cacheContext.tokenCache.deserialize(cacheData);
          } catch (error) {
            console.error("Error reading MSAL token cache file:", error);
          }
        }
      },
      afterCacheAccess: async (cacheContext: any) => {
        if (cacheContext.cacheHasChanged) {
          try {
            const cacheData = cacheContext.tokenCache.serialize();
            // Ensure parent directory exists
            const dir = path.dirname(this.cachePath);
            if (!fs.existsSync(dir)) {
              fs.mkdirSync(dir, { recursive: true });
            }
            fs.writeFileSync(this.cachePath, cacheData, {
              encoding: "utf-8",
              mode: 0o600, // Owner read/write only (ignored on Windows, enforces security on Unix)
            });
          } catch (error) {
            console.error("Error writing MSAL token cache file:", error);
          }
        }
      },
    };
  }
}
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export class OneNoteClient {
  private pca: PublicClientApplication;
  private graphClient: Client | null = null;
  private scopes = ["Notes.ReadWrite", "Notes.ReadWrite.All", "User.Read"];
  private cachePath: string;
  private localCachePath: string;

  constructor() {
    const clientId = process.env.ONENOTE_CLIENT_ID;
    if (!clientId) {
      throw new Error("ONENOTE_CLIENT_ID is not defined in environment variables");
    }

    const tenantId = process.env.ONENOTE_TENANT_ID || "common";
    this.cachePath = path.join(__dirname, "../token-cache.json");
    this.localCachePath = path.join(__dirname, "../local-cache.json");
    const tokenCache = new FileTokenCache(this.cachePath);

    const msalConfig: Configuration = {
      auth: {
        clientId: clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
      cache: {
        cachePlugin: tokenCache.getPlugin(),
      },
    };

    this.pca = new PublicClientApplication(msalConfig);
  }

  /**
   * Get the MSAL PublicClientApplication instance
   */
  getPca() {
    return this.pca;
  }

  /**
   * Get the scopes needed for OneNote
   */
  getScopes() {
    return this.scopes;
  }

  /**
   * Check if authenticated silently
   */
  async isAuthenticated(): Promise<boolean> {
    try {
      const token = await this.getAccessToken();
      return !!token;
    } catch {
      return false;
    }
  }

  /**
   * Silently get the access token
   */
  async getAccessToken(): Promise<string> {
    const accounts = await this.pca.getTokenCache().getAllAccounts();
    if (accounts.length === 0) {
      throw new Error("No accounts found. Please run the CLI auth command first: 'npm run login'");
    }

    try {
      const silentRequest = {
        account: accounts[0],
        scopes: this.scopes,
      };
      const result = await this.pca.acquireTokenSilent(silentRequest);
      if (!result || !result.accessToken) {
        throw new Error("Failed to retrieve access token silently");
      }
      return result.accessToken;
    } catch (error: any) {
      // If silent acquisition fails, propagate the error
      throw new Error(`Authentication required. Please run 'npm run login'. (Details: ${error.message})`);
    }
  }

  /**
   * Initialize and get the Microsoft Graph Client
   */
  private async getGraphClient(): Promise<Client> {
    if (this.graphClient) {
      return this.graphClient;
    }

    this.graphClient = Client.init({
      authProvider: async (done) => {
        try {
          const token = await this.getAccessToken();
          done(null, token);
        } catch (error: any) {
          done(error, null);
        }
      },
    });

    return this.graphClient;
  }

  /**
   * Search OneNote pages across all notebooks
   */
  async search(query: string) {
    try {
      const client = await this.getGraphClient();
      // Search endpoint: GET /me/onenote/pages?$search="query"
      const response = await client
        .api("/me/onenote/pages")
        .query({ $search: `"${query}"` })
        .get();

      return response.value.map((page: any) => ({
        id: page.id,
        title: page.title || "Untitled Page",
        url: page.links?.oneNoteClientUrl?.href || page.links?.oneNoteWebUrl?.href,
        lastModified: page.lastModifiedDateTime,
        sectionId: page.parentSection?.id,
        sectionName: page.parentSection?.displayName,
      }));
    } catch (err: any) {
      console.warn("Online search failed, trying local fallback:", err.message);
      const localPages = await this.loadLocalCache();
      const lowerQuery = query.toLowerCase();
      return localPages
        .filter(
          (p: any) =>
            p.title.toLowerCase().includes(lowerQuery) ||
            p.content.toLowerCase().includes(lowerQuery)
        )
        .map((p: any) => ({
          id: p.id,
          title: p.title,
          url: "",
          lastModified: p.lastModified,
          sectionId: p.sectionId,
          sectionName: p.sectionName,
        }));
    }
  }

  /**
   * List all notebooks
   */
  async listNotebooks() {
    const client = await this.getGraphClient();
    const response = await client
      .api("/me/onenote/notebooks")
      .select("id,displayName,links,lastModifiedDateTime")
      .get();

    return response.value.map((nb: any) => ({
      id: nb.id,
      name: nb.displayName,
      url: nb.links?.oneNoteClientUrl?.href || nb.links?.oneNoteWebUrl?.href,
      lastModified: nb.lastModifiedDateTime,
    }));
  }

  /**
   * List sections in a specific notebook or all sections
   */
  async listSections(notebookId?: string) {
    try {
      const client = await this.getGraphClient();
      const url = notebookId
        ? `/me/onenote/notebooks/${notebookId}/sections`
        : "/me/onenote/sections";
      
      const response = await client
        .api(url)
        .select("id,displayName,lastModifiedDateTime,parentNotebook")
        .get();

      return response.value.map((sec: any) => ({
        id: sec.id,
        name: sec.displayName,
        lastModified: sec.lastModifiedDateTime,
        notebookId: sec.parentNotebook?.id,
        notebookName: sec.parentNotebook?.displayName,
      }));
    } catch (err: any) {
      console.warn("Online listSections failed, trying local fallback:", err.message);
      const localPages = await this.loadLocalCache();
      const sectionsMap = new Map();
      for (const p of localPages) {
        if (notebookId && p.notebookId !== notebookId) continue;
        sectionsMap.set(p.sectionId, {
          id: p.sectionId,
          name: p.sectionName,
          lastModified: p.lastModified,
          notebookId: p.notebookId,
          notebookName: p.notebookName,
        });
      }
      return Array.from(sectionsMap.values());
    }
  }

  async listPages(sectionId?: string, notebookId?: string) {
    try {
      const client = await this.getGraphClient();
      let url = "/me/onenote/pages";
      if (sectionId) {
        url = `/me/onenote/sections/${sectionId}/pages`;
      } else if (notebookId) {
        url = `/me/onenote/notebooks/${notebookId}/pages`;
      }

      let allPages: any[] = [];
      let response = await client
        .api(url)
        .select("id,title,links,lastModifiedDateTime,parentSection")
        .top(100)
        .get();

      if (response && response.value) {
        allPages = allPages.concat(response.value);
      }

      let nextLink = response["@odata.nextLink"];
      while (nextLink) {
        try {
          response = await client.api(nextLink).get();
          if (response && response.value) {
            allPages = allPages.concat(response.value);
          }
          nextLink = response["@odata.nextLink"];
        } catch (error) {
          console.error("Error fetching next page of results:", error);
          nextLink = null;
        }
      }

      return allPages.map((page: any) => ({
        id: page.id,
        title: page.title || "Untitled Page",
        url: page.links?.oneNoteClientUrl?.href || page.links?.oneNoteWebUrl?.href,
        lastModified: page.lastModifiedDateTime,
        sectionId: page.parentSection?.id,
        sectionName: page.parentSection?.displayName,
      }));
    } catch (err: any) {
      console.warn("Online listPages failed, trying local fallback:", err.message);
      const localPages = await this.loadLocalCache();
      return localPages
        .filter((p: any) => {
          if (sectionId && p.sectionId !== sectionId) return false;
          if (notebookId && p.notebookId !== notebookId) return false;
          return true;
        })
        .map((p: any) => ({
          id: p.id,
          title: p.title,
          url: "",
          lastModified: p.lastModified,
          sectionId: p.sectionId,
          sectionName: p.sectionName,
        }));
    }
  }

  async getPageContent(pageId: string): Promise<string> {
    if (pageId.startsWith("local_")) {
      const localPages = await this.loadLocalCache();
      const page = localPages.find((p: any) => p.id === pageId);
      if (!page) {
        throw new Error(`Local page ${pageId} not found in cache.`);
      }
      return this.htmlToMarkdown(page.content);
    }

    const client = await this.getGraphClient();
    // Fetch HTML page content
    const response = await client
      .api(`/me/onenote/pages/${pageId}/content`)
      .responseType("text" as any)
      .get();

    return this.htmlToMarkdown(response);
  }

  /**
   * Create a OneNote page inside a specific section
   */
  async createPage(sectionId: string, title: string, contentMarkdown: string) {
    const client = await this.getGraphClient();
    const htmlBody = this.markdownToHtml(title, contentMarkdown);

    const response = await client
      .api(`/me/onenote/sections/${sectionId}/pages`)
      .headers({ "Content-Type": "text/html" })
      .post(htmlBody);

    return {
      id: response.id,
      title: response.title,
      url: response.links?.oneNoteClientUrl?.href || response.links?.oneNoteWebUrl?.href,
    };
  }

  /**
   * Update page content (Appends new content to the body)
   */
  async updatePage(pageId: string, contentMarkdown: string) {
    const client = await this.getGraphClient();
    // Convert markdown to clean HTML tags to append
    const htmlSnippet = this.markdownToHtmlSnippet(contentMarkdown);

    // PATCH requires commands array
    const patchCommand = [
      {
        target: "body",
        action: "append",
        content: htmlSnippet,
      },
    ];

    await client
      .api(`/me/onenote/pages/${pageId}/content`)
      .patch(patchCommand);

    return { success: true };
  }

  /**
   * Delete a OneNote page
   */
  async deletePage(pageId: string) {
    const client = await this.getGraphClient();
    await client.api(`/me/onenote/pages/${pageId}`).delete();
    return { success: true };
  }

  /**
   * Clean HTML helper to convert standard OneNote output to clean readable Markdown
   */
  private htmlToMarkdown(html: string): string {
    if (!html) return "";

    // Extract Title if present
    const titleMatch = html.match(/<title>(.*?)<\/title>/i);
    const title = titleMatch ? titleMatch[1] : "";

    // Extract body content
    const bodyMatch = html.match(/<body>([\s\S]*?)<\/body>/i);
    const bodyHtml = bodyMatch ? bodyMatch[1] : html;

    // Convert HTML to Markdown using node-html-markdown
    const markdown = NodeHtmlMarkdown.translate(bodyHtml).trim();

    return title ? `# ${title}\n\n${markdown}` : markdown;
  }

  /**
   * Helper to convert Title and Markdown content to a full HTML document
   */
  private markdownToHtml(title: string, markdown: string): string {
    const bodyHtml = this.markdownToHtmlSnippet(markdown);
    return `<!DOCTYPE html>
<html>
  <head>
    <title>${title}</title>
    <meta name="created" content="${new Date().toISOString()}" />
  </head>
  <body>
    ${bodyHtml}
  </body>
</html>`;
  }

  /**
   * Helper to convert Markdown lines into HTML tags for insertion/PATCH
   */
  private markdownToHtmlSnippet(markdown: string): string {
    // Convert Markdown to HTML using marked. Since we do not pass async configuration,
    // marked.parse returns a string. We cast it as a string.
    return marked.parse(markdown) as string;
  }

  private async loadLocalCache(force = false): Promise<any[]> {
    const cacheExists = fs.existsSync(this.localCachePath);
    let cacheData: any = null;

    if (cacheExists && !force) {
      try {
        const stats = fs.statSync(this.localCachePath);
        const ageInMs = Date.now() - stats.mtime.getTime();
        // Cache is valid for 5 minutes (300,000 ms)
        if (ageInMs < 300000) {
          const raw = fs.readFileSync(this.localCachePath, "utf-8");
          cacheData = JSON.parse(raw);
          if (cacheData && cacheData.pages) {
            return cacheData.pages;
          }
        }
      } catch (err: any) {
        console.error("Error reading local cache:", err.message);
      }
    }

    // Refresh cache
    const pages = await this.scanLocalFiles();
    try {
      fs.writeFileSync(this.localCachePath, JSON.stringify({ timestamp: Date.now(), pages }, null, 2), "utf-8");
    } catch (err: any) {
      console.error("Error writing local cache:", err.message);
    }
    return pages;
  }

  private async scanLocalFiles(): Promise<any[]> {
    const userProfile = process.env.USERPROFILE || "C:\\Users\\shmohamm";
    const scanPaths = [
      path.join(userProfile, "Documents"),
      path.join(userProfile, "OneDrive - Informatica\\Documents\\OneNote Notebooks"),
      path.join(userProfile, "Downloads")
    ];

    const allPages: any[] = [];
    
    // Search both dist and src directory for python script
    let scriptPath = path.join(__dirname, "extract_onenote.py");
    if (!fs.existsSync(scriptPath)) {
      scriptPath = path.join(__dirname, "../src/extract_onenote.py");
    }

    for (const scanPath of scanPaths) {
      if (!fs.existsSync(scanPath)) continue;
      try {
        const files = fs.readdirSync(scanPath);
        for (const file of files) {
          const ext = path.extname(file).toLowerCase();
          if (ext === ".one" || ext === ".onepkg") {
            const filePath = path.join(scanPath, file);
            try {
              const res = child_process.spawnSync("python", [scriptPath, filePath], { encoding: "utf-8" });
              if (res.status === 0 && res.stdout) {
                const extractedPages = JSON.parse(res.stdout.trim());
                for (const p of extractedPages) {
                  const title = p.title || "Untitled Page";
                  const cleanTitle = title.replace(/[^a-zA-Z0-9]/g, "_");
                  const cleanFile = file.replace(/[^a-zA-Z0-9]/g, "_");
                  const localId = `local_${cleanFile}_${cleanTitle}`;
                  allPages.push({
                    id: localId,
                    title: title,
                    content: p.content || "",
                    lastModified: new Date().toISOString(),
                    notebookId: `local_nb_${cleanFile}`,
                    notebookName: file,
                    sectionId: `local_sec_${cleanFile}`,
                    sectionName: file.replace(ext, "")
                  });
                }
              }
            } catch (err: any) {
              console.error(`Failed to scan file ${filePath}:`, err.message);
            }
          }
        }
      } catch (err: any) {
        console.error(`Failed to read directory ${scanPath}:`, err.message);
      }
    }
    return allPages;
  }
}
