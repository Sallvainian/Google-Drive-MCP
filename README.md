# Ultimate Google Docs, Sheets, Slides & Drive MCP Server

![Demo Animation](assets/google.docs.mcp.1.gif)

Connect Claude Desktop (or other MCP clients) to your Google Docs, Google Sheets, Google Slides, and Google Drive!

> 🔥 **Check out [15 powerful tasks](SAMPLE_TASKS.md) you can accomplish with this enhanced server!**
> 📁 **NEW:** Complete Google Drive file management capabilities!
> 📊 **NEW:** Full Google Sheets support for reading, writing, and managing spreadsheets!
> 🎨 **NEW:** Full Google Slides support for creating and editing presentations!

This comprehensive server uses the Model Context Protocol (MCP) and the `fastmcp` library to provide tools for reading, writing, formatting, structuring Google Documents, Spreadsheets, and Presentations, and managing your entire Google Drive. It acts as a powerful bridge, allowing AI assistants like Claude to interact with your documents, spreadsheets, presentations, and files programmatically with advanced capabilities.

**Features:**

### Document Access & Editing

- **Read Documents:** Read content with `readGoogleDoc` (plain text, JSON structure, or markdown)
- **Append to Documents:** Add text to documents with `appendToGoogleDoc`
- **Insert Text:** Place text at specific positions with `insertText`
- **Delete Content:** Remove content from a document with `deleteRange`
- **Tab Support:** Work with multi-tab documents using `listDocumentTabs` and optional `tabId` parameter in read/write operations

### Formatting & Styling

- **Text Formatting:** Apply rich styling with `applyTextStyle` (bold, italic, colors, etc.)
- **Paragraph Formatting:** Control paragraph layout with `applyParagraphStyle` (alignment, spacing, etc.)
- **Find & Format:** Format by text content using `formatMatchingText` (legacy support)

### Document Structure

- **Tables:** Create tables with `insertTable`
- **Page Breaks:** Insert page breaks with `insertPageBreak`
- **Images:** Insert images from URLs with `insertImageFromUrl`, or upload local images with `insertLocalImage`
- **Experimental Features:** Tools like `fixListFormatting` for automatic list detection

### 🆕 Comment Management

- **List Comments:** View all comments in a document with `listComments` (shows author, date, and quoted text)
- **Get Comment Details:** Get specific comment with replies using `getComment`
- **Add Comments:** Create new comments anchored to text with `addComment`
- **Reply to Comments:** Add replies to existing comments with `replyToComment`
- **Resolve Comments:** Mark comments as resolved with `resolveComment`
- **Delete Comments:** Remove comments from documents with `deleteComment`

### 🆕 Google Sheets Support

- **Read Spreadsheets:** Read data from ranges with `readSpreadsheet` (supports A1 notation like "A1:B10" or "Sheet1!A1:B10")
- **Write Data:** Write data to ranges with `writeSpreadsheet` (overwrites existing data)
- **Append Rows:** Add new rows to sheets with `appendSpreadsheetRows`
- **Clear Ranges:** Clear cell values with `clearSpreadsheetRange`
- **Spreadsheet Info:** Get detailed metadata and sheet list with `getSpreadsheetInfo`
- **Create Spreadsheets:** Create new spreadsheets with `createSpreadsheet` (optionally with initial data)
- **Add Sheets:** Add new sheets/tabs to spreadsheets with `addSpreadsheetSheet`
- **List Spreadsheets:** Find and list spreadsheets with `listGoogleSheets`

### 🆕 Google Drive File Management

- **Document Discovery:** Find and list documents with `listGoogleDocs`, `searchGoogleDocs`, `getRecentGoogleDocs`
- **Document Information:** Get detailed metadata with `getDocumentInfo`
- **Folder Management:** Create folders (`createFolder`), list contents (`listFolderContents`), get info (`getFolderInfo`)
- **File Operations:** Move (`moveFile`), copy (`copyFile`), rename (`renameFile`), delete (`deleteFile`)
- **Document Creation:** Create new docs (`createDocument`) or from templates (`createFromTemplate`)

### 🆕 Google Slides Support

- **Read Presentations:** Get presentation metadata and content with `getPresentation`, list slides with `listSlides`, get detailed slide content with `getSlide`
- **Create Presentations:** Create new presentations with `createPresentation`
- **Slide Management:** Add slides with `addSlide`, duplicate with `duplicateSlide`, delete with `deleteSlide`, reorder with `moveSlide`
- **Elements:** Add text boxes (`addTextBox`), shapes (`addShape`), images (`addImage`), tables (`addTable`)
- **Content Editing:** Insert or replace text in elements with `insertTextInElement`, update speaker notes with `updateSpeakerNotes`
- **Element Management:** Delete elements with `deleteElement`

### Integration

- **Google Authentication:** Secure OAuth 2.0 authentication with full Drive, Docs, Sheets, and Slides access
- **MCP Compliant:** Designed for use with Claude and other MCP clients
- **VS Code Integration:** [Setup guide](vscode.md) for VS Code MCP extension

---

## Prerequisites

Before you start, make sure you have:

1.  **Node.js and npm:** A recent version of Node.js (which includes npm) installed on your computer. You can download it from [nodejs.org](https://nodejs.org/). (Version 18 or higher recommended).
2.  **Git:** Required for cloning this repository. ([Download Git](https://git-scm.com/downloads)).
3.  **A Google Account:** The account that owns or has access to the Google Docs you want to interact with.
4.  **Command Line Familiarity:** Basic comfort using a terminal or command prompt (like Terminal on macOS/Linux, or Command Prompt/PowerShell on Windows).
5.  **Claude Desktop (Optional):** If your goal is to connect this server to Claude, you'll need the Claude Desktop application installed.

---

## Setup Instructions

Follow these steps carefully to get your own instance of the server running.

### Step 1: Google Cloud Project & Credentials (The Important Bit!)

This server needs permission to talk to Google APIs on your behalf. You'll create special "keys" (credentials) that only your server will use.

1.  **Go to Google Cloud Console:** Open your web browser and go to the [Google Cloud Console](https://console.cloud.google.com/). You might need to log in with your Google Account.
2.  **Create or Select a Project:**
    - If you don't have a project, click the project dropdown near the top and select "NEW PROJECT". Give it a name (e.g., "My MCP Docs Server") and click "CREATE".
    - If you have existing projects, you can select one or create a new one.
3.  **Enable APIs:** You need to turn on the specific Google services this server uses.
    - In the search bar at the top, type "APIs & Services" and select "Library".
    - Search for "**Google Docs API**" and click on it. Then click the "**ENABLE**" button.
    - Search for "**Google Sheets API**" and click on it. Then click the "**ENABLE**" button.
    - Search for "**Google Slides API**" and click on it. Then click the "**ENABLE**" button.
    - Search for "**Google Drive API**" and click on it. Then click the "**ENABLE**" button (this is often needed for finding files or permissions).
4.  **Configure OAuth Consent Screen:** This screen tells users (usually just you) what your app wants permission for.
    - On the left menu, click "APIs & Services" -> "**OAuth consent screen**".
    - Choose User Type: Select "**External**" and click "CREATE".
    - Fill in App Information:
      - **App name:** Give it a name users will see (e.g., "Claude Docs MCP Access").
      - **User support email:** Select your email address.
      - **Developer contact information:** Enter your email address.
    - Click "**SAVE AND CONTINUE**".
    - **Scopes:** Click "**ADD OR REMOVE SCOPES**". Search for and add the following scopes:
      - `https://www.googleapis.com/auth/documents` (Allows reading/writing docs)
      - `https://www.googleapis.com/auth/spreadsheets` (Allows reading/writing spreadsheets)
      - `https://www.googleapis.com/auth/presentations` (Allows reading/writing presentations)
      - `https://www.googleapis.com/auth/drive.file` (Allows access to specific files opened/created by the app)
      - Click "**UPDATE**".
    - Click "**SAVE AND CONTINUE**".
    - **Test Users:** Click "**ADD USERS**". Enter the same Google email address you are logged in with. Click "**ADD**". This allows _you_ to use the app while it's in "testing" mode.
    - Click "**SAVE AND CONTINUE**". Review the summary and click "**BACK TO DASHBOARD**".
5.  **Create Credentials (The Keys!):**
    - On the left menu, click "APIs & Services" -> "**Credentials**".
    - Click "**+ CREATE CREDENTIALS**" at the top and choose "**OAuth client ID**".
    - **Application type:** Select "**Desktop app**" from the dropdown.
    - **Name:** Give it a name (e.g., "MCP Docs Desktop Client").
    - Click "**CREATE**".
6.  **⬇️ DOWNLOAD THE CREDENTIALS FILE:** A box will pop up showing your Client ID. Click the "**DOWNLOAD JSON**" button.
    - Save this file. It will likely be named something like `client_secret_....json`.
    - **IMPORTANT:** Rename the downloaded file to exactly `credentials.json`.
7.  ⚠️ **SECURITY WARNING:** Treat this `credentials.json` file like a password! Do not share it publicly, and **never commit it to GitHub.** Anyone with this file could potentially pretend to be _your application_ (though they'd still need user consent to access data).

### Step 2: Get the Server Code

1.  **Clone the Repository:** Open your terminal/command prompt and run:
    ```bash
    git clone https://github.com/a-bonus/google-docs-mcp.git mcp-googledocs-server
    ```
2.  **Navigate into Directory:**
    ```bash
    cd mcp-googledocs-server
    ```
3.  **Place Credentials:** Move or copy the `credentials.json` file you downloaded and renamed (from Step 1.6) directly into this `mcp-googledocs-server` folder.

### Step 3: Install Dependencies

Your server needs some helper libraries specified in the `package.json` file.

1.  In your terminal (make sure you are inside the `mcp-googledocs-server` directory), run:
    ```bash
    npm install
    ```
    This will download and install all the necessary packages into a `node_modules` folder.

### Step 4: Build the Server Code

The server is written in TypeScript (`.ts`), but we need to compile it into JavaScript (`.js`) that Node.js can run directly.

1.  In your terminal, run:
    ```bash
    npm run build
    ```
    This uses the TypeScript compiler (`tsc`) to create a `dist` folder containing the compiled JavaScript files.

### Step 5: First Run & Google Authorization (One Time Only)

Now you need to run the server once manually to grant it permission to access your Google account data. This will create a `token.json` file that saves your permission grant.

1.  In your terminal, run the _compiled_ server using `node`:
    ```bash
    node ./dist/server.js
    ```
2.  **Watch the Terminal:** The script will print:
    - Status messages (like "Attempting to authorize...").
    - An "Authorize this app by visiting this url:" message followed by a long `https://accounts.google.com/...` URL.
3.  **Authorize in Browser:**
    - Copy the entire long URL from the terminal.
    - Paste the URL into your web browser and press Enter.
    - Log in with the **same Google account** you added as a Test User in Step 1.4.
      - Google will show a screen asking for permission for your app ("Claude Docs MCP Access" or similar) to access Google Docs, Sheets, and Drive. Review and click "**Allow**" or "**Grant**".
4.  **Get the Authorization Code:**
    - After clicking Allow, your browser will likely try to redirect to `http://localhost` and show a **"This site can't be reached" error**. **THIS IS NORMAL!**
    - Look **carefully** at the URL in your browser's address bar. It will look like `http://localhost/?code=4/0Axxxxxxxxxxxxxx&scope=...`
    - Copy the long string of characters **between `code=` and the `&scope` part**. This is your single-use authorization code.
5.  **Paste Code into Terminal:** Go back to your terminal where the script is waiting ("Enter the code from that page here:"). Paste the code you just copied.
6.  **Press Enter.**
7.  **Success!** The script should print:
    - "Authentication successful!"
    - "Token stored to .../token.json"
    - It will then finish starting and likely print "Awaiting MCP client connection via stdio..." or similar, and then exit (or you can press `Ctrl+C` to stop it).
8.  ✅ **Check:** You should now see a new file named `token.json` in your `mcp-googledocs-server` folder.
9.  ⚠️ **SECURITY WARNING:** This `token.json` file contains the key that allows the server to access your Google account _without_ asking again. Protect it like a password. **Do not commit it to GitHub.** The included `.gitignore` file should prevent this automatically.

### Alternative: Service Account with Domain-Wide Delegation (Enterprise)

For Google Workspace organizations that need to access documents across the domain without individual user OAuth flows, you can use a service account with domain-wide delegation.

**Prerequisites:**
- Google Workspace admin access to configure domain-wide delegation
- A service account with domain-wide delegation enabled

**Setup:**

1. **Create a Service Account** in Google Cloud Console:
   - Go to "APIs & Services" → "Credentials" → "Create Credentials" → "Service Account"
   - Download the JSON key file

2. **Enable Domain-Wide Delegation** in Google Workspace Admin Console:
   - Go to Security → API Controls → Domain-wide delegation
   - Add the service account's client ID with the required scopes:
     - `https://www.googleapis.com/auth/documents`
     - `https://www.googleapis.com/auth/drive`
     - `https://www.googleapis.com/auth/spreadsheets`

3. **Configure Environment Variables:**
   ```bash
   # Path to your service account key file
   export SERVICE_ACCOUNT_PATH="/path/to/service-account-key.json"

   # Email of the user to impersonate (required for domain-wide delegation)
   export GOOGLE_IMPERSONATE_USER="user@yourdomain.com"
   ```

4. **Update Claude Desktop Config** (add environment variables):
   ```json
   {
     "mcpServers": {
       "google-docs-mcp": {
         "command": "node",
         "args": ["/PATH/TO/mcp-googledocs-server/dist/server.js"],
         "env": {
           "SERVICE_ACCOUNT_PATH": "/path/to/service-account-key.json",
           "GOOGLE_IMPERSONATE_USER": "user@yourdomain.com"
         }
       }
     }
   }
   ```

When `GOOGLE_IMPERSONATE_USER` is set, the server will impersonate that user when accessing Google APIs, allowing access to that user's documents and Drive.

### Step 6: Configure Claude Desktop (Optional)

If you want to use this server with Claude Desktop, you need to tell Claude how to run it.

1.  **Find Your Absolute Path:** You need the full path to the server code.
    - In your terminal, make sure you are still inside the `mcp-googledocs-server` directory.
    - Run the `pwd` command (on macOS/Linux) or `cd` (on Windows, just displays the path).
    - Copy the full path (e.g., `/Users/yourname/projects/mcp-googledocs-server` or `C:\Users\yourname\projects\mcp-googledocs-server`).
2.  **Locate `mcp_config.json`:** Find Claude's configuration file:
    - **macOS:** `~/Library/Application Support/Claude/mcp_config.json` (You might need to use Finder's "Go" -> "Go to Folder..." menu and paste `~/Library/Application Support/Claude/`)
    - **Windows:** `%APPDATA%\Claude\mcp_config.json` (Paste `%APPDATA%\Claude` into File Explorer's address bar)
    - **Linux:** `~/.config/Claude/mcp_config.json`
    - _If the `Claude` folder or `mcp_config.json` file doesn't exist, create them._
3.  **Edit `mcp_config.json`:** Open the file in a text editor. Add or modify the `mcpServers` section like this, **replacing `/PATH/TO/YOUR/CLONED/REPO` with the actual absolute path you copied in Step 6.1**:

    ```json
    {
      "mcpServers": {
        "google-docs-mcp": {
          "command": "node",
          "args": [
            "/PATH/TO/YOUR/CLONED/REPO/mcp-googledocs-server/dist/server.js"
          ],
          "env": {}
        }
        // Add commas here if you have other servers defined
      }
      // Other Claude settings might be here
    }
    ```

    - **Make sure the path in `"args"` is correct and absolute!**
    - If the file already existed, carefully merge this entry into the existing `mcpServers` object. Ensure the JSON is valid (check commas!).

4.  **Save `mcp_config.json`.**
5.  **Restart Claude Desktop:** Close Claude completely and reopen it.

---

## Usage with Claude Desktop

Once configured, you should be able to use the tools in your chats with Claude:

- "Use the `google-docs-mcp` server to read the document with ID `YOUR_GOOGLE_DOC_ID`."
- "Can you get the content of Google Doc `YOUR_GOOGLE_DOC_ID`?"
- "Append 'This was added by Claude!' to document `YOUR_GOOGLE_DOC_ID` using the `google-docs-mcp` tool."

### Working with Tabs

Google Docs now supports multi-tab documents. This MCP server provides full support for working with tabs:

**Listing Tabs:**

- "List all tabs in document `YOUR_GOOGLE_DOC_ID` using the `listDocumentTabs` tool."
- "Show me the tab structure with content summary for document `YOUR_GOOGLE_DOC_ID`."

**Reading from Specific Tabs:**

- "Read the content from tab `TAB_ID` in document `YOUR_GOOGLE_DOC_ID` using the `readGoogleDoc` tool."
- "Get the markdown content from tab `TAB_ID` in document `YOUR_GOOGLE_DOC_ID`."

**Writing to Specific Tabs:**

- "Append 'New content' to tab `TAB_ID` in document `YOUR_GOOGLE_DOC_ID`."
- "Insert text at index 100 in tab `TAB_ID` of document `YOUR_GOOGLE_DOC_ID`."
- "Delete content from range 50-100 in tab `TAB_ID` of document `YOUR_GOOGLE_DOC_ID`."

**Note:** The following tools support the optional `tabId` parameter:

- `readGoogleDoc` - Read from a specific tab
- `appendToGoogleDoc` - Append to a specific tab
- `insertText` - Insert text into a specific tab
- `deleteRange` - Delete content from a specific tab

When `tabId` is not specified, operations target the first tab (or the legacy document body for older documents without tabs).

### Advanced Usage Examples:

**Google Docs:**
- **Text Styling**: "Use `applyTextStyle` to make the text 'Important Section' bold and red (#FF0000) in document `YOUR_GOOGLE_DOC_ID`."
- **Paragraph Styling**: "Use `applyParagraphStyle` to center-align the paragraph containing 'Title Here' in document `YOUR_GOOGLE_DOC_ID`."
- **Table Creation**: "Insert a 3x4 table at index 500 in document `YOUR_GOOGLE_DOC_ID` using the `insertTable` tool."
- **Image Insertion**: "Use `insertImageFromUrl` to insert an image from 'https://example.com/image.png' at index 100 in document `YOUR_GOOGLE_DOC_ID`."
- **Local Image Upload**: "Use `insertLocalImage` to upload '/path/to/image.jpg' and insert it at index 200 in document `YOUR_GOOGLE_DOC_ID`."
- **Legacy Formatting**: "Use `formatMatchingText` to find the second instance of 'Project Alpha' and make it blue (#0000FF) in doc `YOUR_GOOGLE_DOC_ID`."

**Google Sheets:**
- **Read Data**: "Read range A1:B10 from spreadsheet `YOUR_SPREADSHEET_ID` using `readSpreadsheet`."
- **Write Data**: "Write data [[1, 2], [3, 4]] to range A1:B2 in spreadsheet `YOUR_SPREADSHEET_ID`."
- **Append Rows**: "Append rows [[5, 6], [7, 8]] to spreadsheet `YOUR_SPREADSHEET_ID` starting at A1."
- **Create Spreadsheet**: "Create a new spreadsheet titled 'Sales Data' with initial data [[Name, Amount], [Product A, 100]]."
- **Get Info**: "Get information about spreadsheet `YOUR_SPREADSHEET_ID` including all sheets."
- **Add Sheet**: "Add a new sheet named 'Summary' to spreadsheet `YOUR_SPREADSHEET_ID`."
- **Clear Range**: "Clear the range A1:B10 in spreadsheet `YOUR_SPREADSHEET_ID`."
- **List Spreadsheets**: "List all my Google Spreadsheets modified in the last 30 days."

Remember to replace:
- `YOUR_GOOGLE_DOC_ID` with the actual ID from a Google Doc's URL (the long string between `/d/` and `/edit`)
- `YOUR_SPREADSHEET_ID` with the actual ID from a Google Sheet's URL (the long string between `/d/` and `/edit`)

Claude will automatically launch your server in the background when needed using the command you provided. You do **not** need to run `node ./dist/server.js` manually anymore.

---

## Image Insertion

This server provides two ways to insert images into Google Documents:

### 1. Insert from Public URL (`insertImageFromUrl`)

Inserts an image directly from a publicly accessible URL. The image URL must be accessible without authentication.

**Parameters:**

- `documentId`: The Google Document ID
- `imageUrl`: Publicly accessible URL (http:// or https://)
- `index`: Position in the document (1-based indexing)
- `width` (optional): Image width in points
- `height` (optional): Image height in points

**Example:**

```
"Insert an image from https://example.com/logo.png at index 100 in document YOUR_DOC_ID"
```

### 2. Upload Local Image (`insertLocalImage`)

Uploads a local image file to Google Drive and inserts it into the document. This is a two-step process that:

1. Uploads the image to Google Drive (by default to the same folder as the document)
2. Makes the image publicly readable
3. Inserts the image into the document using its Drive URL

**Parameters:**

- `documentId`: The Google Document ID
- `localImagePath`: Absolute path to the local image file
- `index`: Position in the document (1-based indexing)
- `width` (optional): Image width in points
- `height` (optional): Image height in points
- `uploadToSameFolder` (optional, default: true): If true, uploads to the document's folder; if false, uploads to Drive root

**Supported formats:** .jpg, .jpeg, .png, .gif, .bmp, .webp, .svg

**Example:**

```
"Upload and insert the image at /Users/myname/Pictures/chart.png at index 200 in document YOUR_DOC_ID with width 400 and height 300"
```

**Note:** The uploaded image will be made publicly readable so it can be displayed in the document. The image file will remain in your Google Drive and can be managed separately.

---

## Security & Token Storage

- **`.gitignore`:** This repository includes a `.gitignore` file which should prevent you from accidentally committing your sensitive `credentials.json` and `token.json` files. **Do not remove these lines from `.gitignore`**.
- **Token Storage:** This server stores the Google authorization token (`token.json`) directly in the project folder for simplicity during setup. In production or more security-sensitive environments, consider storing this token more securely, such as using system keychains, encrypted files, or dedicated secret management services.

---

## Testing

The multi-tab support features have been thoroughly tested and verified:

✅ **Tested Features:**

- `listDocumentTabs` - Lists all tabs with IDs, titles, positions, and content summaries
- `readGoogleDoc` with `tabId` - Reads specific tabs; backward compatible without `tabId`
- `appendToGoogleDoc` with `tabId` - Appends to specific tabs without affecting others
- `insertText` with `tabId` - Inserts text at specific positions in specific tabs
- `deleteRange` with `tabId` - Deletes content from specific tabs in isolation
- Multi-tab operations - Sequential operations on different tabs work independently
- Error handling - Invalid tab IDs return clear, helpful error messages
- Backward compatibility - Operations without `tabId` default to first tab (legacy documents supported)

All tab-related features have been validated with real Google Docs containing multiple tabs, confirming:

- Tab isolation (operations on one tab don't affect others)
- Proper tab ID validation and error messages
- Correct content retrieval and manipulation per tab
- Full backward compatibility with single-tab and legacy documents

## Google Sheets Usage

### A1 Notation

Google Sheets uses A1 notation to specify ranges. Examples:
- `A1` - Single cell
- `A1:B10` - Range from A1 to B10
- `Sheet1!A1:B10` - Range on a specific sheet named "Sheet1"
- `A:A` - Entire column A
- `1:1` - Entire row 1

### Value Input Options

When writing data to spreadsheets, you can choose how values are interpreted:
- **USER_ENTERED** (default): Values are parsed as if typed by a user (formulas work, dates are recognized, etc.)
- **RAW**: Values are stored exactly as provided (no parsing)

### Example Workflow

```bash
# 1. Create a new spreadsheet
"Create a spreadsheet titled 'Monthly Report'"

# 2. Write headers
"Write [[Date, Sales, Expenses]] to range A1:C1 in spreadsheet YOUR_SPREADSHEET_ID"

# 3. Append data rows
"Append rows [[2024-01-01, 1000, 500], [2024-01-02, 1200, 600]] to spreadsheet YOUR_SPREADSHEET_ID starting at A2"

# 4. Read the data back
"Read range A1:C3 from spreadsheet YOUR_SPREADSHEET_ID"

# 5. Add a new sheet for analysis
"Add a new sheet named 'Analysis' to spreadsheet YOUR_SPREADSHEET_ID"
```

## Known Limitations

While this MCP server provides comprehensive Google Docs, Sheets, and Drive functionality, there are some limitations imposed by the Google APIs themselves:

### Comment Anchoring

**Programmatically Created Comments Are Not Anchored**: Comments created via the `addComment` tool appear in the "All Comments" list but are not visibly anchored to specific text in the Google Docs UI. They will show "original content deleted" instead of highlighting the intended text range. This is a limitation of the Google Drive API v3 when working with Google Docs files.

- **Workaround**: Comments created manually in the Google Docs UI are properly anchored
- **Other Operations**: Reply, delete, and list operations work correctly on all comments regardless of how they were created

### Comment Resolution

**Resolved Status May Not Persist**: The `resolveComment` tool attempts to mark comments as resolved, but the Drive API v3 does not fully support this operation for Google Docs files. The resolved status may not persist or be visible in the Google Docs UI.

- **Workaround**: Resolve comments manually in the Google Docs web interface

### Converted Documents

**Limited Support for Converted Documents**: Some Google Docs that were converted from other formats (especially Microsoft Word documents) may not support all Docs API operations. You may encounter errors like "This operation is not supported for this document" when trying to read or modify these files.

## Troubleshooting

- **Claude shows "Failed" or "Could not attach":**
  - Double-check the absolute path in `mcp_config.json`.
  - Ensure you ran `npm run build` successfully and the `dist` folder exists.
  - Try running the command from `mcp_config.json` manually in your terminal: `node /PATH/TO/YOUR/CLONED/REPO/mcp-googledocs-server/dist/server.js`. Look for any errors printed.
  - Check the Claude Desktop logs (see the official MCP debugging guide).
  - Make sure all `console.log` status messages in the server code were changed to `console.error`.
- **Google Authorization Errors:**
  - Ensure you enabled the correct APIs (Docs, Sheets, Drive).
  - Make sure you added your email as a Test User on the OAuth Consent Screen.
  - Verify the `credentials.json` file is correctly placed in the project root.
  - **If you're upgrading from an older version:** You may need to delete your existing `token.json` file and re-authenticate to grant the new Sheets API scope.
- **Tab-related Errors:**
  - If you get "Tab with ID not found", use `listDocumentTabs` to see all available tab IDs
  - Ensure you're using the correct tab ID format (typically a short alphanumeric string)
  - Single-tab documents don't require `tabId` - operations work on the document body automatically

---

## License

This project is licensed under the MIT License - see the `LICENSE` file for details. (Note: You should add a `LICENSE` file containing the MIT License text to your repository).

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=a-bonus/google-docs-mcp&type=date&legend=top-left)](https://www.star-history.com/#a-bonus/google-docs-mcp&type=date&legend=top-left)
