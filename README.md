# Google Sheets MCP Server

A powerful Model Context Protocol (MCP) server implementation for seamless Google Sheets API integration, enabling AI assistants to create, read, update, and manage Google Sheets.

## Features

- Create, read, update, and delete Google Sheets
- Manage sheet data with cell-level operations
- Format cells and ranges
- Share spreadsheets with other users
- Search for spreadsheets
- Comprehensive authentication options

## Authentication Options

This MCP server supports multiple authentication methods:

1. **Service Account Authentication** (recommended for production)
   - Provide a path to a service account key file
   - Or provide the service account JSON content directly

2. **API Key Authentication** (simpler for development)
   - Provide a Google API key

3. **OAuth2 Authentication** (for user-specific operations)
   - Provide OAuth client ID, client secret, and refresh token

## Setup

### Prerequisites

- Node.js (v16 or higher)
- A Google Cloud Project with the Google Sheets API enabled
- Authentication credentials (see below)

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/google-sheets-mcp-server.git
   cd google-sheets-mcp-server
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Set up authentication:
   - Create a `.env` file based on the `.env.example` template
   - Add your Google Cloud credentials

4. Build the project:
   ```bash
   npm run build
   ```

5. Start the server:
   ```bash
   npm start
   ```

### OAuth Setup

For operations that require user consent (like creating/editing spreadsheets), you'll need OAuth credentials:

1. Create OAuth credentials in the Google Cloud Console
2. Update the `CLIENT_ID` and `CLIENT_SECRET` in `src/get-refresh-token.js`
3. Install required dependencies:
   ```bash
   npm install open server-destroy
   ```
4. Run the script to get a refresh token:
   ```bash
   node src/get-refresh-token.js
   ```
5. Follow the browser prompts to authorize the application
6. Copy the refresh token to your `.env` file

## Available Tools

The server provides the following tools:

- `google_sheets_create`: Create a new Google Sheet
- `google_sheets_get`: Get a Google Sheet by ID
- `google_sheets_update_values`: Update values in a Google Sheet
- `google_sheets_append_values`: Append values to a Google Sheet
- `google_sheets_get_values`: Get values from a Google Sheet
- `google_sheets_clear_values`: Clear values from a Google Sheet
- `google_sheets_add_sheet`: Add a new sheet to an existing spreadsheet
- `google_sheets_delete_sheet`: Delete a sheet from a spreadsheet
- `google_sheets_list`: List Google Sheets accessible to the authenticated user
- `google_sheets_delete`: Delete a Google Sheet
- `google_sheets_share`: Share a Google Sheet with specific users
- `google_sheets_search`: Search for Google Sheets by title
- `google_sheets_format_cells`: Format cells in a Google Sheet
- `google_sheets_verify_connection`: Verify connection with Google Sheets API

## Example Usage

Here's an example of how to use this MCP server with Claude:

```
You can now use Google Sheets! Try these commands:

1. Create a new spreadsheet:
   "Create a new Google Sheet titled 'Monthly Budget'"

2. Add data to the spreadsheet:
   "Add expense categories and amounts to my Monthly Budget spreadsheet"

3. Format the spreadsheet:
   "Format the header row to be bold and centered"

4. Share the spreadsheet:
   "Share my Monthly Budget spreadsheet with john@example.com"
```

## License

MIT
