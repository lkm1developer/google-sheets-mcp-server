#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
  Tool
} from '@modelcontextprotocol/sdk/types.js';
import { GoogleSheetsClient } from './google-sheets-client.js';
import * as dotenv from 'dotenv';
import { parseArgs } from 'node:util';

// Load environment variables
dotenv.config();

// Parse command line arguments
const { values } = parseArgs({
  options: {
    'GOOGLE_APPLICATION_CREDENTIALS': { type: 'string' },
    'GOOGLE_APPLICATION_CREDENTIALS_JSON': { type: 'string' },
    'GOOGLE_API_KEY': { type: 'string' },
    'GOOGLE_CLOUD_PROJECT_ID': { type: 'string' },
    'client_id': { type: 'string' },
    'client_secret': { type: 'string' },
    'refresh_token': { type: 'string' }
  }
});

// Get authentication parameters
const serviceAccountPath = values['GOOGLE_APPLICATION_CREDENTIALS'] || process.env.GOOGLE_APPLICATION_CREDENTIALS;
const serviceAccountJson = values['GOOGLE_APPLICATION_CREDENTIALS_JSON'] || process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON;
const apiKey = values['GOOGLE_API_KEY'] || process.env.GOOGLE_API_KEY;
const projectId = values['GOOGLE_CLOUD_PROJECT_ID'] || process.env.GOOGLE_CLOUD_PROJECT_ID;
const oauthClientId = values['client_id'] || process.env.client_id;
const oauthClientSecret = values['client_secret'] || process.env.client_secret;
const oauthRefreshToken = values['refresh_token'] || process.env.refresh_token;
let isOAuthEnabled = false;
if (!serviceAccountPath && !serviceAccountJson && !apiKey && !(oauthClientId && oauthClientSecret && oauthRefreshToken)) {
  throw new Error('Either GOOGLE_APPLICATION_CREDENTIALS, GOOGLE_APPLICATION_CREDENTIALS_JSON, GOOGLE_API_KEY, or OAuth credentials are required');
}
if(oauthClientId && oauthClientSecret && oauthRefreshToken) {
  isOAuthEnabled = true
}
console.log({ oauthClientId , oauthClientSecret , oauthRefreshToken, isOAuthEnabled})
if (!projectId && !isOAuthEnabled) {
  throw new Error('GOOGLE_CLOUD_PROJECT_ID environment variable is required');
}

class GoogleSheetsServer {
  // Core server properties
  private server: Server;
  private googleSheets: GoogleSheetsClient;

  constructor() {
    this.server = new Server(
      {
        name: 'google-sheets-manager',
        version: '0.1.0',
      },
      {
        capabilities: {
          resources: {},
          tools: {},
        },
      }
    );
    if (serviceAccountPath) {
      console.log(`Using service account: ${serviceAccountPath}`);
    }
    if (apiKey) {
      console.log(`Using API key: ${apiKey.substring(0, 4)}...`);
    }
    if (oauthClientId && oauthClientSecret && oauthRefreshToken) {
      console.log(`Using OAuth credentials with client ID: ${oauthClientId.substring(0, 4)}...`);
    }
    console.log(`Using project ID: ${projectId}`);
    let isOAuthEnabled = false;
    if(oauthClientId && oauthClientSecret && oauthRefreshToken) {
      isOAuthEnabled = true
    }
    this.googleSheets = new GoogleSheetsClient({
      serviceAccountPath,
      serviceAccountJson,
      apiKey,
      projectId,
      oauthClientId,
      oauthClientSecret,
      oauthRefreshToken,
      isOAuth: isOAuthEnabled
    });
    
    this.setupToolHandlers();
    this.setupErrorHandling();
  }

  private setupErrorHandling(): void {
    this.server.onerror = (error) => {
      console.error('[MCP Error]', error);
    };

    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });

    process.on('uncaughtException', (error) => {
      console.error('Uncaught exception:', error);
    });

    process.on('unhandledRejection', (reason, promise) => {
      console.error('Unhandled rejection at:', promise, 'reason:', reason);
    });
  }

  private setupToolHandlers(): void {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      // Define available tools
      const tools: Tool[] = [
        {
          name: 'google_sheets_create',
          description: 'Create a new Google Sheet',
          inputSchema: {
            type: 'object',
            properties: {
              title: { 
                type: 'string', 
                description: 'Title of the spreadsheet' 
              },
              sheets: { 
                type: 'array', 
                items: { type: 'string' },
                description: 'Array of sheet names to create (optional)' 
              }
            },
            required: ['title']
          }
        },
        {
          name: 'google_sheets_get',
          description: 'Get a Google Sheet by ID',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: { 
                type: 'string', 
                description: 'ID of the spreadsheet to retrieve' 
              },
              includeGridData: { 
                type: 'boolean', 
                description: 'Whether to include grid data (cell values)' 
              }
            },
            required: ['spreadsheetId']
          }
        },
        {
          name: 'google_sheets_update_values',
          description: 'Update values in a Google Sheet',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: { 
                type: 'string', 
                description: 'ID of the spreadsheet to update' 
              },
              range: { 
                type: 'string', 
                description: 'Range to update (e.g., "Sheet1!A1:B2")' 
              },
              values: { 
                type: 'array', 
                description: '2D array of values to update' 
              },
              valueInputOption: { 
                type: 'string', 
                description: 'How to interpret the values (RAW or USER_ENTERED)' 
              }
            },
            required: ['spreadsheetId', 'range', 'values']
          }
        },
        {
          name: 'google_sheets_append_values',
          description: 'Append values to a Google Sheet',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: { 
                type: 'string', 
                description: 'ID of the spreadsheet to update' 
              },
              range: { 
                type: 'string', 
                description: 'Range to append to (e.g., "Sheet1!A1")' 
              },
              values: { 
                type: 'array', 
                description: '2D array of values to append' 
              },
              valueInputOption: { 
                type: 'string', 
                description: 'How to interpret the values (RAW or USER_ENTERED)' 
              }
            },
            required: ['spreadsheetId', 'range', 'values']
          }
        },
        {
          name: 'google_sheets_get_values',
          description: 'Get values from a Google Sheet',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: { 
                type: 'string', 
                description: 'ID of the spreadsheet to read' 
              },
              range: { 
                type: 'string', 
                description: 'Range to read (e.g., "Sheet1!A1:B2")' 
              }
            },
            required: ['spreadsheetId', 'range']
          }
        },
        {
          name: 'google_sheets_clear_values',
          description: 'Clear values from a Google Sheet',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: { 
                type: 'string', 
                description: 'ID of the spreadsheet to clear' 
              },
              range: { 
                type: 'string', 
                description: 'Range to clear (e.g., "Sheet1!A1:B2")' 
              }
            },
            required: ['spreadsheetId', 'range']
          }
        },
        {
          name: 'google_sheets_add_sheet',
          description: 'Add a new sheet to an existing spreadsheet',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: { 
                type: 'string', 
                description: 'ID of the spreadsheet to update' 
              },
              sheetTitle: { 
                type: 'string', 
                description: 'Title of the new sheet' 
              }
            },
            required: ['spreadsheetId', 'sheetTitle']
          }
        },
        {
          name: 'google_sheets_delete_sheet',
          description: 'Delete a sheet from a spreadsheet',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: { 
                type: 'string', 
                description: 'ID of the spreadsheet' 
              },
              sheetId: { 
                type: 'number', 
                description: 'ID of the sheet to delete' 
              }
            },
            required: ['spreadsheetId', 'sheetId']
          }
        },
        {
          name: 'google_sheets_list',
          description: 'List Google Sheets accessible to the authenticated user',
          inputSchema: {
            type: 'object',
            properties: {
              pageSize: { 
                type: 'number', 
                description: 'Number of spreadsheets to return (default: 10)' 
              },
              pageToken: { 
                type: 'string', 
                description: 'Token for pagination' 
              }
            }
          }
        },
        {
          name: 'google_sheets_delete',
          description: 'Delete a Google Sheet',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: { 
                type: 'string', 
                description: 'ID of the spreadsheet to delete' 
              }
            },
            required: ['spreadsheetId']
          }
        },
        {
          name: 'google_sheets_share',
          description: 'Share a Google Sheet with specific users',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: { 
                type: 'string', 
                description: 'ID of the spreadsheet to share' 
              },
              emailAddress: { 
                type: 'string', 
                description: 'Email address to share with' 
              },
              role: { 
                type: 'string', 
                description: 'Role to assign (reader, writer, commenter)' 
              }
            },
            required: ['spreadsheetId', 'emailAddress']
          }
        },
        {
          name: 'google_sheets_search',
          description: 'Search for Google Sheets by title',
          inputSchema: {
            type: 'object',
            properties: {
              query: { 
                type: 'string', 
                description: 'Search query for spreadsheet title' 
              },
              pageSize: { 
                type: 'number', 
                description: 'Number of results to return (default: 10)' 
              },
              pageToken: { 
                type: 'string', 
                description: 'Token for pagination' 
              }
            },
            required: ['query']
          }
        },
        {
          name: 'google_sheets_format_cells',
          description: 'Format cells in a Google Sheet',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: { 
                type: 'string', 
                description: 'ID of the spreadsheet' 
              },
              sheetId: { 
                type: 'number', 
                description: 'ID of the sheet' 
              },
              range: { 
                type: 'object', 
                description: 'Range to format',
                properties: {
                  startRowIndex: { type: 'number' },
                  endRowIndex: { type: 'number' },
                  startColumnIndex: { type: 'number' },
                  endColumnIndex: { type: 'number' }
                },
                required: ['startRowIndex', 'endRowIndex', 'startColumnIndex', 'endColumnIndex']
              },
              format: { 
                type: 'object', 
                description: 'Format to apply' 
              }
            },
            required: ['spreadsheetId', 'sheetId', 'range', 'format']
          }
        },
        {
          name: 'google_sheets_verify_connection',
          description: 'Verify connection with Google Sheets API and check credentials',
          inputSchema: {
            type: 'object',
            properties: {}
          }
        },
        {
          name: 'write_to_sheet',
          description: 'Write data with headers and rows to a Google Sheet',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: { 
                type: 'string', 
                description: 'ID of the spreadsheet to write to' 
              },
              sheetName: { 
                type: 'string', 
                description: 'Name of the sheet to write to' 
              },
              data: { 
                type: 'object',
                description: 'Data to write to the sheet',
                properties: {
                  headers: {
                    type: 'array',
                    items: { type: 'string' },
                    description: 'Array of column headers'
                  },
                  rows: {
                    type: 'array',
                    items: {
                      type: 'array',
                      items: {}
                    },
                    description: 'Array of row data (2D array)'
                  }
                },
                required: ['headers', 'rows']
              },
              clearExisting: { 
                type: 'boolean', 
                description: 'Whether to clear existing data before writing (default: false)' 
              }
            },
            required: ['spreadsheetId', 'sheetName', 'data']
          }
        }
      ];
      
      return { tools };
    });

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        const args = request.params.arguments ?? {};

        switch (request.params.name) {
          case 'google_sheets_create': {
            const result = await this.googleSheets.createSpreadsheet(
              args.title as string,
              args.sheets as string[] | undefined
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'google_sheets_get': {
            const result = await this.googleSheets.getSpreadsheet(
              args.spreadsheetId as string,
              args.includeGridData as boolean | undefined
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'google_sheets_update_values': {
            const result = await this.googleSheets.updateValues(
              args.spreadsheetId as string,
              args.range as string,
              args.values as any[][],
              args.valueInputOption as string | undefined
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'google_sheets_append_values': {
            const result = await this.googleSheets.appendValues(
              args.spreadsheetId as string,
              args.range as string,
              args.values as any[][],
              args.valueInputOption as string | undefined
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'google_sheets_get_values': {
            const result = await this.googleSheets.getValues(
              args.spreadsheetId as string,
              args.range as string
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'google_sheets_clear_values': {
            const result = await this.googleSheets.clearValues(
              args.spreadsheetId as string,
              args.range as string
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'google_sheets_add_sheet': {
            const result = await this.googleSheets.addSheet(
              args.spreadsheetId as string,
              args.sheetTitle as string
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'google_sheets_delete_sheet': {
            const result = await this.googleSheets.deleteSheet(
              args.spreadsheetId as string,
              args.sheetId as number
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'google_sheets_list': {
            const result = await this.googleSheets.listSpreadsheets(
              args.pageSize as number | undefined,
              args.pageToken as string | undefined
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'google_sheets_delete': {
            const result = await this.googleSheets.deleteSpreadsheet(
              args.spreadsheetId as string
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'google_sheets_share': {
            const result = await this.googleSheets.shareSpreadsheet(
              args.spreadsheetId as string,
              args.emailAddress as string,
              args.role as string | undefined
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'google_sheets_search': {
            const result = await this.googleSheets.searchSpreadsheets(
              args.query as string,
              args.pageSize as number | undefined,
              args.pageToken as string | undefined
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'google_sheets_format_cells': {
            const result = await this.googleSheets.formatCells(
              args.spreadsheetId as string,
              args.sheetId as number,
              args.range as { 
                startRowIndex: number, 
                endRowIndex: number, 
                startColumnIndex: number, 
                endColumnIndex: number 
              },
              args.format as any
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'google_sheets_verify_connection': {
            const result = await this.googleSheets.verifyConnection();
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }
          
          case 'write_to_sheet': {
            const result = await this.googleSheets.writeToSheet(
              args.spreadsheetId as string,
              args.sheetName as string,
              args.data as { headers: string[], rows: any[][] },
              args.clearExisting as boolean | undefined
            );
            return {
              content: [{
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }]
            };
          }

          default:
            throw new McpError(
              ErrorCode.MethodNotFound,
              `Unknown tool: ${request.params.name}`
            );
        }
      } catch (error: any) {
        console.error(`Error executing tool ${request.params.name}:`, error);
        return {
          content: [{
            type: 'text',
            text: `Google Sheets API error: ${error.message}`
          }],
          isError: true,
        };
      }
    });
  }

  async run(): Promise<void> {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.log('Google Sheets MCP server started');
  }
}

export async function serve(): Promise<void> {
  const server = new GoogleSheetsServer();
  await server.run();
}

const server = new GoogleSheetsServer();
server.run().catch(console.error);
