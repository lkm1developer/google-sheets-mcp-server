import { google, sheets_v4 } from 'googleapis';
import * as dotenv from 'dotenv';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import * as crypto from 'crypto';

dotenv.config();

export class GoogleSheetsClient {
  private sheetsClient: sheets_v4.Sheets;
  private driveClient: any; // Using any type to bypass TypeScript errors
  private projectId: string;
  private auth: any;
  private tempFilePath: string | null = null;
  private isOAuth = false;

  // Clean up temporary files when the process exits
  constructor(options: {
    serviceAccountPath?: string;
    serviceAccountJson?: string;
    apiKey?: string;
    projectId?: string;
    oauthClientId?: string;
    oauthClientSecret?: string;
    oauthRefreshToken?: string;
    isOAuth: boolean;
  }) {
    // Set up cleanup handlers for temporary files
    process.on('exit', () => this.cleanup());
    process.on('SIGINT', () => {
      this.cleanup();
      process.exit(0);
    });
    process.on('SIGTERM', () => {
      this.cleanup();
      process.exit(0);
    });
    // Initialize the client
    // Get project ID from environment variables or constructor parameters
    this.projectId = options.projectId || process.env.GOOGLE_CLOUD_PROJECT_ID || '';
    console.log(`Using project ID: ${this.projectId} for Google Sheets API ${this.isOAuth}`);
    if (!this.projectId && !options.isOAuth) {
      throw new Error('GOOGLE_CLOUD_PROJECT_ID environment variable or projectId parameter is required');
    }
    
    // Check if API key is provided
    const apiKey = options.apiKey || process.env.GOOGLE_API_KEY;
    
    // Check if service account JSON content is provided
    const serviceAccountJson = options.serviceAccountJson || process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON;
    
    // Check if service account path is provided
    const serviceAccountPath = options.serviceAccountPath || process.env.GOOGLE_APPLICATION_CREDENTIALS;
    
    // Check if OAuth credentials are provided
    const oauthClientId = options.oauthClientId || process.env.client_id;
    const oauthClientSecret = options.oauthClientSecret || process.env.client_secret;
    const oauthRefreshToken = options.oauthRefreshToken || process.env.refresh_token;
    
    // We need either an API key, a service account (path or JSON), or OAuth credentials
    if (!apiKey && !serviceAccountPath && !serviceAccountJson && !(oauthClientId && oauthClientSecret && oauthRefreshToken)) {
      throw new Error('Either GOOGLE_API_KEY, GOOGLE_APPLICATION_CREDENTIALS, GOOGLE_APPLICATION_CREDENTIALS_JSON, or OAuth credentials must be provided');
    }
    
    if (apiKey) {
      // Use API key authentication
      console.log('Using API key authentication');
      this.auth = apiKey;
    } else if (serviceAccountJson) {
      // Use service account JSON content
      console.log('Using service account JSON content for authentication');
      
      try {
        // Create a temporary file with the JSON content
        const tempDir = os.tmpdir();
        const randomId = crypto.randomBytes(16).toString('hex');
        this.tempFilePath = path.join(tempDir, `google-credentials-${randomId}.json`);
        
        // Write the JSON content to the temporary file
        fs.writeFileSync(this.tempFilePath, serviceAccountJson);
        
        // Use the temporary file for authentication
        this.auth = new google.auth.GoogleAuth({
          keyFile: this.tempFilePath,
          scopes: [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
          ]
        });
      } catch (error: any) {
        throw new Error(`Failed to create temporary credentials file: ${error.message}`);
      }
    } else if (serviceAccountPath) {
      // Check if the credentials file exists
      if (!fs.existsSync(serviceAccountPath)) {
        throw new Error(`Service account key file not found at: ${serviceAccountPath}`);
      }
      
      // Use service account authentication
      console.log('Using service account authentication from file');
      this.auth = new google.auth.GoogleAuth({
        keyFile: serviceAccountPath,
        scopes: [
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive'
        ]
      });
    } else if (oauthClientId && oauthClientSecret && oauthRefreshToken) {
      // Use OAuth2 authentication
      console.log('Using OAuth2 authentication');
      const oauth2Client = new google.auth.OAuth2(
        oauthClientId,
        oauthClientSecret
      );
      
      oauth2Client.setCredentials({
        refresh_token: oauthRefreshToken
      });
      
      this.auth = oauth2Client;
    }
    
    // Create the sheets client
    this.sheetsClient = google.sheets({
      version: 'v4',
      auth: this.auth
    });
    
    // Create the drive client (needed for some operations like creating spreadsheets)
    this.driveClient = google.drive({
      version: 'v3',
      auth: this.auth
    });
    
    console.log('Google Sheets client initialized');
  }
  
  /**
   * Create a new Google Sheet
   * @param title Title of the spreadsheet
   * @param sheets Array of sheet names to create (optional)
   * @returns Created spreadsheet details
   */
  async createSpreadsheet(title: string, sheets?: string[]): Promise<any> {
    try {
      // Create a new spreadsheet
      const resource: sheets_v4.Schema$Spreadsheet = {
        properties: {
          title
        }
      };
      
      // If sheets are provided, add them to the resource
      if (sheets && sheets.length > 0) {
        resource.sheets = sheets.map(sheetName => ({
          properties: {
            title: sheetName
          }
        }));
      }
      
      const response = await this.sheetsClient.spreadsheets.create({
        requestBody: resource
      });
      
      const spreadsheetId = response.data.spreadsheetId;
      
      return {
        spreadsheetId,
        title,
        url: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
        spreadsheet: response.data
      };
    } catch (error: any) {
      console.error('Error creating spreadsheet:', error);
      return { error: error.message };
    }
  }
  
  /**
   * Get a Google Sheet by ID
   * @param spreadsheetId ID of the spreadsheet to retrieve
   * @param includeGridData Whether to include grid data (cell values)
   * @returns Spreadsheet details
   */
  async getSpreadsheet(spreadsheetId: string, includeGridData: boolean = false): Promise<any> {
    try {
      const response = await this.sheetsClient.spreadsheets.get({
        spreadsheetId,
        includeGridData
      });
      
      return response.data;
    } catch (error: any) {
      console.error('Error getting spreadsheet:', error);
      return { error: error.message };
    }
  }
  
  /**
   * Update values in a Google Sheet
   * @param spreadsheetId ID of the spreadsheet to update
   * @param range Range to update (e.g., 'Sheet1!A1:B2')
   * @param values 2D array of values to update
   * @param valueInputOption How to interpret the values (RAW or USER_ENTERED)
   * @returns Updated spreadsheet details
   */
  async updateValues(
    spreadsheetId: string, 
    range: string, 
    values: any[][], 
    valueInputOption: string = 'USER_ENTERED'
  ): Promise<any> {
    try {
      const response = await this.sheetsClient.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption,
        requestBody: {
          values
        }
      });
      
      return response.data;
    } catch (error: any) {
      console.error('Error updating values:', error);
      return { error: error.message };
    }
  }
  
  /**
   * Append values to a Google Sheet
   * @param spreadsheetId ID of the spreadsheet to update
   * @param range Range to append to (e.g., 'Sheet1!A1')
   * @param values 2D array of values to append
   * @param valueInputOption How to interpret the values (RAW or USER_ENTERED)
   * @returns Updated spreadsheet details
   */
  async appendValues(
    spreadsheetId: string, 
    range: string, 
    values: any[][], 
    valueInputOption: string = 'USER_ENTERED'
  ): Promise<any> {
    try {
      const response = await this.sheetsClient.spreadsheets.values.append({
        spreadsheetId,
        range,
        valueInputOption,
        requestBody: {
          values
        }
      });
      
      return response.data;
    } catch (error: any) {
      console.error('Error appending values:', error);
      return { error: error.message };
    }
  }
  
  /**
   * Get values from a Google Sheet
   * @param spreadsheetId ID of the spreadsheet to read
   * @param range Range to read (e.g., 'Sheet1!A1:B2')
   * @returns Values from the specified range
   */
  async getValues(spreadsheetId: string, range: string): Promise<any> {
    try {
      const response = await this.sheetsClient.spreadsheets.values.get({
        spreadsheetId,
        range
      });
      
      return response.data;
    } catch (error: any) {
      console.error('Error getting values:', error);
      return { error: error.message };
    }
  }
  
  /**
   * Clear values from a Google Sheet
   * @param spreadsheetId ID of the spreadsheet to clear
   * @param range Range to clear (e.g., 'Sheet1!A1:B2')
   * @returns Clear operation result
   */
  async clearValues(spreadsheetId: string, range: string): Promise<any> {
    try {
      const response = await this.sheetsClient.spreadsheets.values.clear({
        spreadsheetId,
        range
      });
      
      return response.data;
    } catch (error: any) {
      console.error('Error clearing values:', error);
      return { error: error.message };
    }
  }
  
  /**
   * Add a new sheet to an existing spreadsheet
   * @param spreadsheetId ID of the spreadsheet to update
   * @param sheetTitle Title of the new sheet
   * @returns Updated spreadsheet details
   */
  async addSheet(spreadsheetId: string, sheetTitle: string): Promise<any> {
    try {
      const response = await this.sheetsClient.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: sheetTitle
                }
              }
            }
          ]
        }
      });
      
      return response.data;
    } catch (error: any) {
      console.error('Error adding sheet:', error);
      return { error: error.message };
    }
  }
  
  /**
   * Delete a sheet from a spreadsheet
   * @param spreadsheetId ID of the spreadsheet
   * @param sheetId ID of the sheet to delete
   * @returns Updated spreadsheet details
   */
  async deleteSheet(spreadsheetId: string, sheetId: number): Promise<any> {
    try {
      const response = await this.sheetsClient.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              deleteSheet: {
                sheetId
              }
            }
          ]
        }
      });
      
      return response.data;
    } catch (error: any) {
      console.error('Error deleting sheet:', error);
      return { error: error.message };
    }
  }
  
  /**
   * List spreadsheets accessible to the authenticated user
   * @param pageSize Number of spreadsheets to return
   * @param pageToken Token for pagination
   * @returns List of spreadsheets
   */
  async listSpreadsheets(pageSize: number = 10, pageToken?: string): Promise<any> {
    try {
      const response = await this.driveClient.files.list({
        pageSize,
        pageToken,
        q: "mimeType='application/vnd.google-apps.spreadsheet'",
        fields: 'nextPageToken, files(id, name, createdTime, modifiedTime, webViewLink)'
      });
      
      return {
        spreadsheets: response.data.files,
        nextPageToken: response.data.nextPageToken
      };
    } catch (error: any) {
      console.error('Error listing spreadsheets:', error);
      return { error: error.message };
    }
  }
  
  /**
   * Delete a Google Sheet
   * @param spreadsheetId ID of the spreadsheet to delete
   * @returns Success status
   */
  async deleteSpreadsheet(spreadsheetId: string): Promise<any> {
    try {
      await this.driveClient.files.delete({
        fileId: spreadsheetId
      });
      
      return {
        success: true,
        spreadsheetId,
        message: `Spreadsheet ${spreadsheetId} successfully deleted`
      };
    } catch (error: any) {
      console.error('Error deleting spreadsheet:', error);
      return { 
        success: false,
        error: error.message 
      };
    }
  }
  
  /**
   * Share a Google Sheet with specific users
   * @param spreadsheetId ID of the spreadsheet to share
   * @param emailAddress Email address to share with
   * @param role Role to assign (reader, writer, commenter)
   * @returns Share details
   */
  async shareSpreadsheet(spreadsheetId: string, emailAddress: string, role: string = 'reader'): Promise<any> {
    try {
      const response = await this.driveClient.permissions.create({
        fileId: spreadsheetId,
        requestBody: {
          type: 'user',
          role,
          emailAddress
        }
      });
      
      return {
        success: true,
        spreadsheetId,
        permission: response.data
      };
    } catch (error: any) {
      console.error('Error sharing spreadsheet:', error);
      return { 
        success: false,
        error: error.message 
      };
    }
  }
  
  /**
   * Search for spreadsheets by title
   * @param query Search query
   * @param pageSize Number of results to return
   * @param pageToken Token for pagination
   * @returns Search results
   */
  async searchSpreadsheets(query: string, pageSize: number = 10, pageToken?: string): Promise<any> {
    try {
      const response = await this.driveClient.files.list({
        pageSize,
        pageToken,
        q: `mimeType='application/vnd.google-apps.spreadsheet' and name contains '${query}'`,
        fields: 'nextPageToken, files(id, name, createdTime, modifiedTime, webViewLink)'
      });
      
      return {
        spreadsheets: response.data.files,
        nextPageToken: response.data.nextPageToken
      };
    } catch (error: any) {
      console.error('Error searching spreadsheets:', error);
      return { error: error.message };
    }
  }
  
  /**
   * Format cells in a Google Sheet
   * @param spreadsheetId ID of the spreadsheet
   * @param sheetId ID of the sheet
   * @param range Range to format (startRowIndex, endRowIndex, startColumnIndex, endColumnIndex)
   * @param format Format to apply
   * @returns Updated spreadsheet details
   */
  async formatCells(
    spreadsheetId: string, 
    sheetId: number, 
    range: { 
      startRowIndex: number, 
      endRowIndex: number, 
      startColumnIndex: number, 
      endColumnIndex: number 
    }, 
    format: any
  ): Promise<any> {
    try {
      const response = await this.sheetsClient.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              repeatCell: {
                range: {
                  sheetId,
                  ...range
                },
                cell: {
                  userEnteredFormat: format
                },
                fields: 'userEnteredFormat'
              }
            }
          ]
        }
      });
      
      return response.data;
    } catch (error: any) {
      console.error('Error formatting cells:', error);
      return { error: error.message };
    }
  }
  
  /**
   * Verify connection with Google Sheets API
   * @returns Connection status and details
   */
  async verifyConnection(): Promise<any> {
    try {
      // Try to list a single spreadsheet to verify connectivity
      const response = await this.driveClient.files.list({
        pageSize: 1,
        q: "mimeType='application/vnd.google-apps.spreadsheet'",
        fields: 'files(id, name)'
      });
      
      return {
        connected: true,
        projectId: this.projectId,
        timestamp: new Date().toISOString(),
        details: {
          authType: this.getAuthType(),
          apiVersion: 'v4',
          spreadsheetCount: response.data.files?.length || 0
        }
      };
    } catch (error: any) {
      console.error('Error verifying Google Sheets API connection:', error);
      
      // Return detailed error information
      return {
        connected: false,
        projectId: this.projectId,
        timestamp: new Date().toISOString(),
        error: {
          message: error.message,
          code: error.code || 'UNKNOWN_ERROR',
          details: error.details || error.stack
        }
      };
    }
  }
  
  /**
   * Helper method to determine the authentication type being used
   * @returns Authentication type string
   */
  private getAuthType(): string {
    if (typeof this.auth === 'string') {
      return 'API Key';
    } else if (this.auth instanceof google.auth.GoogleAuth) {
      return 'Service Account';
    } else if (this.auth instanceof google.auth.OAuth2) {
      return 'OAuth2';
    }
    return 'Unknown';
  }

  /**
   * Clean up any temporary files created for service account JSON
   */
  private cleanup(): void {
    if (this.tempFilePath && fs.existsSync(this.tempFilePath)) {
      try {
        fs.unlinkSync(this.tempFilePath);
        console.log(`Cleaned up temporary credentials file: ${this.tempFilePath}`);
        this.tempFilePath = null;
      } catch (error: any) {
        console.error(`Failed to clean up temporary credentials file: ${error.message}`);
      }
    }
  }
}
