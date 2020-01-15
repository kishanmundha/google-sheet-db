/* tslint:disable:member-ordering */
import * as fs from 'fs';
import { google } from 'googleapis';
import * as readline from 'readline';
import { ISheetProperty } from './types';

const debug = (message: string, ...optionalParams: any[]) => {
  // tslint:disable-next-line: no-console
  console.log('GoogleSheetDb:', message, ...optionalParams);
};

export interface IOptions {
  tokenPath?: string;
  credentialPath?: string;
  scopes?: string[];
  auth?: any;
}

const DEFAULT_SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const DEFAULT_TOKEN_PATH = 'token.json';
const DEFAULT_CREDENTIAL_PATH = 'credentials.json';

class GoogleSheet {
  private SPREADSHEET_ID: string = null;

  // If modifying these scopes, delete token.json.
  private SCOPES = DEFAULT_SCOPES;
  // The file token.json stores the user's access and refresh tokens, and is
  // created automatically when the authorization flow completes for the first
  // time.
  private TOKEN_PATH = DEFAULT_TOKEN_PATH;
  private CREDENTIAL_PATH = DEFAULT_CREDENTIAL_PATH;

  private auth = null;

  constructor(spreadsheetId: string, options?: IOptions) {
    this.SPREADSHEET_ID = spreadsheetId;

    if (options) {
      if (options.scopes) {
        this.SCOPES = options.scopes;
      }
      if (options.tokenPath) {
        this.TOKEN_PATH = options.tokenPath;
      }
      if (options.credentialPath) {
        this.CREDENTIAL_PATH = options.credentialPath;
      }

      if (options.auth) {
        this.auth = options.auth;
      }
    }
  }

  /**
   * Get and store new token after prompting for user authorization, and then
   * execute the given callback with the authorized OAuth2 client.
   * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
   * @param {getEventsCallback} callback The callback for the authorized client.
   */
  private getNewToken(oAuth2Client: any, callback: any) {
    const authUrl = oAuth2Client.generateAuthUrl({
      access_type: 'offline',
      scope: this.SCOPES,
    });
    debug('Authorize this app by visiting this url:', authUrl);
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });
    rl.question('Enter the code from that page here: ', (code) => {
      rl.close();
      oAuth2Client.getToken(code, (err: any, token: any) => {
        if (err) {
          return debug('Error while trying to retrieve access token', err);
        }
        oAuth2Client.setCredentials(token);
        // Store the token to disk for later program executions
        // tslint:disable-next-line:no-shadowed-variable
        fs.writeFile(this.TOKEN_PATH, JSON.stringify(token), (err: any) => {
          if (err) {
            return debug(err);
          }
          debug('Token stored to', this.TOKEN_PATH);
        });
        callback(oAuth2Client);
      });
    });
  }

  /**
   * Create an OAuth2 client with the given credentials, and then execute the
   * given callback function.
   * @param {Object} credentials The authorization client credentials.
   * @param {function} callback The callback to call with the authorized client.
   */
  private authorize(credentials: any, callback: any) {
    const { client_secret, client_id, redirect_uris } = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(
      client_id, client_secret, redirect_uris[0]);

    // Check if we have previously stored a token.
    fs.readFile(this.TOKEN_PATH, (err, token) => {
      if (err) {
        return this.getNewToken(oAuth2Client, callback);
      }
      oAuth2Client.setCredentials(JSON.parse(token.toString()));
      callback(oAuth2Client);
    });
  }

  public authenticate() {
    if (this.auth) {
      return Promise.resolve();
    }
    
    return new Promise((resolve, reject) => {
      // Load client secrets from a local file.
      fs.readFile(this.CREDENTIAL_PATH, (err, content) => {
        if (err) {
          return reject('Error loading client secret file:' + err.message);
        }
        // Authorize a client with credentials, then call the Google Sheets API.
        this.authorize(JSON.parse(content.toString()), (result: any) => {
          this.auth = result;
          resolve();
        });
      });
    });
  }

  public async getSheets(): Promise<ISheetProperty[]> {
    const sheets = google.sheets({ version: 'v4', auth: this.auth });
    const result = await sheets.spreadsheets.get({
      spreadsheetId: this.SPREADSHEET_ID,
      includeGridData: false,
    });

    return result.data.sheets.map(x => x.properties);
  }

  public async createSheet(title: string): Promise<ISheetProperty[]> {
    const sheets = google.sheets({ version: 'v4', auth: this.auth });
    const result = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: this.SPREADSHEET_ID,
      requestBody: {
        includeSpreadsheetInResponse: true,
        requests: [{
          addSheet: {
            properties: {
              title,
            },
          },
        }],
      },
    });

    return result.data.updatedSpreadsheet.sheets.map(x => x.properties);
  }

  public async insertRows(sheetId: number, startIndex: number, count: number): Promise<void> {
    const sheets = google.sheets({ version: 'v4', auth: this.auth });
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: this.SPREADSHEET_ID,
      requestBody: {
        requests: [{
          insertDimension: {
            range: {
              sheetId,
              dimension: 'ROWS',
              startIndex,
              endIndex: startIndex + count,
            },
            inheritFromBefore: false,
          },
        }],
      },
    });
  }

  public async removeRows(sheetId: number, startIndex: number, count: number): Promise<void> {
    console.log('remove row', startIndex, count);
    const sheets = google.sheets({ version: 'v4', auth: this.auth });
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: this.SPREADSHEET_ID,
      requestBody: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId,
              dimension: 'ROWS',
              startIndex,
              endIndex: startIndex + count,
            },
          },
        }],
      },
    });
  }

  public async writeData(range: string, values: string[][]): Promise<void> {
    const sheets = google.sheets({ version: 'v4', auth: this.auth });
    await sheets.spreadsheets.values.update({
      spreadsheetId: this.SPREADSHEET_ID,
      range,
      valueInputOption: 'RAW',
      requestBody: {
        values,
      },
    });
  }

  public async applyHeaderStyle(sheetId: number): Promise<void> {
    const sheets = google.sheets({ version: 'v4', auth: this.auth });
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: this.SPREADSHEET_ID,
      requestBody: {
        requests: [{
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 0,
              endRowIndex: 1,
            },
            cell: {
              userEnteredFormat: {
                textFormat: {
                  bold: true,
                },
              },
            },
            fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
          },
        }],
      },
    });
  }

  public async readData(range: string) {
    const sheets = google.sheets({ version: 'v4', auth: this.auth });
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: this.SPREADSHEET_ID,
      range,
    });

    return result.data.values;
  }
}

export default GoogleSheet;
