// src/auth.ts
import { google } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import { JWT } from 'google-auth-library'; // ADDED: Import for Service Account client
import * as fs from 'fs/promises';
import * as path from 'path';
import * as readline from 'readline/promises';
import { fileURLToPath } from 'url';

// --- Calculate paths relative to this script file (ESM way) ---
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRootDir = path.resolve(__dirname, '..');

// Support custom paths via env vars for multi-account setups
const TOKEN_PATH = process.env.TOKEN_PATH || path.join(projectRootDir, 'token.json');
const CREDENTIALS_PATH = process.env.CREDENTIALS_PATH || path.join(projectRootDir, 'credentials.json');
// --- End of path calculation ---

const SCOPES = [
  'https://www.googleapis.com/auth/documents',
  'https://www.googleapis.com/auth/drive', // Full Drive access for listing, searching, and document discovery
  'https://www.googleapis.com/auth/spreadsheets', // Google Sheets API access
  'https://www.googleapis.com/auth/presentations' // Google Slides API access
];

// --- NEW FUNCTION: Handles Service Account Authentication ---
// This entire function is new. It is called only when the
// SERVICE_ACCOUNT_PATH environment variable is set.
// Supports domain-wide delegation via GOOGLE_IMPERSONATE_USER env var.
async function authorizeWithServiceAccount(): Promise<JWT> {
  const serviceAccountPath = process.env.SERVICE_ACCOUNT_PATH!; // We know this is set if we are in this function
  const impersonateUser = process.env.GOOGLE_IMPERSONATE_USER; // Optional: email of user to impersonate
  try {
    const keyFileContent = await fs.readFile(serviceAccountPath, 'utf8');
    const serviceAccountKey = JSON.parse(keyFileContent);

    const auth = new JWT({
      email: serviceAccountKey.client_email,
      key: serviceAccountKey.private_key,
      scopes: SCOPES,
      subject: impersonateUser, // Enables domain-wide delegation when set
    });
    await auth.authorize();
    if (impersonateUser) {
      console.error(`Service Account authentication successful, impersonating: ${impersonateUser}`);
    } else {
      console.error('Service Account authentication successful!');
    }
    return auth;
  } catch (error: any) {
    if (error.code === 'ENOENT') {
      console.error(`FATAL: Service account key file not found at path: ${serviceAccountPath}`);
      throw new Error(`Service account key file not found. Please check the path in SERVICE_ACCOUNT_PATH.`);
    }
    console.error('FATAL: Error loading or authorizing the service account key:', error.message);
    throw new Error('Failed to authorize using the service account. Ensure the key file is valid and the path is correct.');
  }
}
// --- END OF NEW FUNCTION---

async function loadSavedCredentialsIfExist(): Promise<OAuth2Client | null> {
  try {
    const content = await fs.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content.toString());
    const { client_secret, client_id, redirect_uris } = await loadClientSecrets();
    const client = new google.auth.OAuth2(client_id, client_secret, redirect_uris?.[0]);
    client.setCredentials(credentials);
    return client;
  } catch (err) {
    return null;
  }
}

async function loadClientSecrets() {
  const content = await fs.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content.toString());
  const key = keys.installed || keys.web;
   if (!key) throw new Error("Could not find client secrets in credentials.json.");
  return {
      client_id: key.client_id,
      client_secret: key.client_secret,
      redirect_uris: key.redirect_uris || ['http://localhost:3000/'], // Default for web clients
      client_type: keys.web ? 'web' : 'installed'
  };
}

async function saveCredentials(client: OAuth2Client): Promise<void> {
  const { client_secret, client_id } = await loadClientSecrets();
  const payload = JSON.stringify({
    type: 'authorized_user',
    client_id: client_id,
    client_secret: client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFile(TOKEN_PATH, payload);
  console.error('Token stored to', TOKEN_PATH);
}

async function authenticate(): Promise<OAuth2Client> {
  const { client_secret, client_id, redirect_uris, client_type } = await loadClientSecrets();
  // For web clients, use the configured redirect URI; for desktop clients, use 'urn:ietf:wg:oauth:2.0:oob'
  const redirectUri = client_type === 'web' ? redirect_uris[0] : 'urn:ietf:wg:oauth:2.0:oob';
  console.error(`DEBUG: Using redirect URI: ${redirectUri}`);
  console.error(`DEBUG: Client type: ${client_type}`);
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirectUri);

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

  const authorizeUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES.join(' '),
  });

  console.error('DEBUG: Generated auth URL:', authorizeUrl);
  console.error('Authorize this app by visiting this url:', authorizeUrl);
  const code = await rl.question('Enter the code from that page here: ');
  rl.close();

  try {
    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);
    if (tokens.refresh_token) { // Save only if we got a refresh token
         await saveCredentials(oAuth2Client);
    } else {
         console.error("Did not receive refresh token. Token might expire.");
    }
    console.error('Authentication successful!');
    return oAuth2Client;
  } catch (err) {
    console.error('Error retrieving access token', err);
    throw new Error('Authentication failed');
  }
}

// --- MODIFIED: The Main Exported Function ---
// This function now acts as a router. It checks for the environment
// variable and decides which authentication method to use.
export async function authorize(): Promise<OAuth2Client | JWT> {
  // Check if the Service Account environment variable is set.
  if (process.env.SERVICE_ACCOUNT_PATH) {
    console.error('Service account path detected. Attempting service account authentication...');
    return authorizeWithServiceAccount();
  } else {
    // If not, execute the original OAuth 2.0 flow exactly as it was.
    console.error('No service account path detected. Falling back to standard OAuth 2.0 flow...');
    let client = await loadSavedCredentialsIfExist();
    if (client) {
      // Optional: Add token refresh logic here if needed, though library often handles it.
      console.error('Using saved credentials.');
      return client;
    }
    console.error('Starting authentication flow...');
    client = await authenticate();
    return client;
  }
}
// --- END OF MODIFIED: The Main Exported Function ---
