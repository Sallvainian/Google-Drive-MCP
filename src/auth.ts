// src/auth.ts
import { google } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import { JWT } from 'google-auth-library'; // ADDED: Import for Service Account client
import { randomBytes } from 'crypto';
import * as fs from 'fs/promises';
import * as path from 'path';
import * as http from 'http';
import { fileURLToPath, URL } from 'url';

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
  'https://www.googleapis.com/auth/presentations', // Google Slides API access
  'https://mail.google.com/', // Full Gmail access for email operations
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

function installTokenRefreshListener(client: OAuth2Client): void {
  client.removeAllListeners('tokens');
  client.on('tokens', (tokens) => {
    if (tokens.refresh_token) {
      console.error('Received rotated refresh token, persisting to disk...');
      client.setCredentials(tokens);
      saveCredentials(client).catch((err) => {
        console.error('Failed to persist rotated refresh token:', err.message);
      });
    }
  });
}

function isRecoverableAuthError(msg: string): boolean {
  return (
    msg.includes('invalid_grant') ||
    msg.includes('invalid_client') ||
    msg.includes('token has been expired') ||
    msg.includes('token has been revoked')
  );
}

async function getAuthenticatedEmail(client: OAuth2Client | JWT): Promise<string> {
  const drive = google.drive({ version: 'v3', auth: client as OAuth2Client });
  const { data } = await drive.about.get({ fields: 'user' });
  return data.user?.emailAddress ?? '(unknown)';
}

export function accountMatchesRequired(email: string): boolean {
  const required = process.env.REQUIRED_ACCOUNT_EMAIL;
  if (!required) return true;
  return email.toLowerCase() === required.toLowerCase();
}

async function isClientForRequiredAccount(client: OAuth2Client): Promise<boolean> {
  if (!process.env.REQUIRED_ACCOUNT_EMAIL) return true;
  const actual = await getAuthenticatedEmail(client);
  if (!accountMatchesRequired(actual)) {
    console.error(
      `Saved credentials are for "${actual}" but do not match REQUIRED_ACCOUNT_EMAIL. Will re-authenticate.`
    );
    return false;
  }
  return true;
}

async function loadSavedCredentialsIfExist(): Promise<OAuth2Client | null> {
  // Prefer GOOGLE_REFRESH_TOKEN env var over token.json file
  if (process.env.GOOGLE_REFRESH_TOKEN) {
    try {
      const { client_id, client_secret, redirect_uris } = await loadClientSecrets();
      const client = new google.auth.OAuth2(client_id, client_secret, redirect_uris?.[0]);
      client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN, expiry_date: 1 });
      installTokenRefreshListener(client);
      const { credentials } = await client.refreshAccessToken();
      // Persist if Google rotated the refresh token
      if (credentials.refresh_token && credentials.refresh_token !== process.env.GOOGLE_REFRESH_TOKEN) {
        console.error('Refresh token was rotated during startup verification, saving...');
        await saveCredentials(client);
      }
      if (!(await isClientForRequiredAccount(client))) {
        // Fall through to token file, then interactive OAuth (with login_hint pinning the right account)
        return null;
      }
      return client;
    } catch (err: unknown) {
      const msg = ((err as Error).message || '').toLowerCase();
      if (isRecoverableAuthError(msg)) {
        // Fall through to token file — don't trigger interactive OAuth for a stale env var
        console.error('GOOGLE_REFRESH_TOKEN is invalid, falling back to token file:', (err as Error).message);
      } else {
        throw new Error(`Failed to verify env var credentials: ${(err as Error).message}`);
      }
    }
  }

  // Fall back to token.json file
  let content: Buffer;
  try {
    content = await fs.readFile(TOKEN_PATH);
  } catch {
    return null;
  }

  try {
    const credentials = JSON.parse(content.toString());
    const { client_secret, client_id, redirect_uris } = await loadClientSecrets();
    const client = new google.auth.OAuth2(client_id, client_secret, redirect_uris?.[0]);
    client.setCredentials(credentials);
    installTokenRefreshListener(client);

    // Verify token is still valid
    const { credentials: refreshed } = await client.refreshAccessToken();
    // Persist if Google rotated the refresh token
    if (refreshed.refresh_token && refreshed.refresh_token !== credentials.refresh_token) {
      console.error('Refresh token was rotated during startup verification, saving...');
      await saveCredentials(client);
    }
    if (!(await isClientForRequiredAccount(client))) {
      // Treat wrong-account cached token as invalid so the interactive flow runs and overwrites it
      return null;
    }
    return client;
  } catch (err: unknown) {
    const msg = ((err as Error).message || '').toLowerCase();
    if (isRecoverableAuthError(msg)) {
      console.error('Saved credentials are invalid, will re-authenticate:', (err as Error).message);
      return null;
    }
    throw new Error(`Failed to verify saved credentials: ${(err as Error).message}`);
  }
}

async function loadClientSecrets() {
  // Prefer env vars over credentials.json file
  if (process.env.GOOGLE_CLIENT_ID && process.env.GOOGLE_CLIENT_SECRET) {
    return {
      client_id: process.env.GOOGLE_CLIENT_ID,
      client_secret: process.env.GOOGLE_CLIENT_SECRET,
      redirect_uris: [process.env.GOOGLE_REDIRECT_URI || 'http://localhost:3000/'],
      client_type: 'installed' as const,
    };
  }

  // Fall back to credentials.json file
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

export async function saveCredentials(client: OAuth2Client): Promise<void> {
  const { client_secret, client_id } = await loadClientSecrets();
  const payload = JSON.stringify({
    type: 'authorized_user',
    client_id: client_id,
    client_secret: client_secret,
    refresh_token: client.credentials.refresh_token,
    access_token: client.credentials.access_token,
    expiry_date: client.credentials.expiry_date,
    token_type: client.credentials.token_type,
    scope: client.credentials.scope,
  });
  await fs.writeFile(TOKEN_PATH, payload, { mode: 0o600 });
  console.error('Token stored to', TOKEN_PATH);

  // Hint for env var users — point at the field; never interpolate the live refresh token,
  // and never print an assignment whose value is a placeholder someone could paste verbatim
  if (process.env.GOOGLE_CLIENT_ID) {
    console.error(
      `To use env vars instead of token.json, set GOOGLE_REFRESH_TOKEN to the refresh_token field in ${TOKEN_PATH}`
    );
  }
}

export function listenForOAuthCode(
  expectedState: string,
  port: number = 3000
): { code: Promise<string>; listening: Promise<number>; close: () => void } {
  let closeFn = () => {};
  let listeningSettled = false;
  let resolveListening: (boundPort: number) => void = () => {};
  let rejectListening: (err: Error) => void = () => {};
  const listening = new Promise<number>((resolve, reject) => {
    resolveListening = resolve;
    rejectListening = reject;
  });

  const code = new Promise<string>((resolve, reject) => {
    let settled = false;
    let timeoutHandle: NodeJS.Timeout;
    const server = http.createServer((req, res) => {
      try {
        const reqUrl = new URL(req.url || '', `http://localhost:${port}`);
        const receivedCode = reqUrl.searchParams.get('code');
        const error = reqUrl.searchParams.get('error');

        if (error) {
          res.writeHead(400, { 'Content-Type': 'text/html' });
          res.end(`<html><body><h1>Authorization Failed</h1><p>Error: ${error}</p><p>You can close this window.</p></body></html>`);
          settle('reject', new Error(`Authorization error: ${error}`));
          return;
        }

        if (receivedCode) {
          const receivedState = reqUrl.searchParams.get('state');
          if (receivedState !== expectedState) {
            console.error('Rejected OAuth callback: state parameter did not match this flow. Still waiting...');
            res.writeHead(400, { 'Content-Type': 'text/html' });
            res.end('<html><body><h1>Invalid OAuth state</h1><p>You can close this window.</p></body></html>');
            return;
          }
          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end('<html><body><h1>Authorization Successful!</h1><p>You can close this window and return to the terminal.</p></body></html>');
          settle('resolve', receivedCode);
          return;
        }

        res.writeHead(400, { 'Content-Type': 'text/html' });
        res.end('<html><body><h1>No authorization code received</h1></body></html>');
      } catch (err) {
        settle('reject', err as Error);
      }
    });

    function settle(kind: 'resolve' | 'reject', value: string | Error) {
      if (settled) return;
      settled = true;
      clearTimeout(timeoutHandle);
      server.close();
      if (kind === 'resolve') {
        resolve(value as string);
        return;
      }
      const err = value instanceof Error ? value : new Error(String(value));
      if (!listeningSettled) {
        listeningSettled = true;
        rejectListening(err);
      }
      reject(err);
    }

    timeoutHandle = setTimeout(() => {
      settle('reject', new Error('Authentication timed out after 5 minutes'));
    }, 5 * 60 * 1000);

    server.listen(port, () => {
      const address = server.address();
      const boundPort = typeof address === 'object' && address !== null ? address.port : port;
      console.error(`Local server listening on port ${boundPort}...`);
      if (!listeningSettled) {
        listeningSettled = true;
        resolveListening(boundPort);
      }
    });

    server.on('error', (err: NodeJS.ErrnoException) => {
      if (err.code === 'EADDRINUSE') {
        settle('reject', new Error(`Port ${port} is already in use. Please close any application using it and try again.`));
      } else {
        settle('reject', err);
      }
    });

    closeFn = () => {
      settle('reject', new Error('OAuth callback listener closed'));
    };
  });

  return { code, listening, close: () => closeFn() };
}

async function authenticate(): Promise<OAuth2Client> {
  const { client_secret, client_id, redirect_uris, client_type } = await loadClientSecrets();

  // Use localhost redirect for installed apps (OOB flow deprecated by Google)
  // For web clients, use the configured redirect URI
  const PORT = 3000;
  const redirectUri = client_type === 'web' ? redirect_uris[0] : `http://localhost:${PORT}`;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirectUri);

  const authUrlOptions: Parameters<typeof oAuth2Client.generateAuthUrl>[0] = {
    access_type: 'offline',
    scope: SCOPES.join(' '),
    prompt: 'consent', // Force consent screen to ensure we get refresh_token
  };
  if (process.env.REQUIRED_ACCOUNT_EMAIL) {
    // Pre-select the required account in Google's picker so the wrong-account pathway is closed off
    authUrlOptions.login_hint = process.env.REQUIRED_ACCOUNT_EMAIL;
  }
  const state = randomBytes(32).toString('hex');
  authUrlOptions.state = state;
  const authorizeUrl = oAuth2Client.generateAuthUrl(authUrlOptions);

  console.error('\n=== Google OAuth Authentication ===');
  console.error('Opening browser for authentication...');
  console.error('If browser does not open, visit this URL manually:');
  console.error(authorizeUrl);
  console.error('\nWaiting for authorization...\n');

  // For installed apps, start local server to capture the code
  if (client_type !== 'web') {
    const waiter = listenForOAuthCode(state, PORT);
    void waiter.listening.then(() => {
      import('child_process').then(({ exec }) => {
        const platform = process.platform;
        const openCommand = platform === 'darwin' ? 'open' : platform === 'win32' ? 'start' : 'xdg-open';
        exec(`${openCommand} "${authorizeUrl}"`, (err) => {
          if (err) {
            console.error('Could not open browser automatically. Please open the URL manually.');
          }
        });
      });
    }, () => {});
    const code = await waiter.code;

    try {
      const { tokens } = await oAuth2Client.getToken(code);
      oAuth2Client.setCredentials(tokens);
      installTokenRefreshListener(oAuth2Client);
      // Validate before persisting so a wrong-account OAuth never reaches disk
      await enforceRequiredAccount(oAuth2Client);
      if (tokens.refresh_token) {
        await saveCredentials(oAuth2Client);
      } else {
        console.error("Did not receive refresh token. Token might expire.");
      }
      console.error('Authentication successful!');
      return oAuth2Client;
    } catch (err) {
      console.error('Error retrieving access token');
      throw new Error(`Authentication failed: ${(err as Error).message}`);
    }
  } else {
    // For web clients, the redirect will happen externally
    throw new Error('Web client authentication requires external redirect handling');
  }
}

async function enforceRequiredAccount(client: OAuth2Client | JWT): Promise<void> {
  const required = process.env.REQUIRED_ACCOUNT_EMAIL;
  if (!required) return;

  const actual = await getAuthenticatedEmail(client);
  if (!accountMatchesRequired(actual)) {
    throw new Error(
      `Account mismatch: expected "${required}" but authenticated as "${actual}". ` +
      `Re-run authentication and pick "${required}" in Google's account picker.`
    );
  }
  console.error(`Account verified: ${actual}`);
}

// --- MODIFIED: The Main Exported Function ---
// This function now acts as a router. It checks for the environment
// variable and decides which authentication method to use.
export async function authorize(): Promise<OAuth2Client | JWT> {
  let client: OAuth2Client | JWT;

  // Check if the Service Account environment variable is set.
  if (process.env.SERVICE_ACCOUNT_PATH) {
    console.error('Service account path detected. Attempting service account authentication...');
    client = await authorizeWithServiceAccount();
  } else {
    // If not, execute the original OAuth 2.0 flow exactly as it was.
    console.error('No service account path detected. Falling back to standard OAuth 2.0 flow...');
    let oauthClient = await loadSavedCredentialsIfExist();
    if (oauthClient) {
      console.error('Using saved credentials.');
      client = oauthClient;
    } else {
      console.error('Starting authentication flow...');
      client = await authenticate();
    }
  }

  await enforceRequiredAccount(client);
  return client;
}
// --- END OF MODIFIED: The Main Exported Function ---
