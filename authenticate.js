import { DeviceCodeCredential } from '@azure/identity';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

// Get current directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path for storing the access token
const tokenFilePath = path.join(__dirname, '.access-token.txt');

// Client ID for Microsoft Graph API access
const clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'; // Microsoft Graph Explorer client ID

// 'common' accepts BOTH personal Microsoft accounts (MSA) and work/school (Azure AD).
// Override with GRAPH_TENANT if you need to lock to a single org.
const tenantId = process.env.GRAPH_TENANT || 'common';

// IMPORTANT: use resource-qualified, NON-".All" delegated scopes.
// Personal Microsoft accounts cannot consent to the ".All" variants, which makes
// the device-code flow return a token Microsoft Graph rejects with 401 (code 40001).
// The non-".All" scopes cover everything this server needs and work for all account types.
const scopes = [
  'https://graph.microsoft.com/Notes.Read',
  'https://graph.microsoft.com/Notes.ReadWrite',
  'https://graph.microsoft.com/User.Read'
];

function isLikelyJwt(token) {
  return typeof token === 'string' && token.split('.').length === 3;
}

async function authenticate() {
  try {
    // Use device code flow
    const credential = new DeviceCodeCredential({
      clientId: clientId,
      tenantId: tenantId,
      userPromptCallback: (info) => {
        // This will show the URL and code to the user
        console.log('\n' + info.message);
      }
    });

    // Get an access token using device code flow
    console.log('Starting authentication...');
    console.log('You will see a URL and code to enter shortly...');

    const tokenResponse = await credential.getToken(scopes);

    // Save the token for future use
    const accessToken = tokenResponse.token;
    fs.writeFileSync(tokenFilePath, JSON.stringify({ token: accessToken }));

    console.log('\nAuthentication successful!');
    console.log('Access token saved to:', tokenFilePath);

    if (isLikelyJwt(accessToken)) {
      console.log('Token format: JWT (work/school account).');
    } else {
      console.log('Token format: compact token (normal for personal Microsoft accounts).');
    }
    console.log("Next: run 'node verify-token.js' to confirm Graph accepts it.");
  } catch (error) {
    console.error('Authentication error:', error);
  }
}

// Run the authentication
authenticate();
