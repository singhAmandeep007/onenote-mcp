import { DeviceCodeCredential } from '@azure/identity';
import { CLIENT_ID, TENANT_ID, SCOPES } from './config.js';
import { saveToken, isLikelyJwt } from './token-store.js';
import { log } from './logger.js';

export interface AuthResult {
  token: string;
  /** True for work/school (JWT) tokens; false for personal-account compact tokens. */
  isJwt: boolean;
}

/**
 * Run the Microsoft device-code flow and cache the resulting access token.
 *
 * @param prompt Callback invoked with the sign-in message (URL + code). Defaults to
 *               logging to stderr.
 */
export async function authenticateWithDeviceCode(
  prompt: (message: string) => void = (message) => log(message),
): Promise<AuthResult> {
  const credential = new DeviceCodeCredential({
    clientId: CLIENT_ID,
    tenantId: TENANT_ID,
    userPromptCallback: (info) => prompt(info.message),
  });

  const response = await credential.getToken(SCOPES);
  if (!response?.token) {
    throw new Error('Device-code authentication did not return an access token.');
  }

  saveToken(response.token);
  return { token: response.token, isJwt: isLikelyJwt(response.token) };
}
