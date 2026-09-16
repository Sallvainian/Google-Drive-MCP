import { OAuth2Client } from 'google-auth-library';

export async function authorize() {
  return new OAuth2Client('capture-stub-client-id', 'capture-stub-client-secret');
}
