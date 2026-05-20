import { ConfidentialClientApplication } from '@azure/msal-node';

import type { MicrosoftAuthBoundary } from './editor-auth-routes';

const microsoftAuthority = 'https://login.microsoftonline.com/common';

export function createMicrosoftAuthBoundary(options: {
  clientId: string;
  clientSecret: string;
}): MicrosoftAuthBoundary {
  const client = new ConfidentialClientApplication({
    auth: {
      authority: microsoftAuthority,
      clientId: options.clientId,
      clientSecret: options.clientSecret,
    },
  });

  return {
    createAuthorizationUrl(input) {
      return client.getAuthCodeUrl({
        redirectUri: input.redirectUri,
        scopes: input.scopes,
        state: input.state,
      });
    },
    async completeAuthorizationCode(input) {
      const result = await client.acquireTokenByCode({
        code: input.code,
        redirectUri: input.redirectUri,
        scopes: input.scopes,
      });

      if (!result?.account) {
        throw new Error('Microsoft sign-in did not return an account.');
      }

      return {
        editorId: result.account.homeAccountId,
        displayName: result.account.name,
        username: result.account.username,
        tokenCache: client.getTokenCache().serialize(),
      };
    },
  };
}
