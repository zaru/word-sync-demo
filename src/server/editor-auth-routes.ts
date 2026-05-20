import type { EditorAuthStore, EditorProfile } from './editor-auth-store';

export const editorAuthScopes = [
  'openid',
  'profile',
  'email',
  'offline_access',
  'Files.ReadWrite.AppFolder',
];

export const editorSessionCookieName = 'word_sync_editor_session';
const oauthStateCookieName = 'word_sync_oauth_state';

export type MicrosoftAuthBoundary = {
  createAuthorizationUrl(input: {
    redirectUri: string;
    scopes: string[];
    state: string;
  }): Promise<string>;
  completeAuthorizationCode(input: {
    code: string;
    redirectUri: string;
    scopes: string[];
  }): Promise<{
    editorId: string;
    displayName?: string;
    username?: string;
    tokenCache: string;
  }>;
};

export function createEditorAuthHandlers(options: {
  appBaseUrl: string;
  createSessionId: () => string;
  createState: () => string;
  microsoftAuth: MicrosoftAuthBoundary;
  store: EditorAuthStore;
}) {
  const redirectUri = `${options.appBaseUrl}/auth/callback`;

  return {
    async login(): Promise<Response> {
      const state = options.createState();
      const authorizationUrl = await options.microsoftAuth.createAuthorizationUrl({
        redirectUri,
        scopes: editorAuthScopes,
        state,
      });

      return redirectResponse(authorizationUrl, [
        serializeCookie(oauthStateCookieName, state, {
          httpOnly: true,
          maxAgeSeconds: 600,
          sameSite: 'Lax',
        }),
      ]);
    },

    async callback(request: Request): Promise<Response> {
      const url = new URL(request.url);
      const code = url.searchParams.get('code');
      const state = url.searchParams.get('state');
      const expectedState = readCookie(request, oauthStateCookieName);

      if (!code || !state || state !== expectedState) {
        return Response.json({ error: 'Invalid Microsoft sign-in callback' }, { status: 400 });
      }

      const signedInEditor = await options.microsoftAuth.completeAuthorizationCode({
        code,
        redirectUri,
        scopes: editorAuthScopes,
      });
      const session = options.store.saveSignedInEditor({
        editor: {
          id: signedInEditor.editorId,
          displayName: signedInEditor.displayName,
          username: signedInEditor.username,
        },
        sessionId: options.createSessionId(),
        tokenCache: signedInEditor.tokenCache,
      });

      return redirectResponse(`${options.appBaseUrl}/`, [
        serializeCookie(editorSessionCookieName, session.sessionId, {
          httpOnly: true,
          maxAgeSeconds: 60 * 60 * 24 * 14,
          sameSite: 'Lax',
        }),
        expireCookie(oauthStateCookieName),
      ]);
    },

    async session(request: Request): Promise<Response> {
      const sessionId = readCookie(request, editorSessionCookieName);
      const session = sessionId ? options.store.readSession(sessionId) : undefined;

      if (!session) {
        return Response.json({ signedIn: false });
      }

      return Response.json({
        signedIn: true,
        editor: session.editor,
      });
    },

    async logout(request: Request): Promise<Response> {
      const sessionId = readCookie(request, editorSessionCookieName);

      if (sessionId) {
        options.store.clearSession(sessionId);
      }

      return redirectResponse(`${options.appBaseUrl}/`, [expireCookie(editorSessionCookieName)]);
    },
  };
}

function redirectResponse(location: string, cookies: string[]): Response {
  const headers = new Headers({ location });

  for (const cookie of cookies) {
    headers.append('set-cookie', cookie);
  }

  return new Response(null, {
    headers,
    status: 302,
  });
}

function readCookie(request: Request, name: string): string | undefined {
  const cookieHeader = request.headers.get('cookie');

  if (!cookieHeader) {
    return undefined;
  }

  for (const cookie of cookieHeader.split(/; */)) {
    const [rawName, ...rawValueParts] = cookie.split('=');

    if (rawName === name) {
      return decodeURIComponent(rawValueParts.join('='));
    }
  }

  return undefined;
}

function serializeCookie(
  name: string,
  value: string,
  options: { httpOnly: boolean; maxAgeSeconds: number; sameSite: 'Lax' },
): string {
  return [
    `${name}=${encodeURIComponent(value)}`,
    'Path=/',
    `Max-Age=${options.maxAgeSeconds}`,
    options.httpOnly ? 'HttpOnly' : undefined,
    `SameSite=${options.sameSite}`,
  ]
    .filter(Boolean)
    .join('; ');
}

function expireCookie(name: string): string {
  return `${name}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax`;
}
