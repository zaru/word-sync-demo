import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { afterEach, describe, expect, it } from 'vitest';

import { EditorAuthStore } from './editor-auth-store';

describe('EditorAuthStore', () => {
  const tempDirs: string[] = [];

  afterEach(() => {
    for (const tempDir of tempDirs.splice(0)) {
      rmSync(tempDir, { recursive: true, force: true });
    }
  });

  it('keeps the MSAL token cache associated with a signed-in 編集者 across store recreation', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'word-sync-demo-auth-store-'));
    tempDirs.push(tempDir);
    const databasePath = join(tempDir, 'auth.sqlite');
    const firstStore = new EditorAuthStore({ databasePath });

    firstStore.saveSignedInEditor({
      editor: {
        id: 'editor-1',
        displayName: '編集者 A',
        username: 'editor@example.com',
      },
      sessionId: 'session-1',
      tokenCache: '{"RefreshToken":{"secret":"persisted"}}',
    });
    firstStore.close();

    const restartedStore = new EditorAuthStore({ databasePath });

    expect(restartedStore.readTokenCache('editor-1')).toBe(
      '{"RefreshToken":{"secret":"persisted"}}',
    );
    expect(restartedStore.readSession('session-1')).toEqual({
      sessionId: 'session-1',
      editor: {
        id: 'editor-1',
        displayName: '編集者 A',
        username: 'editor@example.com',
      },
    });
  });
});
