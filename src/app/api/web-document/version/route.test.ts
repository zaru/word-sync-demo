import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { afterEach, describe, expect, it, vi } from 'vitest';

describe('/api/web-document/version', () => {
  const tempDirs: string[] = [];

  afterEach(() => {
    vi.unstubAllEnvs();
    vi.resetModules();
    for (const tempDir of tempDirs.splice(0)) {
      rmSync(tempDir, { recursive: true, force: true });
    }
  });

  async function importRoutes() {
    const tempDir = mkdtempSync(join(tmpdir(), 'word-sync-demo-version-'));
    tempDirs.push(tempDir);
    vi.stubEnv('WEB_DOCUMENT_DB_PATH', join(tempDir, 'web-document.sqlite'));

    const documentRoute = await import('../route');
    const versionRoute = await import('./route');

    return { documentRoute, versionRoute };
  }

  it('returns only the current Webドキュメント version for polling', async () => {
    const { documentRoute, versionRoute } = await importRoutes();

    await documentRoute.PUT(
      new Request('http://localhost/api/web-document', {
        method: 'PUT',
        body: JSON.stringify({ markdown: '# New version' }),
      }),
    );

    const response = await versionRoute.GET();

    expect(response.status).toBe(200);
    await expect(response.json()).resolves.toEqual({ version: 2 });
  });
});
