import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { afterEach, describe, expect, it, vi } from 'vitest';

describe('/api/web-document', () => {
  const tempDirs: string[] = [];

  afterEach(() => {
    vi.unstubAllEnvs();
    vi.resetModules();
    for (const tempDir of tempDirs.splice(0)) {
      rmSync(tempDir, { recursive: true, force: true });
    }
  });

  async function importRoute() {
    const tempDir = mkdtempSync(join(tmpdir(), 'word-sync-demo-api-'));
    tempDirs.push(tempDir);
    vi.stubEnv('WEB_DOCUMENT_DB_PATH', join(tempDir, 'web-document.sqlite'));

    return import('./route');
  }

  it('reads the versioned shared Webドキュメント', async () => {
    const { GET } = await importRoute();

    const response = await GET();

    expect(response.status).toBe(200);
    await expect(response.json()).resolves.toEqual({
      id: 'shared',
      markdown: expect.stringContaining('Webドキュメント'),
      version: 1,
    });
  });

  it('autosaves Markdown and returns the incremented version', async () => {
    const { GET, PUT } = await importRoute();

    const response = await PUT(
      new Request('http://localhost/api/web-document', {
        method: 'PUT',
        body: JSON.stringify({
          markdown: '## Autosaved\n\nAPIから保存されたMarkdown互換内容。',
        }),
      }),
    );

    expect(response.status).toBe(200);
    await expect(response.json()).resolves.toEqual({
      id: 'shared',
      markdown: '## Autosaved\n\nAPIから保存されたMarkdown互換内容。',
      version: 2,
    });

    await expect((await GET()).json()).resolves.toEqual({
      id: 'shared',
      markdown: '## Autosaved\n\nAPIから保存されたMarkdown互換内容。',
      version: 2,
    });
  });
});
