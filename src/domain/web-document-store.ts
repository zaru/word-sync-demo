import { mkdirSync } from 'node:fs';
import { dirname } from 'node:path';
import { DatabaseSync } from 'node:sqlite';

export type WebDocument = {
  id: 'shared';
  markdown: string;
  version: number;
};

type WebDocumentStoreOptions = {
  databasePath: string;
  seedMarkdown: string;
};

type WebDocumentRow = {
  id: string;
  markdown: string;
  version: number;
};

export class WebDocumentStore {
  public readonly databasePath: string;

  private readonly database: DatabaseSync;

  constructor(options: WebDocumentStoreOptions) {
    this.databasePath = options.databasePath;
    mkdirSync(dirname(options.databasePath), { recursive: true });
    this.database = new DatabaseSync(options.databasePath);
    this.migrate(options.seedMarkdown);
  }

  loadSharedDocument(): WebDocument {
    const row = this.database
      .prepare('SELECT id, markdown, version FROM web_documents WHERE id = ?')
      .get('shared');

    if (!isWebDocumentRow(row)) {
      throw new Error('Shared Webドキュメント was not initialized.');
    }

    return {
      id: 'shared',
      markdown: row.markdown,
      version: row.version,
    };
  }

  saveMarkdown(markdown: string): WebDocument {
    const row = this.database
      .prepare(
        'UPDATE web_documents SET markdown = ?, version = version + 1 WHERE id = ? RETURNING id, markdown, version',
      )
      .get(markdown, 'shared');

    if (!isWebDocumentRow(row)) {
      throw new Error('Shared Webドキュメント could not be saved.');
    }

    return {
      id: 'shared',
      markdown: row.markdown,
      version: row.version,
    };
  }

  readVersion(): number {
    const row = this.database.prepare('SELECT version FROM web_documents WHERE id = ?').get('shared');

    if (!isVersionRow(row)) {
      throw new Error('Shared Webドキュメント version could not be read.');
    }

    return row.version;
  }

  close(): void {
    this.database.close();
  }

  private migrate(seedMarkdown: string): void {
    this.database.exec(`
      CREATE TABLE IF NOT EXISTS web_documents (
        id TEXT PRIMARY KEY,
        markdown TEXT NOT NULL,
        version INTEGER NOT NULL
      ) STRICT;
    `);

    this.database
      .prepare('INSERT OR IGNORE INTO web_documents (id, markdown, version) VALUES (?, ?, ?)')
      .run('shared', seedMarkdown, 1);
  }
}

function isWebDocumentRow(row: unknown): row is WebDocumentRow {
  if (typeof row !== 'object' || row === null) {
    return false;
  }

  const candidate = row as Record<string, unknown>;

  return (
    candidate.id === 'shared' &&
    typeof candidate.markdown === 'string' &&
    typeof candidate.version === 'number'
  );
}

function isVersionRow(row: unknown): row is Pick<WebDocumentRow, 'version'> {
  if (typeof row !== 'object' || row === null) {
    return false;
  }

  return typeof (row as Record<string, unknown>).version === 'number';
}
