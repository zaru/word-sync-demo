import { mkdirSync } from "node:fs";
import { dirname } from "node:path";
import { DatabaseSync } from "node:sqlite";

export type EditorProfile = {
	id: string;
	displayName?: string;
	username?: string;
};

export type EditorSession = {
	sessionId: string;
	editor: EditorProfile;
};

type EditorAuthStoreOptions = {
	databasePath: string;
};

type SessionRow = {
	session_id: string;
	editor_id: string;
	display_name: string | null;
	username: string | null;
};

export class EditorAuthStore {
	public readonly databasePath: string;

	private readonly database: DatabaseSync;

	constructor(options: EditorAuthStoreOptions) {
		this.databasePath = options.databasePath;
		mkdirSync(dirname(options.databasePath), { recursive: true });
		this.database = new DatabaseSync(options.databasePath);
		this.migrate();
	}

	saveSignedInEditor(input: {
		editor: EditorProfile;
		sessionId: string;
		tokenCache: string;
	}): EditorSession {
		this.database
			.prepare(
				`INSERT INTO editor_token_caches (editor_id, display_name, username, token_cache, updated_at)
         VALUES (?, ?, ?, ?, datetime('now'))
         ON CONFLICT(editor_id) DO UPDATE SET
           display_name = excluded.display_name,
           username = excluded.username,
           token_cache = excluded.token_cache,
           updated_at = excluded.updated_at`,
			)
			.run(
				input.editor.id,
				input.editor.displayName ?? null,
				input.editor.username ?? null,
				input.tokenCache,
			);
		this.database
			.prepare(
				`INSERT INTO editor_sessions (session_id, editor_id, created_at)
         VALUES (?, ?, datetime('now'))
         ON CONFLICT(session_id) DO UPDATE SET editor_id = excluded.editor_id`,
			)
			.run(input.sessionId, input.editor.id);

		return {
			sessionId: input.sessionId,
			editor: input.editor,
		};
	}

	readSession(sessionId: string): EditorSession | undefined {
		const row = this.database
			.prepare(
				`SELECT s.session_id, c.editor_id, c.display_name, c.username
         FROM editor_sessions s
         JOIN editor_token_caches c ON c.editor_id = s.editor_id
         WHERE s.session_id = ?`,
			)
			.get(sessionId);

		if (!isSessionRow(row)) {
			return undefined;
		}

		return {
			sessionId: row.session_id,
			editor: {
				id: row.editor_id,
				displayName: row.display_name ?? undefined,
				username: row.username ?? undefined,
			},
		};
	}

	readTokenCache(editorId: string): string | undefined {
		const row = this.database
			.prepare(
				"SELECT token_cache FROM editor_token_caches WHERE editor_id = ?",
			)
			.get(editorId);

		if (!isTokenCacheRow(row)) {
			return undefined;
		}

		return row.token_cache;
	}

	clearSession(sessionId: string): void {
		this.database
			.prepare("DELETE FROM editor_sessions WHERE session_id = ?")
			.run(sessionId);
	}

	close(): void {
		this.database.close();
	}

	private migrate(): void {
		this.database.exec(`
      CREATE TABLE IF NOT EXISTS editor_token_caches (
        editor_id TEXT PRIMARY KEY,
        display_name TEXT,
        username TEXT,
        token_cache TEXT NOT NULL,
        updated_at TEXT NOT NULL
      ) STRICT;

      CREATE TABLE IF NOT EXISTS editor_sessions (
        session_id TEXT PRIMARY KEY,
        editor_id TEXT NOT NULL,
        created_at TEXT NOT NULL,
        FOREIGN KEY (editor_id) REFERENCES editor_token_caches(editor_id)
      ) STRICT;
    `);
	}
}

function isSessionRow(row: unknown): row is SessionRow {
	if (typeof row !== "object" || row === null) {
		return false;
	}

	const candidate = row as Record<string, unknown>;

	return (
		typeof candidate.session_id === "string" &&
		typeof candidate.editor_id === "string" &&
		(typeof candidate.display_name === "string" ||
			candidate.display_name === null) &&
		(typeof candidate.username === "string" || candidate.username === null)
	);
}

function isTokenCacheRow(row: unknown): row is { token_cache: string } {
	if (typeof row !== "object" || row === null) {
		return false;
	}

	return typeof (row as Record<string, unknown>).token_cache === "string";
}
