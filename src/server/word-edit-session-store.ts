import { mkdirSync } from "node:fs";
import { dirname } from "node:path";
import { DatabaseSync } from "node:sqlite";

export type WordEditSession = {
	sessionId: string;
	webDocumentId: "shared";
	webDocumentVersion: number;
	editorId: string;
	driveItemId: string;
	workingCopyFileName: string;
	oneDriveWebUrl: string;
};

type WordEditSessionRow = {
	session_id: string;
	web_document_id: string;
	web_document_version: number;
	editor_id: string;
	drive_item_id: string;
	working_copy_file_name: string;
	one_drive_web_url: string;
};

type WordEditSessionStateRow = {
	discarded_at: string | null;
	finished_at: string | null;
	import_error_at: string | null;
};

export type WordEditSessionState =
	| "active"
	| "discarded"
	| "finished"
	| "importError";

export function createWordEditSessionStore(options: { databasePath: string }) {
	mkdirSync(dirname(options.databasePath), { recursive: true });
	const database = new DatabaseSync(options.databasePath);

	database.exec(`
    CREATE TABLE IF NOT EXISTS word_edit_sessions (
      session_id TEXT PRIMARY KEY,
      web_document_id TEXT NOT NULL,
      web_document_version INTEGER NOT NULL,
      editor_id TEXT NOT NULL,
      drive_item_id TEXT NOT NULL,
      working_copy_file_name TEXT NOT NULL,
      one_drive_web_url TEXT NOT NULL,
      discarded_at TEXT,
      finished_at TEXT,
      import_error_at TEXT,
      created_at TEXT NOT NULL
    ) STRICT;
  `);
	ensureColumn(
		database,
		"discarded_at",
		"ALTER TABLE word_edit_sessions ADD COLUMN discarded_at TEXT;",
	);
	ensureColumn(
		database,
		"finished_at",
		"ALTER TABLE word_edit_sessions ADD COLUMN finished_at TEXT;",
	);
	ensureColumn(
		database,
		"import_error_at",
		"ALTER TABLE word_edit_sessions ADD COLUMN import_error_at TEXT;",
	);

	return {
		saveStartedSession(input: WordEditSession): WordEditSession {
			database
				.prepare(
					`INSERT INTO word_edit_sessions (
             session_id,
             web_document_id,
             web_document_version,
             editor_id,
             drive_item_id,
             working_copy_file_name,
             one_drive_web_url,
             created_at
           ) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))`,
				)
				.run(
					input.sessionId,
					input.webDocumentId,
					input.webDocumentVersion,
					input.editorId,
					input.driveItemId,
					input.workingCopyFileName,
					input.oneDriveWebUrl,
				);

			return input;
		},

		readSession(sessionId: string): WordEditSession | undefined {
			const row = database
				.prepare(
					`SELECT
             session_id,
             web_document_id,
             web_document_version,
             editor_id,
             drive_item_id,
             working_copy_file_name,
             one_drive_web_url
           FROM word_edit_sessions
           WHERE session_id = ?`,
				)
				.get(sessionId);

			if (!isWordEditSessionRow(row)) {
				return undefined;
			}

			return {
				sessionId: row.session_id,
				webDocumentId: "shared",
				webDocumentVersion: row.web_document_version,
				editorId: row.editor_id,
				driveItemId: row.drive_item_id,
				workingCopyFileName: row.working_copy_file_name,
				oneDriveWebUrl: row.one_drive_web_url,
			};
		},

		finishSession(sessionId: string): void {
			const result = database
				.prepare(
					"UPDATE word_edit_sessions SET finished_at = datetime('now'), import_error_at = NULL WHERE session_id = ? AND discarded_at IS NULL AND finished_at IS NULL",
				)
				.run(sessionId);

			if (result.changes !== 1) {
				throw new Error(
					"Word編集セッション could not transition to セッション終了.",
				);
			}
		},

		markSessionImportError(sessionId: string): void {
			const result = database
				.prepare(
					"UPDATE word_edit_sessions SET import_error_at = datetime('now') WHERE session_id = ? AND discarded_at IS NULL AND finished_at IS NULL",
				)
				.run(sessionId);

			if (result.changes !== 1) {
				throw new Error(
					"Word編集セッション could not transition to 取り込みエラー.",
				);
			}
		},

		discardSession(sessionId: string): void {
			const result = database
				.prepare(
					"UPDATE word_edit_sessions SET discarded_at = datetime('now') WHERE session_id = ? AND discarded_at IS NULL AND finished_at IS NULL AND import_error_at IS NOT NULL",
				)
				.run(sessionId);

			if (result.changes !== 1) {
				throw new Error("Word編集セッション could not transition to 破棄.");
			}
		},

		readSessionState(sessionId: string): WordEditSessionState | undefined {
			const row = database
				.prepare(
					"SELECT discarded_at, finished_at, import_error_at FROM word_edit_sessions WHERE session_id = ?",
				)
				.get(sessionId);

			if (!isWordEditSessionStateRow(row)) {
				return undefined;
			}

			if (row.finished_at !== null) {
				return "finished";
			}

			if (row.discarded_at !== null) {
				return "discarded";
			}

			return row.import_error_at === null ? "active" : "importError";
		},

		close(): void {
			database.close();
		},
	};
}

export type WordEditSessionStore = ReturnType<
	typeof createWordEditSessionStore
>;

function ensureColumn(
	database: DatabaseSync,
	columnName: string,
	addColumnSql: string,
): void {
	const columns = database
		.prepare("PRAGMA table_info(word_edit_sessions)")
		.all()
		.flatMap((row) =>
			typeof row === "object" &&
			row !== null &&
			typeof (row as Record<string, unknown>).name === "string"
				? [(row as Record<string, string>).name]
				: [],
		);

	if (!columns.includes(columnName)) {
		database.exec(addColumnSql);
	}
}

function isWordEditSessionRow(row: unknown): row is WordEditSessionRow {
	if (typeof row !== "object" || row === null) {
		return false;
	}

	const candidate = row as Record<string, unknown>;

	return (
		typeof candidate.session_id === "string" &&
		candidate.web_document_id === "shared" &&
		typeof candidate.web_document_version === "number" &&
		typeof candidate.editor_id === "string" &&
		typeof candidate.drive_item_id === "string" &&
		typeof candidate.working_copy_file_name === "string" &&
		typeof candidate.one_drive_web_url === "string"
	);
}

function isWordEditSessionStateRow(
	row: unknown,
): row is WordEditSessionStateRow {
	if (typeof row !== "object" || row === null) {
		return false;
	}

	const finishedAt = (row as Record<string, unknown>).finished_at;
	const importErrorAt = (row as Record<string, unknown>).import_error_at;
	const discardedAt = (row as Record<string, unknown>).discarded_at;

	return (
		(finishedAt === null || typeof finishedAt === "string") &&
		(importErrorAt === null || typeof importErrorAt === "string") &&
		(discardedAt === null || typeof discardedAt === "string")
	);
}
