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
      created_at TEXT NOT NULL
    ) STRICT;
  `);

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

		close(): void {
			database.close();
		},
	};
}

export type WordEditSessionStore = ReturnType<
	typeof createWordEditSessionStore
>;

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
