import { randomUUID } from "node:crypto";
import { join } from "node:path";

import "server-only";

import { getEditorAuthStore } from "./editor-auth";
import { createMicrosoftGraphAppFolderBoundary } from "./microsoft-graph";
import { createPandocMarkdownToDocxConverter } from "./pandoc-markdown-to-docx";
import { getWebDocumentStore } from "./web-document-store";
import { createWordEditSessionHandlers } from "./word-edit-session-routes";
import {
	createWordEditSessionStore,
	type WordEditSessionStore,
} from "./word-edit-session-store";

let cachedStore:
	| { databasePath: string; store: WordEditSessionStore }
	| undefined;

export function getWordEditSessionStore(): WordEditSessionStore {
	const databasePath =
		process.env.WORD_EDIT_SESSION_DB_PATH ??
		join(process.cwd(), ".data", "word-edit-sessions.sqlite");

	if (cachedStore?.databasePath !== databasePath) {
		cachedStore?.store.close();
		cachedStore = {
			databasePath,
			store: createWordEditSessionStore({ databasePath }),
		};
	}

	return cachedStore.store;
}

export function getWordEditSessionHandlers() {
	return createWordEditSessionHandlers({
		converter: createPandocMarkdownToDocxConverter(),
		createSessionId: randomUUID,
		editorAuthStore: getEditorAuthStore(),
		graph: createMicrosoftGraphAppFolderBoundary({
			clientId: requireEnv("MICROSOFT_CLIENT_ID"),
			clientSecret: requireEnv("MICROSOFT_CLIENT_SECRET"),
		}),
		webDocumentStore: getWebDocumentStore(),
		wordEditSessionStore: getWordEditSessionStore(),
	});
}

function requireEnv(name: string): string {
	const value = process.env[name];

	if (!value) {
		throw new Error(`${name} is required for Word編集セッション.`);
	}

	return value;
}
