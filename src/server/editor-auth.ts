import { randomUUID } from "node:crypto";
import { join } from "node:path";

import "server-only";

import {
	createEditorAuthHandlers,
	type MicrosoftAuthBoundary,
} from "./editor-auth-routes";
import { EditorAuthStore, type EditorSession } from "./editor-auth-store";
import { createMicrosoftAuthBoundary } from "./microsoft-auth";

let cachedStore: { databasePath: string; store: EditorAuthStore } | undefined;

export function getEditorAuthStore(): EditorAuthStore {
	const databasePath =
		process.env.EDITOR_AUTH_DB_PATH ??
		join(process.cwd(), ".data", "editor-auth.sqlite");

	if (cachedStore?.databasePath !== databasePath) {
		cachedStore?.store.close();
		cachedStore = {
			databasePath,
			store: new EditorAuthStore({ databasePath }),
		};
	}

	return cachedStore.store;
}

export function getEditorAuthHandlers() {
	return createEditorAuthHandlers({
		appBaseUrl: process.env.APP_BASE_URL ?? "http://localhost:3000",
		createSessionId: randomUUID,
		createState: randomUUID,
		microsoftAuth: createLazyMicrosoftAuthBoundary(),
		store: getEditorAuthStore(),
	});
}

export function readEditorSessionFromCookieValue(
	sessionId: string | undefined,
): EditorSession | undefined {
	if (!sessionId) {
		return undefined;
	}

	return getEditorAuthStore().readSession(sessionId);
}

function createLazyMicrosoftAuthBoundary(): MicrosoftAuthBoundary {
	let boundary: MicrosoftAuthBoundary | undefined;

	function getBoundary() {
		boundary ??= createMicrosoftAuthBoundary({
			clientId: requireEnv("MICROSOFT_CLIENT_ID"),
			clientSecret: requireEnv("MICROSOFT_CLIENT_SECRET"),
		});

		return boundary;
	}

	return {
		createAuthorizationUrl(input) {
			return getBoundary().createAuthorizationUrl(input);
		},
		completeAuthorizationCode(input) {
			return getBoundary().completeAuthorizationCode(input);
		},
	};
}

function requireEnv(name: string): string {
	const value = process.env[name];

	if (!value) {
		throw new Error(`${name} is required for Microsoft sign-in.`);
	}

	return value;
}
