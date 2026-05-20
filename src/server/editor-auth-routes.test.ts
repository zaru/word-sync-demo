import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { afterEach, describe, expect, it } from "vitest";
import {
	createEditorAuthHandlers,
	type MicrosoftAuthBoundary,
} from "./editor-auth-routes";
import { EditorAuthStore } from "./editor-auth-store";

describe("Microsoft sign-in routes", () => {
	const tempDirs: string[] = [];

	afterEach(() => {
		for (const tempDir of tempDirs.splice(0)) {
			rmSync(tempDir, { recursive: true, force: true });
		}
	});

	function createStore() {
		const tempDir = mkdtempSync(join(tmpdir(), "word-sync-demo-auth-"));
		tempDirs.push(tempDir);

		return new EditorAuthStore({
			databasePath: join(tempDir, "auth.sqlite"),
		});
	}

	it("signs an 編集者 in after Microsoft callback without exposing raw token cache to browser code", async () => {
		const tokenCache =
			'{"RefreshToken":{"secret":"browser must not receive this"}}';
		const store = createStore();
		const microsoftAuth: MicrosoftAuthBoundary = {
			async createAuthorizationUrl({ state }) {
				return `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?state=${state}`;
			},
			async completeAuthorizationCode() {
				return {
					editorId: "microsoft-account-1",
					displayName: "編集者 A",
					username: "editor@example.com",
					tokenCache,
				};
			},
		};
		const handlers = createEditorAuthHandlers({
			appBaseUrl: "http://localhost",
			createSessionId: () => "session-1",
			createState: () => "state-1",
			microsoftAuth,
			store,
		});

		const loginResponse = await handlers.login(
			new Request("http://localhost/auth/login"),
		);

		expect(loginResponse.status).toBe(302);
		expect(loginResponse.headers.get("location")).toBe(
			"https://login.microsoftonline.com/common/oauth2/v2.0/authorize?state=state-1",
		);

		const callbackResponse = await handlers.callback(
			new Request(
				"http://localhost/auth/callback?code=auth-code&state=state-1",
				{
					headers: {
						cookie: loginResponse.headers.get("set-cookie") ?? "",
					},
				},
			),
		);

		expect(callbackResponse.status).toBe(302);
		expect(callbackResponse.headers.get("location")).toBe("http://localhost/");

		const sessionResponse = await handlers.session(
			new Request("http://localhost/api/session", {
				headers: {
					cookie: callbackResponse.headers.get("set-cookie") ?? "",
				},
			}),
		);

		const browserSession = await sessionResponse.json();

		expect(browserSession).toEqual({
			signedIn: true,
			editor: {
				id: "microsoft-account-1",
				displayName: "編集者 A",
				username: "editor@example.com",
			},
		});
		expect(JSON.stringify(browserSession)).not.toContain(
			"browser must not receive this",
		);
		expect(store.readTokenCache("microsoft-account-1")).toBe(tokenCache);
	});

	it("logs out the signed-in 編集者 and clears the browser session", async () => {
		const store = createStore();
		store.saveSignedInEditor({
			editor: {
				id: "microsoft-account-1",
				displayName: "編集者 A",
				username: "editor@example.com",
			},
			sessionId: "session-1",
			tokenCache: '{"RefreshToken":{"cached":true}}',
		});
		const handlers = createEditorAuthHandlers({
			appBaseUrl: "http://localhost",
			createSessionId: () => "unused",
			createState: () => "unused",
			microsoftAuth: {
				async createAuthorizationUrl() {
					return "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
				},
				async completeAuthorizationCode() {
					throw new Error("not used");
				},
			},
			store,
		});

		const logoutResponse = await handlers.logout(
			new Request("http://localhost/auth/logout", {
				method: "POST",
				headers: {
					cookie: "word_sync_editor_session=session-1",
				},
			}),
		);

		expect(logoutResponse.status).toBe(302);
		expect(logoutResponse.headers.get("location")).toBe("http://localhost/");
		expect(logoutResponse.headers.get("set-cookie")).toContain(
			"word_sync_editor_session=",
		);
		expect(logoutResponse.headers.get("set-cookie")).toContain("Max-Age=0");

		const sessionResponse = await handlers.session(
			new Request("http://localhost/api/session", {
				headers: {
					cookie: "word_sync_editor_session=session-1",
				},
			}),
		);

		await expect(sessionResponse.json()).resolves.toEqual({ signedIn: false });
		expect(store.readTokenCache("microsoft-account-1")).toBe(
			'{"RefreshToken":{"cached":true}}',
		);
	});
});
