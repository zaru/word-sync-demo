import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { afterEach, describe, expect, it } from "vitest";
import { editorSessionCookieName } from "./editor-auth-routes";
import { EditorAuthStore } from "./editor-auth-store";
import { createWordEditSessionHandlers } from "./word-edit-session-routes";
import { createWordEditSessionStore } from "./word-edit-session-store";

describe("Word編集セッション routes", () => {
	const tempDirs: string[] = [];

	afterEach(() => {
		for (const tempDir of tempDirs.splice(0)) {
			rmSync(tempDir, { recursive: true, force: true });
		}
	});

	function createStores() {
		const tempDir = mkdtempSync(join(tmpdir(), "word-sync-demo-word-edit-"));
		tempDirs.push(tempDir);

		return {
			editorAuthStore: new EditorAuthStore({
				databasePath: join(tempDir, "auth.sqlite"),
			}),
			wordEditSessionStore: createWordEditSessionStore({
				databasePath: join(tempDir, "word-edit-sessions.sqlite"),
			}),
		};
	}

	it("starts a fresh Word編集セッション with a new OneDrive作業コピー for every request", async () => {
		const { editorAuthStore, wordEditSessionStore } = createStores();
		editorAuthStore.saveSignedInEditor({
			editor: {
				id: "editor-1",
				displayName: "編集者 A",
				username: "editor@example.com",
			},
			sessionId: "browser-session-1",
			tokenCache: '{"RefreshToken":{"cached":true}}',
		});
		const uploadedWorkingCopies: Array<{
			content: Uint8Array;
			fileName: string;
			tokenCache: string;
		}> = [];
		const convertedMarkdown: string[] = [];
		const handlers = createWordEditSessionHandlers({
			converter: {
				async convertMarkdownToDocx(input) {
					convertedMarkdown.push(input.markdown);
					return new TextEncoder().encode(input.markdown);
				},
			},
			createSessionId: createSequentialId("word-session"),
			editorAuthStore,
			graph: {
				async uploadAppFolderWorkingCopy(input) {
					uploadedWorkingCopies.push(input);

					return {
						driveItemId: `drive-item-${uploadedWorkingCopies.length}`,
						webUrl: `https://onedrive.example/${input.fileName}`,
					};
				},
			},
			webDocumentStore: {
				loadSharedDocument() {
					return {
						id: "shared",
						markdown: "# Shared Webドキュメント",
						version: 7,
					};
				},
			},
			wordEditSessionStore,
		});
		const request = new Request("http://localhost/api/word-edit-sessions", {
			method: "POST",
			headers: {
				cookie: `${editorSessionCookieName}=browser-session-1`,
			},
		});

		const firstResponse = await handlers.start(request);
		const secondResponse = await handlers.start(request);

		expect(firstResponse.status).toBe(201);
		expect(secondResponse.status).toBe(201);
		expect(convertedMarkdown).toEqual([
			"# Shared Webドキュメント",
			"# Shared Webドキュメント",
		]);
		expect(uploadedWorkingCopies).toEqual([
			{
				content: new TextEncoder().encode("# Shared Webドキュメント"),
				fileName: "Webドキュメント-word-session-1.docx",
				tokenCache: '{"RefreshToken":{"cached":true}}',
			},
			{
				content: new TextEncoder().encode("# Shared Webドキュメント"),
				fileName: "Webドキュメント-word-session-2.docx",
				tokenCache: '{"RefreshToken":{"cached":true}}',
			},
		]);
		expect(wordEditSessionStore.readSession("word-session-1")).toEqual({
			driveItemId: "drive-item-1",
			editorId: "editor-1",
			oneDriveWebUrl:
				"https://onedrive.example/Webドキュメント-word-session-1.docx",
			sessionId: "word-session-1",
			webDocumentId: "shared",
			webDocumentVersion: 7,
			workingCopyFileName: "Webドキュメント-word-session-1.docx",
		});
		expect(wordEditSessionStore.readSession("word-session-2")).toEqual({
			driveItemId: "drive-item-2",
			editorId: "editor-1",
			oneDriveWebUrl:
				"https://onedrive.example/Webドキュメント-word-session-2.docx",
			sessionId: "word-session-2",
			webDocumentId: "shared",
			webDocumentVersion: 7,
			workingCopyFileName: "Webドキュメント-word-session-2.docx",
		});
	});

	it("returns a Word起動導線 with an Office URI and OneDrive fallback link", async () => {
		const { editorAuthStore, wordEditSessionStore } = createStores();
		editorAuthStore.saveSignedInEditor({
			editor: {
				id: "editor-1",
				displayName: "編集者 A",
				username: "editor@example.com",
			},
			sessionId: "browser-session-1",
			tokenCache: '{"RefreshToken":{"cached":true}}',
		});
		const handlers = createWordEditSessionHandlers({
			converter: {
				async convertMarkdownToDocx(input) {
					return new TextEncoder().encode(input.markdown);
				},
			},
			createSessionId: () => "word-session-1",
			editorAuthStore,
			graph: {
				async uploadAppFolderWorkingCopy(input) {
					return {
						driveItemId: "drive-item-1",
						webUrl: `https://onedrive.example/${input.fileName}`,
					};
				},
			},
			webDocumentStore: {
				loadSharedDocument() {
					return {
						id: "shared",
						markdown: "# Shared Webドキュメント",
						version: 7,
					};
				},
			},
			wordEditSessionStore,
		});

		const response = await handlers.start(
			new Request("http://localhost/api/word-edit-sessions", {
				method: "POST",
				headers: {
					cookie: `${editorSessionCookieName}=browser-session-1`,
				},
			}),
		);

		expect(response.status).toBe(201);
		await expect(response.json()).resolves.toEqual({
			launchLinks: {
				officeUri:
					"ms-word:ofe|u|https://onedrive.example/Webドキュメント-word-session-1.docx",
				oneDriveFallbackUrl:
					"https://onedrive.example/Webドキュメント-word-session-1.docx",
			},
			sessionId: "word-session-1",
			workingCopy: {
				driveItemId: "drive-item-1",
				fileName: "Webドキュメント-word-session-1.docx",
			},
		});
	});
});

function createSequentialId(prefix: string): () => string {
	let nextId = 1;

	return () => `${prefix}-${nextId++}`;
}
