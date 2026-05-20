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

	async function finishSessionWithImportFailure(options: {
		conversionError?: Error;
		downloadError?: Error;
		stabilizationError?: Error;
	}) {
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
		const webDocument = {
			id: "shared" as const,
			markdown: "# Before Word編集セッション",
			version: 3,
		};
		let saveMarkdownWasCalled = false;
		const handlers = createWordEditSessionHandlers({
			converter: {
				async convertMarkdownToDocx(input) {
					return new TextEncoder().encode(input.markdown);
				},
				async convertDocxToMarkdown() {
					if (options.conversionError) {
						throw options.conversionError;
					}

					return "# Imported from Word";
				},
			},
			createSessionId: () => "word-session-1",
			editorAuthStore,
			graph: {
				async uploadAppFolderWorkingCopy() {
					return {
						driveItemId: "drive-item-1",
						webUrl:
							"https://onedrive.example/Webドキュメント-word-session-1.docx",
					};
				},
				async downloadAppFolderWorkingCopy() {
					if (options.downloadError) {
						throw options.downloadError;
					}

					return new TextEncoder().encode("imported docx");
				},
			},
			async waitForWorkingCopyToStabilize() {
				if (options.stabilizationError) {
					throw options.stabilizationError;
				}
			},
			webDocumentStore: {
				loadSharedDocument() {
					return webDocument;
				},
				saveMarkdown() {
					saveMarkdownWasCalled = true;
					return {
						id: "shared",
						markdown: "# Should not be saved",
						version: 4,
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

		await handlers.start(request);
		const finishResponse = await handlers.finish(request, {
			sessionId: "word-session-1",
		});

		return {
			finishResponse,
			saveMarkdownWasCalled,
			webDocument,
			wordEditSessionStore,
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

	it("finishes a Word編集セッション by importing the latest OneDrive作業コピー into the Webドキュメント", async () => {
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
		const exportedDocx = new TextEncoder().encode("exported docx");
		const importedDocx = new TextEncoder().encode("imported docx");
		let workingCopyIsStable = false;
		let webDocument = {
			id: "shared" as const,
			markdown: "# Before Word編集セッション",
			version: 3,
		};
		const handlers = createWordEditSessionHandlers({
			converter: {
				async convertMarkdownToDocx() {
					return exportedDocx;
				},
				async convertDocxToMarkdown(input) {
					expect(input.content).toEqual(importedDocx);

					return "# Imported from Word";
				},
			},
			createSessionId: () => "word-session-1",
			editorAuthStore,
			graph: {
				async uploadAppFolderWorkingCopy() {
					return {
						driveItemId: "drive-item-1",
						webUrl:
							"https://onedrive.example/Webドキュメント-word-session-1.docx",
					};
				},
				async downloadAppFolderWorkingCopy(input) {
					expect(input).toEqual({
						driveItemId: "drive-item-1",
						tokenCache: '{"RefreshToken":{"cached":true}}',
					});
					expect(workingCopyIsStable).toBe(true);

					return importedDocx;
				},
			},
			async waitForWorkingCopyToStabilize() {
				workingCopyIsStable = true;
			},
			webDocumentStore: {
				loadSharedDocument() {
					return webDocument;
				},
				saveMarkdown(markdown) {
					webDocument = {
						id: "shared",
						markdown,
						version: webDocument.version + 1,
					};

					return webDocument;
				},
			},
			wordEditSessionStore,
		});
		const startRequest = new Request(
			"http://localhost/api/word-edit-sessions",
			{
				method: "POST",
				headers: {
					cookie: `${editorSessionCookieName}=browser-session-1`,
				},
			},
		);

		await handlers.start(startRequest);
		const finishResponse = await handlers.finish(
			new Request(
				"http://localhost/api/word-edit-sessions/word-session-1/finish",
				{
					method: "POST",
					headers: {
						cookie: `${editorSessionCookieName}=browser-session-1`,
					},
				},
			),
			{ sessionId: "word-session-1" },
		);

		expect(finishResponse.status).toBe(200);
		await expect(finishResponse.json()).resolves.toEqual({
			webDocument: {
				id: "shared",
				markdown: "# Imported from Word",
				version: 4,
			},
		});
		expect(webDocument).toEqual({
			id: "shared",
			markdown: "# Imported from Word",
			version: 4,
		});
	});

	it("finishes successfully with a visible 互換外破棄通知 when unsupported Word content is discarded", async () => {
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
		let webDocument = {
			id: "shared" as const,
			markdown: "# Before Word編集セッション",
			version: 3,
		};
		const handlers = createWordEditSessionHandlers({
			converter: {
				async convertMarkdownToDocx(input) {
					return new TextEncoder().encode(input.markdown);
				},
				async convertDocxToMarkdown() {
					return {
						markdown: "# Imported from Word",
						notifications: [
							{
								message:
									"基本Markdown要素として取り込めないWord編集を破棄しました。",
								type: "unsupportedContentDiscarded",
							},
						],
					};
				},
			},
			createSessionId: () => "word-session-1",
			editorAuthStore,
			graph: {
				async uploadAppFolderWorkingCopy() {
					return {
						driveItemId: "drive-item-1",
						webUrl:
							"https://onedrive.example/Webドキュメント-word-session-1.docx",
					};
				},
				async downloadAppFolderWorkingCopy() {
					return new TextEncoder().encode("imported docx");
				},
			},
			webDocumentStore: {
				loadSharedDocument() {
					return webDocument;
				},
				saveMarkdown(markdown) {
					webDocument = {
						id: "shared",
						markdown,
						version: webDocument.version + 1,
					};

					return webDocument;
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

		await handlers.start(request);
		const finishResponse = await handlers.finish(request, {
			sessionId: "word-session-1",
		});

		expect(finishResponse.status).toBe(200);
		await expect(finishResponse.json()).resolves.toEqual({
			notifications: [
				{
					message: "基本Markdown要素として取り込めないWord編集を破棄しました。",
					type: "unsupportedContentDiscarded",
				},
			],
			webDocument: {
				id: "shared",
				markdown: "# Imported from Word",
				version: 4,
			},
		});
		expect(webDocument).toEqual({
			id: "shared",
			markdown: "# Imported from Word",
			version: 4,
		});
		expect(wordEditSessionStore.readSessionState("word-session-1")).toBe(
			"finished",
		);
	});

	it("surfaces 取り込みエラー when the OneDrive作業コピー cannot be downloaded without changing the Webドキュメント", async () => {
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
		const webDocument = {
			id: "shared" as const,
			markdown: "# Before Word編集セッション",
			version: 3,
		};
		let saveMarkdownWasCalled = false;
		const handlers = createWordEditSessionHandlers({
			converter: {
				async convertMarkdownToDocx(input) {
					return new TextEncoder().encode(input.markdown);
				},
				async convertDocxToMarkdown() {
					throw new Error("DOCX conversion should not run.");
				},
			},
			createSessionId: () => "word-session-1",
			editorAuthStore,
			graph: {
				async uploadAppFolderWorkingCopy() {
					return {
						driveItemId: "drive-item-1",
						webUrl:
							"https://onedrive.example/Webドキュメント-word-session-1.docx",
					};
				},
				async downloadAppFolderWorkingCopy() {
					throw new Error("OneDrive作業コピー was missing or deleted.");
				},
			},
			webDocumentStore: {
				loadSharedDocument() {
					return webDocument;
				},
				saveMarkdown() {
					saveMarkdownWasCalled = true;
					return {
						id: "shared",
						markdown: "# Should not be saved",
						version: 4,
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

		await handlers.start(request);
		const finishResponse = await handlers.finish(request, {
			sessionId: "word-session-1",
		});

		expect(finishResponse.status).toBe(409);
		await expect(finishResponse.json()).resolves.toEqual({
			error: {
				message: "OneDrive作業コピー was missing or deleted.",
				type: "importError",
			},
			session: {
				sessionId: "word-session-1",
				status: "importError",
			},
		});
		expect(saveMarkdownWasCalled).toBe(false);
		expect(webDocument).toEqual({
			id: "shared",
			markdown: "# Before Word編集セッション",
			version: 3,
		});
		expect(wordEditSessionStore.readSessionState("word-session-1")).toBe(
			"importError",
		);
	});

	it.each([
		{
			errorOptions: {
				downloadError: new Error("Graph download access was denied."),
			},
			message: "Graph download access was denied.",
			name: "Graph access or download failure",
		},
		{
			errorOptions: {
				stabilizationError: new Error("OneDrive作業コピー did not stabilize."),
			},
			message: "OneDrive作業コピー did not stabilize.",
			name: "unstable OneDrive作業コピー",
		},
		{
			errorOptions: {
				conversionError: new Error("DOCX conversion failed."),
			},
			message: "DOCX conversion failed.",
			name: "DOCX conversion failure",
		},
	])("surfaces 取り込みエラー for $name", async ({ errorOptions, message }) => {
		const {
			finishResponse,
			saveMarkdownWasCalled,
			webDocument,
			wordEditSessionStore,
		} = await finishSessionWithImportFailure(errorOptions);

		expect(finishResponse.status).toBe(409);
		await expect(finishResponse.json()).resolves.toEqual({
			error: {
				message,
				type: "importError",
			},
			session: {
				sessionId: "word-session-1",
				status: "importError",
			},
		});
		expect(saveMarkdownWasCalled).toBe(false);
		expect(webDocument).toEqual({
			id: "shared",
			markdown: "# Before Word編集セッション",
			version: 3,
		});
		expect(wordEditSessionStore.readSessionState("word-session-1")).toBe(
			"importError",
		);
	});

	it("lets the 編集者 discard an errored Word編集セッション without changing the Webドキュメント", async () => {
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
		const webDocument = {
			id: "shared" as const,
			markdown: "# Before Word編集セッション",
			version: 3,
		};
		let saveMarkdownWasCalled = false;
		const handlers = createWordEditSessionHandlers({
			converter: {
				async convertMarkdownToDocx(input) {
					return new TextEncoder().encode(input.markdown);
				},
				async convertDocxToMarkdown() {
					throw new Error("DOCX conversion should not run.");
				},
			},
			createSessionId: () => "word-session-1",
			editorAuthStore,
			graph: {
				async uploadAppFolderWorkingCopy() {
					return {
						driveItemId: "drive-item-1",
						webUrl:
							"https://onedrive.example/Webドキュメント-word-session-1.docx",
					};
				},
				async downloadAppFolderWorkingCopy() {
					throw new Error("OneDrive作業コピー was missing or deleted.");
				},
			},
			webDocumentStore: {
				loadSharedDocument() {
					return webDocument;
				},
				saveMarkdown() {
					saveMarkdownWasCalled = true;
					return {
						id: "shared",
						markdown: "# Should not be saved",
						version: 4,
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

		await handlers.start(request);
		await handlers.finish(request, { sessionId: "word-session-1" });
		const discardResponse = await handlers.discard(request, {
			sessionId: "word-session-1",
		});

		expect(discardResponse.status).toBe(200);
		await expect(discardResponse.json()).resolves.toEqual({
			session: {
				sessionId: "word-session-1",
				status: "discarded",
			},
		});
		expect(saveMarkdownWasCalled).toBe(false);
		expect(webDocument).toEqual({
			id: "shared",
			markdown: "# Before Word編集セッション",
			version: 3,
		});
		expect(wordEditSessionStore.readSessionState("word-session-1")).toBe(
			"discarded",
		);
	});

	it("transitions the Word編集セッション to セッション終了 after successful 終了取り込み", async () => {
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
				async convertDocxToMarkdown() {
					return "# Imported from Word";
				},
			},
			createSessionId: () => "word-session-1",
			editorAuthStore,
			graph: {
				async uploadAppFolderWorkingCopy() {
					return {
						driveItemId: "drive-item-1",
						webUrl:
							"https://onedrive.example/Webドキュメント-word-session-1.docx",
					};
				},
				async downloadAppFolderWorkingCopy() {
					return new TextEncoder().encode("imported docx");
				},
			},
			webDocumentStore: {
				loadSharedDocument() {
					return {
						id: "shared",
						markdown: "# Before Word編集セッション",
						version: 3,
					};
				},
				saveMarkdown(markdown) {
					return {
						id: "shared",
						markdown,
						version: 4,
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

		await handlers.start(request);
		expect(wordEditSessionStore.readSessionState("word-session-1")).toBe(
			"active",
		);

		await handlers.finish(request, { sessionId: "word-session-1" });

		expect(wordEditSessionStore.readSessionState("word-session-1")).toBe(
			"finished",
		);
	});
});

function createSequentialId(prefix: string): () => string {
	let nextId = 1;

	return () => `${prefix}-${nextId++}`;
}
