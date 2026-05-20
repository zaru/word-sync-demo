import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { afterEach, describe, expect, it } from "vitest";
import { WebDocumentStore } from "../domain/web-document-store";
import {
	createWordEditSessionLifecycle,
	type EditorIdentity,
	type OneDriveWorkingCopyAdapter,
} from "./word-edit-session-lifecycle";
import { createWordEditSessionStore } from "./word-edit-session-store";

describe("Word編集セッション lifecycle", () => {
	const tempDirs: string[] = [];
	const editor: EditorIdentity = {
		displayName: "編集者 A",
		id: "editor-1",
		username: "editor@example.com",
	};

	afterEach(() => {
		for (const tempDir of tempDirs.splice(0)) {
			rmSync(tempDir, { recursive: true, force: true });
		}
	});

	it("starts a Word編集セッション by creating a OneDrive作業コピー from the current Webドキュメント", async () => {
		const { webDocumentStore, wordEditSessionStore } = createStores();
		const uploadedWorkingCopies: Array<{
			fileName: string;
			markdown: string;
		}> = [];
		const lifecycle = createWordEditSessionLifecycle({
			createOfficeUriTimestamp: () => 1_779_288_599_962,
			createSessionId: () => "word-session-1",
			oneDriveWorkingCopies: {
				async create(input) {
					uploadedWorkingCopies.push({
						fileName: input.fileName,
						markdown: input.markdown,
					});

					return {
						driveItemId: "drive-item-1",
						officeUriMetadata: {
							contentId: "01f5d442-f410-4df9-9f58-d1ac3f8d46ba",
							objectResourceId: "81dd2b71-fb82-4b33-ac71-fed46bf0f87a",
						},
						webDavUrl:
							"https://d.docs.live.net/editor-drive/Webドキュメント-word-session-1.docx",
						webUrl: `https://onedrive.example/${input.fileName}`,
					};
				},
				async delete() {},
				async read() {
					throw new Error("read should not be called while starting.");
				},
			},
			webDocumentStore,
			wordEditSessionStore,
		});

		const outcome = await lifecycle.start({ editor });

		expect(outcome).toEqual({
			kind: "started",
			launchLinks: {
				officeUri:
					"ms-word:ofe|or|81dd2b71-fb82-4b33-ac71-fed46bf0f87a|cid|01f5d442-f410-4df9-9f58-d1ac3f8d46ba|ct|1779288599962|u|https://d.docs.live.net/editor-drive/Web%E3%83%89%E3%82%AD%E3%83%A5%E3%83%A1%E3%83%B3%E3%83%88-word-session-1.docx",
				oneDriveFallbackUrl:
					"https://onedrive.example/Webドキュメント-word-session-1.docx",
			},
			sessionId: "word-session-1",
			workingCopy: {
				driveItemId: "drive-item-1",
				fileName: "Webドキュメント-word-session-1.docx",
			},
		});
		expect(uploadedWorkingCopies).toEqual([
			{
				fileName: "Webドキュメント-word-session-1.docx",
				markdown: "# Before Word編集セッション",
			},
		]);
		expect(wordEditSessionStore.readSession("word-session-1")).toEqual({
			driveItemId: "drive-item-1",
			editorId: "editor-1",
			oneDriveWebUrl:
				"https://onedrive.example/Webドキュメント-word-session-1.docx",
			sessionId: "word-session-1",
			webDocumentId: "shared",
			webDocumentVersion: 1,
			workingCopyFileName: "Webドキュメント-word-session-1.docx",
		});
	});

	it("uses 最後取り込み優先 when later 終了取り込み replaces Web autosave changes", async () => {
		const { webDocumentStore, wordEditSessionStore } = createStores();
		const importedMarkdownByDriveItem = new Map([
			["drive-item-1", "# First Word import"],
			["drive-item-2", "# Second Word import"],
		]);
		const lifecycle = createWordEditSessionLifecycle({
			createSessionId: createSequentialId("word-session"),
			oneDriveWorkingCopies: createMemoryOneDriveWorkingCopies({
				readMarkdown(input) {
					const markdown = importedMarkdownByDriveItem.get(input.driveItemId);

					if (!markdown) {
						throw new Error(`Unexpected drive item: ${input.driveItemId}`);
					}

					return markdown;
				},
			}),
			webDocumentStore,
			wordEditSessionStore,
		});

		await lifecycle.start({ editor });
		await lifecycle.start({ editor });
		await lifecycle.finish({ editor, sessionId: "word-session-1" });
		webDocumentStore.saveMarkdown("# Web autosave between imports");
		const finalOutcome = await lifecycle.finish({
			editor,
			sessionId: "word-session-2",
		});

		expect(finalOutcome).toMatchObject({
			kind: "finished",
			webDocument: {
				id: "shared",
				markdown: "# Second Word import",
				version: 4,
			},
		});
		expect(webDocumentStore.loadSharedDocument()).toEqual({
			id: "shared",
			markdown: "# Second Word import",
			version: 4,
		});
	});

	it("enters 取り込みエラー without changing the Webドキュメント when OneDrive作業コピー import fails", async () => {
		const { webDocumentStore, wordEditSessionStore } = createStores();
		const lifecycle = createWordEditSessionLifecycle({
			createSessionId: () => "word-session-1",
			oneDriveWorkingCopies: createMemoryOneDriveWorkingCopies({
				readMarkdown() {
					throw new Error("OneDrive作業コピー was missing or deleted.");
				},
			}),
			webDocumentStore,
			wordEditSessionStore,
		});

		await lifecycle.start({ editor });
		const outcome = await lifecycle.finish({
			editor,
			sessionId: "word-session-1",
		});

		expect(outcome).toEqual({
			error: {
				message: "OneDrive作業コピー was missing or deleted.",
				type: "importError",
			},
			kind: "importError",
			session: {
				sessionId: "word-session-1",
				status: "importError",
			},
		});
		expect(webDocumentStore.loadSharedDocument()).toEqual({
			id: "shared",
			markdown: "# Before Word編集セッション",
			version: 1,
		});
		expect(wordEditSessionStore.readSessionState("word-session-1")).toBe(
			"importError",
		);
	});

	it("moves inactive sessions to 放置終了 and deletes due OneDrive作業コピー after 作業コピー削除猶予", async () => {
		let currentTime = new Date("2026-05-20T00:00:00.000Z");
		const { webDocumentStore, wordEditSessionStore } = createStores({
			now: () => currentTime,
		});
		const deletedWorkingCopies: Array<{
			driveItemId: string;
			editorId: string;
		}> = [];
		const lifecycle = createWordEditSessionLifecycle({
			createSessionId: createSequentialId("word-session"),
			oneDriveWorkingCopies: createMemoryOneDriveWorkingCopies({
				deletedWorkingCopies,
				readMarkdown(input) {
					return input.driveItemId === "drive-item-1"
						? "# Imported from finished session"
						: "# Should not import abandoned session";
				},
			}),
			webDocumentStore,
			wordEditSessionStore,
		});

		await lifecycle.start({ editor });
		await lifecycle.start({ editor });
		currentTime = new Date("2026-05-20T01:59:59.000Z");
		await lifecycle.finish({ editor, sessionId: "word-session-1" });
		currentTime = new Date("2026-05-20T02:00:01.000Z");
		const abandonOutcome = await lifecycle.cleanup();
		currentTime = new Date("2026-05-21T01:59:58.000Z");
		const beforeGraceOutcome = await lifecycle.cleanup();
		currentTime = new Date("2026-05-21T02:00:02.000Z");
		const afterGraceOutcome = await lifecycle.cleanup();

		expect(abandonOutcome).toEqual({
			abandonedSessions: ["word-session-2"],
			deletedWorkingCopies: [],
			failures: [],
		});
		expect(beforeGraceOutcome).toEqual({
			abandonedSessions: [],
			deletedWorkingCopies: [],
			failures: [],
		});
		expect(afterGraceOutcome).toEqual({
			abandonedSessions: [],
			deletedWorkingCopies: [
				{
					driveItemId: "drive-item-1",
					sessionId: "word-session-1",
				},
				{
					driveItemId: "drive-item-2",
					sessionId: "word-session-2",
				},
			],
			failures: [],
		});
		expect(deletedWorkingCopies).toEqual([
			{
				driveItemId: "drive-item-1",
				editorId: "editor-1",
			},
			{
				driveItemId: "drive-item-2",
				editorId: "editor-1",
			},
		]);
	});

	function createStores(options: { now?: () => Date } = {}) {
		const tempDir = mkdtempSync(join(tmpdir(), "word-sync-demo-lifecycle-"));
		tempDirs.push(tempDir);

		return {
			webDocumentStore: new WebDocumentStore({
				databasePath: join(tempDir, "web-document.sqlite"),
				seedMarkdown: "# Before Word編集セッション",
			}),
			wordEditSessionStore: createWordEditSessionStore({
				databasePath: join(tempDir, "word-edit-sessions.sqlite"),
				now: options.now,
			}),
		};
	}
});

function createMemoryOneDriveWorkingCopies(options: {
	deletedWorkingCopies?: Array<{
		driveItemId: string;
		editorId: string;
	}>;
	readMarkdown(input: { driveItemId: string }): string;
}): OneDriveWorkingCopyAdapter {
	let nextDriveItemId = 1;

	return {
		async create(input) {
			const driveItemId = `drive-item-${nextDriveItemId++}`;

			return {
				driveItemId,
				webUrl: `https://onedrive.example/${input.fileName}`,
			};
		},
		async delete(input) {
			options.deletedWorkingCopies?.push({
				driveItemId: input.driveItemId,
				editorId: input.editor.id,
			});
		},
		async read(input) {
			return {
				markdown: options.readMarkdown(input),
				notifications: [],
			};
		},
	};
}

function createSequentialId(prefix: string): () => string {
	let nextId = 1;

	return () => `${prefix}-${nextId++}`;
}
