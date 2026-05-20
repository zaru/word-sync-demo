import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { afterEach, describe, expect, it } from "vitest";

import { WebDocumentStore } from "./web-document-store";

describe("WebDocumentStore", () => {
	const tempDirs: string[] = [];

	afterEach(() => {
		for (const tempDir of tempDirs.splice(0)) {
			rmSync(tempDir, { recursive: true, force: true });
		}
	});

	function createStore(
		seedMarkdown = "# Shared Webドキュメント\n\n最初のMarkdown互換内容。",
	) {
		const tempDir = mkdtempSync(join(tmpdir(), "word-sync-demo-"));
		tempDirs.push(tempDir);

		return new WebDocumentStore({
			databasePath: join(tempDir, "web-document.sqlite"),
			seedMarkdown,
		});
	}

	it("creates and reloads the single shared Webドキュメント", () => {
		const firstStore = createStore();

		const firstLoad = firstStore.loadSharedDocument();

		expect(firstLoad).toEqual({
			id: "shared",
			markdown: "# Shared Webドキュメント\n\n最初のMarkdown互換内容。",
			version: 1,
		});

		const secondStore = new WebDocumentStore({
			databasePath: firstStore.databasePath,
			seedMarkdown: "this seed must not replace the persisted Webドキュメント",
		});

		expect(secondStore.loadSharedDocument()).toEqual(firstLoad);
	});

	it("persists autosaved Markdown and increments the version each time", () => {
		const store = createStore();

		const firstSave = store.saveMarkdown(
			"## Updated Webドキュメント\n\nAutosaved Markdown互換内容。",
		);
		const secondSave = store.saveMarkdown(
			"## Updated again\n\n別タブにも見える内容。",
		);

		expect(firstSave).toEqual({
			id: "shared",
			markdown: "## Updated Webドキュメント\n\nAutosaved Markdown互換内容。",
			version: 2,
		});
		expect(secondSave).toEqual({
			id: "shared",
			markdown: "## Updated again\n\n別タブにも見える内容。",
			version: 3,
		});
		expect(store.loadSharedDocument()).toEqual(secondSave);
	});

	it("reads the current version without loading Markdown", () => {
		const store = createStore();

		expect(store.readVersion()).toBe(1);

		store.saveMarkdown("A later autosave.");

		expect(store.readVersion()).toBe(2);
	});
});
