import { describe, expect, it, vi } from "vitest";

import {
	createDebouncedAutosave,
	loadChangedWebDocument,
} from "./web-document-client";

describe("Webドキュメント client behavior", () => {
	it("autosaves only the latest Markdown after the debounce delay", async () => {
		vi.useFakeTimers();
		const saveMarkdown = vi.fn(async (markdown: string) => ({
			id: "shared" as const,
			markdown,
			version: 2,
		}));
		const onSaved = vi.fn();

		const autosave = createDebouncedAutosave({
			delayMs: 500,
			saveMarkdown,
			onSaved,
		});

		autosave.schedule("# first draft");
		autosave.schedule("# final draft");
		await vi.advanceTimersByTimeAsync(499);

		expect(saveMarkdown).not.toHaveBeenCalled();

		await vi.advanceTimersByTimeAsync(1);

		expect(saveMarkdown).toHaveBeenCalledTimes(1);
		expect(saveMarkdown).toHaveBeenCalledWith("# final draft");
		expect(onSaved).toHaveBeenCalledWith({
			id: "shared",
			markdown: "# final draft",
			version: 2,
		});

		vi.useRealTimers();
	});

	it("loads the Webドキュメント only when polling observes a changed version", async () => {
		const fetchVersion = vi.fn(async () => ({ version: 3 }));
		const fetchDocument = vi.fn(async () => ({
			id: "shared" as const,
			markdown: "# Reloaded",
			version: 3,
		}));

		await expect(
			loadChangedWebDocument({
				currentVersion: 2,
				fetchVersion,
				fetchDocument,
			}),
		).resolves.toEqual({
			id: "shared",
			markdown: "# Reloaded",
			version: 3,
		});

		fetchVersion.mockResolvedValueOnce({ version: 3 });

		await expect(
			loadChangedWebDocument({
				currentVersion: 3,
				fetchVersion,
				fetchDocument,
			}),
		).resolves.toBeUndefined();
		expect(fetchDocument).toHaveBeenCalledTimes(1);
	});
});
