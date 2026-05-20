// @vitest-environment happy-dom

import { afterEach, describe, expect, it, vi } from "vitest";

import type { WebDocument } from "../../domain/web-document-store";

describe("WebDocumentEditor", () => {
	afterEach(() => {
		vi.restoreAllMocks();
		document.body.replaceChildren();
	});

	it("hydrates a non-empty Markdown互換内容 projection without a React mismatch", async () => {
		const initialDocument: WebDocument = {
			id: "shared",
			markdown: "# Existing Webドキュメント",
			version: 1,
		};
		const consoleError = vi
			.spyOn(console, "error")
			.mockImplementation(() => {});
		const host = document.createElement("div");

		host.innerHTML = await renderServerWebDocumentEditor(initialDocument);
		expect(host.innerHTML).toContain("editor-shell");
		expect(host.innerHTML).not.toContain("editor-placeholder");

		document.body.append(host);

		await hydrateClientWebDocumentEditor(host, initialDocument);

		expect(consoleError.mock.calls.flat()).not.toContainEqual(
			expect.stringContaining(
				"Hydration failed because the server rendered HTML didn't match the client",
			),
		);
	});
});

async function renderServerWebDocumentEditor(
	initialDocument: WebDocument,
): Promise<string> {
	vi.resetModules();
	vi.stubGlobal("window", undefined);
	vi.stubGlobal("document", undefined);

	try {
		const [{ createElement }, { renderToString }, { WebDocumentEditor }] =
			await Promise.all([
				import("react"),
				import("react-dom/server"),
				import("./web-document-editor"),
			]);

		return renderToString(
			createElement(WebDocumentEditor, { initialDocument }),
		);
	} finally {
		vi.unstubAllGlobals();
	}
}

async function hydrateClientWebDocumentEditor(
	host: HTMLElement,
	initialDocument: WebDocument,
): Promise<void> {
	vi.resetModules();
	Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true });

	const [{ act, createElement }, { hydrateRoot }, { WebDocumentEditor }] =
		await Promise.all([
			import("react"),
			import("react-dom/client"),
			import("./web-document-editor"),
		]);

	await act(async () => {
		hydrateRoot(host, createElement(WebDocumentEditor, { initialDocument }));
	});
}
