import { describe, expect, it } from "vitest";

import { discardUnsupportedMarkdown } from "./basic-markdown-import";

describe("basic Markdown import", () => {
	it("keeps 基本Markdown要素 and discards common unsupported Markdown produced from Word", () => {
		const result = discardUnsupportedMarkdown(
			[
				"# Imported",
				"",
				"Paragraph with **strong** text.",
				"",
				"![diagram](media/image1.png)",
				"",
				"| unsupported | table |",
				"| --- | --- |",
				"| A | B |",
				"",
				"<aside>unsupported HTML</aside>",
			].join("\n"),
		);

		expect(result).toEqual({
			discardedUnsupportedContent: true,
			markdown: "# Imported\n\nParagraph with **strong** text.",
		});
	});
});
