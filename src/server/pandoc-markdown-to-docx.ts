import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";

import type { MarkdownToDocxConverter } from "./word-edit-session-routes";

const execFileAsync = promisify(execFile);

export function createPandocMarkdownToDocxConverter(): MarkdownToDocxConverter {
	return {
		async convertMarkdownToDocx(input) {
			const workingDirectory = await mkdtemp(
				join(tmpdir(), "word-sync-pandoc-"),
			);
			const markdownPath = join(workingDirectory, "web-document.md");
			const docxPath = join(workingDirectory, "onedrive-working-copy.docx");

			try {
				await writeFile(markdownPath, input.markdown, "utf8");
				await execFileAsync("pandoc", [markdownPath, "-o", docxPath]);

				return await readFile(docxPath);
			} finally {
				await rm(workingDirectory, { force: true, recursive: true });
			}
		},
	};
}
