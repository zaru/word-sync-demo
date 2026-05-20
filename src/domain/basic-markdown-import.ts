export type BasicMarkdownImport = {
	discardedUnsupportedContent: boolean;
	markdown: string;
};

const inlineImagePattern = /!\[[^\]]*]\([^)]+\)(?:\{[^}]*})?/g;

export function discardUnsupportedMarkdown(
	markdown: string,
): BasicMarkdownImport {
	const lines = markdown.split(/\r?\n/);
	const keptLines: string[] = [];
	let discardedUnsupportedContent = false;

	for (let index = 0; index < lines.length; index += 1) {
		const line = lines[index] ?? "";

		if (isTableStart(line, lines[index + 1])) {
			discardedUnsupportedContent = true;
			index += 1;

			while (index + 1 < lines.length && isTableRow(lines[index + 1] ?? "")) {
				index += 1;
			}

			continue;
		}

		if (isRawHtmlLine(line)) {
			discardedUnsupportedContent = true;
			continue;
		}

		const withoutImages = line.replace(inlineImagePattern, "").trimEnd();

		if (withoutImages !== line) {
			discardedUnsupportedContent = true;
		}

		if (withoutImages.trim() === "" && line.trim() !== "") {
			continue;
		}

		keptLines.push(withoutImages);
	}

	return {
		discardedUnsupportedContent,
		markdown: trimBlankLines(keptLines).join("\n"),
	};
}

function isTableStart(line: string, nextLine: string | undefined): boolean {
	return (
		isTableRow(line) && nextLine !== undefined && isTableSeparator(nextLine)
	);
}

function isTableRow(line: string): boolean {
	return line.includes("|") && line.trim().length > 0;
}

function isTableSeparator(line: string): boolean {
	return /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/.test(line);
}

function isRawHtmlLine(line: string): boolean {
	return /^\s*<\/?[a-z][^>]*>.*$/i.test(line);
}

function trimBlankLines(lines: string[]): string[] {
	let start = 0;
	let end = lines.length;

	while (start < end && lines[start]?.trim() === "") {
		start += 1;
	}

	while (end > start && lines[end - 1]?.trim() === "") {
		end -= 1;
	}

	return lines.slice(start, end);
}
