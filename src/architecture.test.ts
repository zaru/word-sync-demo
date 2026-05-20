import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

const repoRoot = path.resolve(__dirname, "..");
const appDir = path.join(repoRoot, "src/app");
const srcDir = path.join(repoRoot, "src");

const nextSpecialFileNames = new Set([
	"default.tsx",
	"error.tsx",
	"global-error.tsx",
	"head.tsx",
	"layout.tsx",
	"loading.tsx",
	"not-found.tsx",
	"page.tsx",
	"route.ts",
	"template.tsx",
]);

const legacyClassFiles = new Set([
	"src/domain/web-document-store.ts",
	"src/server/editor-auth-store.ts",
]);

describe("architecture conventions", () => {
	it("keeps non-route app files in underscored colocation directories", () => {
		expect(findAppPlacementViolations()).toEqual([]);
	});

	it("does not introduce new application classes", () => {
		expect(findClassDeclarationViolations()).toEqual([]);
	});
});

function findAppPlacementViolations(): string[] {
	const violations: string[] = [];

	walkAppSegments(appDir, violations);

	return violations.sort();
}

function walkAppSegments(currentDir: string, violations: string[]): void {
	const entries = readdirSync(currentDir, { withFileTypes: true });
	const fileNames = new Set(
		entries.filter((entry) => entry.isFile()).map((entry) => entry.name),
	);

	for (const entry of entries) {
		const entryPath = path.join(currentDir, entry.name);

		if (entry.isDirectory()) {
			if (!entry.name.startsWith("_")) {
				walkAppSegments(entryPath, violations);
			}
			continue;
		}

		if (
			!entry.isFile() ||
			isAllowedAppSegmentFile(currentDir, entry.name, fileNames)
		) {
			continue;
		}

		violations.push(toRepoPath(entryPath));
	}
}

function isAllowedAppSegmentFile(
	currentDir: string,
	fileName: string,
	siblingFileNames: Set<string>,
): boolean {
	if (isNextSpecialFile(fileName)) {
		return true;
	}

	if (currentDir === appDir && fileName === "globals.css") {
		return true;
	}

	return isTestBesideTarget(fileName, siblingFileNames);
}

function isNextSpecialFile(fileName: string): boolean {
	return (
		nextSpecialFileNames.has(fileName) ||
		/^(manifest|robots|sitemap)\.(js|ts)$/.test(fileName) ||
		/^(icon|apple-icon|opengraph-image|twitter-image)\.(ico|jpg|jpeg|png|svg|js|jsx|ts|tsx)$/.test(
			fileName,
		) ||
		fileName === "favicon.ico"
	);
}

function isTestBesideTarget(
	fileName: string,
	siblingFileNames: Set<string>,
): boolean {
	const match = /^(?<target>.+)\.test\.(?<extension>ts|tsx)$/.exec(fileName);

	if (!match?.groups) {
		return false;
	}

	const { target, extension } = match.groups;

	return siblingFileNames.has(`${target}.${extension}`);
}

function findClassDeclarationViolations(): string[] {
	return findSourceFiles(srcDir)
		.filter(
			(filePath) =>
				!filePath.endsWith(".test.ts") && !filePath.endsWith(".test.tsx"),
		)
		.filter((filePath) => !legacyClassFiles.has(toRepoPath(filePath)))
		.flatMap((filePath) => {
			const source = readFileSync(filePath, "utf8");
			const matches = [...source.matchAll(/\bclass\s+[A-Z]\w*/g)];

			return matches.map(
				(match) =>
					`${toRepoPath(filePath)}:${lineNumberAt(source, match.index ?? 0)}`,
			);
		})
		.sort();
}

function findSourceFiles(currentDir: string): string[] {
	return readdirSync(currentDir, { withFileTypes: true }).flatMap((entry) => {
		const entryPath = path.join(currentDir, entry.name);

		if (entry.isDirectory()) {
			return findSourceFiles(entryPath);
		}

		if (entry.isFile() && /\.(ts|tsx)$/.test(entry.name)) {
			return [entryPath];
		}

		return [];
	});
}

function lineNumberAt(source: string, index: number): number {
	return source.slice(0, index).split("\n").length;
}

function toRepoPath(filePath: string): string {
	return path.relative(repoRoot, filePath).split(path.sep).join("/");
}
