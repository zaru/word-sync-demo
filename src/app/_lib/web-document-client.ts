import type { WebDocument } from "../../domain/web-document-store";

type DebouncedAutosaveOptions = {
	delayMs: number;
	saveMarkdown: (markdown: string) => Promise<WebDocument>;
	onSaved: (document: WebDocument) => void;
	onError?: (error: unknown) => void;
};

type WebDocumentVersion = Pick<WebDocument, "version">;

export function createDebouncedAutosave(options: DebouncedAutosaveOptions) {
	let timeout: ReturnType<typeof setTimeout> | undefined;

	return {
		schedule(markdown: string): void {
			if (timeout !== undefined) {
				clearTimeout(timeout);
			}

			timeout = setTimeout(() => {
				timeout = undefined;
				options
					.saveMarkdown(markdown)
					.then(options.onSaved, (error: unknown) => {
						if (options.onError) {
							options.onError(error);
							return;
						}

						throw error;
					});
			}, options.delayMs);
		},
		cancel(): void {
			if (timeout !== undefined) {
				clearTimeout(timeout);
				timeout = undefined;
			}
		},
	};
}

export async function loadChangedWebDocument(options: {
	currentVersion: number;
	fetchVersion: () => Promise<WebDocumentVersion>;
	fetchDocument: () => Promise<WebDocument>;
}): Promise<WebDocument | undefined> {
	const { version } = await options.fetchVersion();

	if (version === options.currentVersion) {
		return undefined;
	}

	return options.fetchDocument();
}
