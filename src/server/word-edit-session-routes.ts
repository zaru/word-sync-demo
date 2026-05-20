import type { WebDocument } from "../domain/web-document-store";
import { editorSessionCookieName } from "./editor-auth-routes";
import type { EditorAuthStore } from "./editor-auth-store";
import type { WordEditSessionStore } from "./word-edit-session-store";

export type MarkdownDocxConverter = {
	convertMarkdownToDocx(input: { markdown: string }): Promise<Uint8Array>;
	convertDocxToMarkdown(input: { content: Uint8Array }): Promise<string>;
};

export type GraphAppFolderBoundary = {
	uploadAppFolderWorkingCopy(input: {
		content: Uint8Array;
		fileName: string;
		tokenCache: string;
	}): Promise<{
		driveItemId: string;
		webUrl: string;
	}>;
	downloadAppFolderWorkingCopy(input: {
		driveItemId: string;
		tokenCache: string;
	}): Promise<Uint8Array>;
};

export function createWordEditSessionHandlers(options: {
	converter: MarkdownDocxConverter;
	createSessionId: () => string;
	editorAuthStore: Pick<EditorAuthStore, "readSession" | "readTokenCache">;
	graph: GraphAppFolderBoundary;
	waitForWorkingCopyToStabilize?: () => Promise<void>;
	webDocumentStore: {
		loadSharedDocument(): WebDocument;
		saveMarkdown(markdown: string): WebDocument;
	};
	wordEditSessionStore: WordEditSessionStore;
}) {
	return {
		async start(request: Request): Promise<Response> {
			const browserSessionId = readCookie(request, editorSessionCookieName);
			const editorSession = browserSessionId
				? options.editorAuthStore.readSession(browserSessionId)
				: undefined;

			if (!editorSession) {
				return Response.json(
					{ error: "編集者 sign-in is required" },
					{ status: 401 },
				);
			}

			const tokenCache = options.editorAuthStore.readTokenCache(
				editorSession.editor.id,
			);

			if (!tokenCache) {
				return Response.json(
					{ error: "Graph access is required" },
					{ status: 401 },
				);
			}

			const sessionId = options.createSessionId();
			const webDocument = options.webDocumentStore.loadSharedDocument();
			const content = await options.converter.convertMarkdownToDocx({
				markdown: webDocument.markdown,
			});
			const workingCopyFileName = `Webドキュメント-${sessionId}.docx`;
			const uploadedWorkingCopy =
				await options.graph.uploadAppFolderWorkingCopy({
					content,
					fileName: workingCopyFileName,
					tokenCache,
				});
			const wordEditSession = options.wordEditSessionStore.saveStartedSession({
				driveItemId: uploadedWorkingCopy.driveItemId,
				editorId: editorSession.editor.id,
				oneDriveWebUrl: uploadedWorkingCopy.webUrl,
				sessionId,
				webDocumentId: webDocument.id,
				webDocumentVersion: webDocument.version,
				workingCopyFileName,
			});

			return Response.json(
				{
					launchLinks: {
						officeUri: createWordOfficeUri(wordEditSession.oneDriveWebUrl),
						oneDriveFallbackUrl: wordEditSession.oneDriveWebUrl,
					},
					sessionId: wordEditSession.sessionId,
					workingCopy: {
						driveItemId: wordEditSession.driveItemId,
						fileName: wordEditSession.workingCopyFileName,
					},
				},
				{ status: 201 },
			);
		},

		async finish(
			request: Request,
			params: { sessionId: string },
		): Promise<Response> {
			const browserSessionId = readCookie(request, editorSessionCookieName);
			const editorSession = browserSessionId
				? options.editorAuthStore.readSession(browserSessionId)
				: undefined;

			if (!editorSession) {
				return Response.json(
					{ error: "編集者 sign-in is required" },
					{ status: 401 },
				);
			}

			const wordEditSession = options.wordEditSessionStore.readSession(
				params.sessionId,
			);

			if (!wordEditSession) {
				return Response.json(
					{ error: "Word編集セッション was not found" },
					{ status: 404 },
				);
			}

			if (wordEditSession.editorId !== editorSession.editor.id) {
				return Response.json(
					{ error: "Word編集セッション belongs to a different 編集者" },
					{ status: 403 },
				);
			}

			const tokenCache = options.editorAuthStore.readTokenCache(
				editorSession.editor.id,
			);

			if (!tokenCache) {
				return Response.json(
					{ error: "Graph access is required" },
					{ status: 401 },
				);
			}

			await options.waitForWorkingCopyToStabilize?.();
			const content = await options.graph.downloadAppFolderWorkingCopy({
				driveItemId: wordEditSession.driveItemId,
				tokenCache,
			});
			const markdown = await options.converter.convertDocxToMarkdown({
				content,
			});
			const webDocument = options.webDocumentStore.saveMarkdown(markdown);
			options.wordEditSessionStore.finishSession(wordEditSession.sessionId);

			return Response.json({ webDocument });
		},
	};
}

function createWordOfficeUri(oneDriveWebUrl: string): string {
	return `ms-word:ofe|u|${oneDriveWebUrl}`;
}

function readCookie(request: Request, name: string): string | undefined {
	const cookieHeader = request.headers.get("cookie");

	if (!cookieHeader) {
		return undefined;
	}

	for (const cookie of cookieHeader.split(/; */)) {
		const [rawName, ...rawValueParts] = cookie.split("=");

		if (rawName === name) {
			return decodeURIComponent(rawValueParts.join("="));
		}
	}

	return undefined;
}
