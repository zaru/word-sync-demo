import type { WebDocument } from "../domain/web-document-store";
import { editorSessionCookieName } from "./editor-auth-routes";
import type { EditorAuthStore } from "./editor-auth-store";
import type {
	WordEditSessionState,
	WordEditSessionStore,
} from "./word-edit-session-store";

export type MarkdownDocxConverter = {
	convertMarkdownToDocx(input: { markdown: string }): Promise<Uint8Array>;
	convertDocxToMarkdown(input: {
		content: Uint8Array;
	}): Promise<string | DocxToMarkdownResult>;
};

export type DocxToMarkdownResult = {
	markdown: string;
	notifications?: ImportNotification[];
};

export type ImportNotification = {
	message: string;
	type: "unsupportedContentDiscarded";
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
	deleteAppFolderWorkingCopy(input: {
		driveItemId: string;
		tokenCache: string;
	}): Promise<void>;
};

const abandonedSessionTimeoutMs = 2 * 60 * 60 * 1_000;
const workingCopyDeletionGraceMs = 24 * 60 * 60 * 1_000;

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

			const sessionState = options.wordEditSessionStore.readSessionState(
				wordEditSession.sessionId,
			);

			if (sessionState === "discarded" || sessionState === "finished") {
				return terminalSessionResponse({
					message: "Word編集セッション is already closed.",
					sessionId: wordEditSession.sessionId,
					status: sessionState,
				});
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

			let markdown: string;
			let notifications: ImportNotification[] = [];
			try {
				await options.waitForWorkingCopyToStabilize?.();
				const content = await options.graph.downloadAppFolderWorkingCopy({
					driveItemId: wordEditSession.driveItemId,
					tokenCache,
				});
				const result = await options.converter.convertDocxToMarkdown({
					content,
				});
				({ markdown, notifications } = normalizeDocxToMarkdownResult(result));
			} catch (error) {
				options.wordEditSessionStore.markSessionImportError(
					wordEditSession.sessionId,
				);

				return Response.json(
					{
						error: {
							message: importErrorMessage(error),
							type: "importError",
						},
						session: {
							sessionId: wordEditSession.sessionId,
							status: "importError",
						},
					},
					{ status: 409 },
				);
			}
			const webDocument = options.webDocumentStore.saveMarkdown(markdown);
			options.wordEditSessionStore.finishSession(wordEditSession.sessionId);

			return Response.json(
				notifications.length > 0
					? { notifications, webDocument }
					: { webDocument },
			);
		},

		async discard(
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

			const sessionState = options.wordEditSessionStore.readSessionState(
				wordEditSession.sessionId,
			);

			if (sessionState !== "importError") {
				return terminalSessionResponse({
					message: "Only a 取り込みエラー Word編集セッション can be discarded.",
					sessionId: wordEditSession.sessionId,
					status: sessionState ?? "active",
				});
			}

			options.wordEditSessionStore.discardSession(wordEditSession.sessionId);

			return Response.json({
				session: {
					sessionId: wordEditSession.sessionId,
					status: "discarded",
				},
			});
		},

		async cleanup(): Promise<Response> {
			const abandonedSessions =
				options.wordEditSessionStore.abandonInactiveSessions(
					abandonedSessionTimeoutMs,
				);
			const deletedWorkingCopies: Array<{
				driveItemId: string;
				sessionId: string;
			}> = [];
			const failures: Array<{
				driveItemId: string;
				message: string;
				sessionId: string;
			}> = [];

			for (const candidate of options.wordEditSessionStore.listWorkingCopiesEligibleForDeletion(
				workingCopyDeletionGraceMs,
			)) {
				const tokenCache = options.editorAuthStore.readTokenCache(
					candidate.editorId,
				);

				if (!tokenCache) {
					failures.push({
						driveItemId: candidate.driveItemId,
						message: "Graph access is required",
						sessionId: candidate.sessionId,
					});
					continue;
				}

				try {
					await options.graph.deleteAppFolderWorkingCopy({
						driveItemId: candidate.driveItemId,
						tokenCache,
					});
					options.wordEditSessionStore.markWorkingCopyDeleted(
						candidate.sessionId,
					);
					deletedWorkingCopies.push({
						driveItemId: candidate.driveItemId,
						sessionId: candidate.sessionId,
					});
				} catch (error) {
					failures.push({
						driveItemId: candidate.driveItemId,
						message: cleanupErrorMessage(error),
						sessionId: candidate.sessionId,
					});
				}
			}

			return Response.json(
				{
					abandonedSessions,
					deletedWorkingCopies,
					failures,
				},
				{ status: failures.length > 0 ? 207 : 200 },
			);
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

function importErrorMessage(error: unknown): string {
	return error instanceof Error
		? error.message
		: "終了取り込みでOneDrive作業コピーを取り込めませんでした。";
}

function cleanupErrorMessage(error: unknown): string {
	return error instanceof Error
		? error.message
		: "OneDrive作業コピー cleanup failed.";
}

function normalizeDocxToMarkdownResult(
	result: string | DocxToMarkdownResult,
): Required<DocxToMarkdownResult> {
	if (typeof result === "string") {
		return { markdown: result, notifications: [] };
	}

	return {
		markdown: result.markdown,
		notifications: result.notifications ?? [],
	};
}

function terminalSessionResponse(input: {
	message: string;
	sessionId: string;
	status: WordEditSessionState;
}): Response {
	return Response.json(
		{
			error: {
				message: input.message,
				type: "invalidSessionState",
			},
			session: {
				sessionId: input.sessionId,
				status: input.status,
			},
		},
		{ status: 409 },
	);
}
