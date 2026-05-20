import type { WebDocument } from "../domain/web-document-store";
import { editorSessionCookieName } from "./editor-auth-routes";
import type { EditorAuthStore } from "./editor-auth-store";
import type {
	DiscardImportErrorOutcome,
	DocxToMarkdownResult,
	FinishWordEditSessionOutcome,
	GraphAppFolderBoundary,
	MarkdownDocxConverter,
	OneDriveWorkingCopyAdapter,
	StartWordEditSessionOutcome,
} from "./word-edit-session-lifecycle";
import {
	createGraphAccessRequiredError,
	createWordEditSessionLifecycle,
} from "./word-edit-session-lifecycle";
import type { WordEditSessionStore } from "./word-edit-session-store";

export function createWordEditSessionHandlers(options: {
	converter: MarkdownDocxConverter;
	createOfficeUriTimestamp?: () => number;
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
	const lifecycle = createWordEditSessionLifecycle({
		createOfficeUriTimestamp: options.createOfficeUriTimestamp,
		createSessionId: options.createSessionId,
		oneDriveWorkingCopies: createOneDriveWorkingCopyAdapter(options),
		webDocumentStore: options.webDocumentStore,
		wordEditSessionStore: options.wordEditSessionStore,
	});

	return {
		async start(request: Request): Promise<Response> {
			const editorSession = readEditorSession(request, options.editorAuthStore);

			if (!editorSession) {
				return Response.json(
					{ error: "編集者 sign-in is required" },
					{ status: 401 },
				);
			}

			return startOutcomeResponse(
				await lifecycle.start({ editor: editorSession.editor }),
			);
		},

		async finish(
			request: Request,
			params: { sessionId: string },
		): Promise<Response> {
			const editorSession = readEditorSession(request, options.editorAuthStore);

			if (!editorSession) {
				return Response.json(
					{ error: "編集者 sign-in is required" },
					{ status: 401 },
				);
			}

			return finishOutcomeResponse(
				await lifecycle.finish({
					editor: editorSession.editor,
					sessionId: params.sessionId,
				}),
			);
		},

		async discard(
			request: Request,
			params: { sessionId: string },
		): Promise<Response> {
			const editorSession = readEditorSession(request, options.editorAuthStore);

			if (!editorSession) {
				return Response.json(
					{ error: "編集者 sign-in is required" },
					{ status: 401 },
				);
			}

			return discardOutcomeResponse(
				await lifecycle.discardImportError({
					editor: editorSession.editor,
					sessionId: params.sessionId,
				}),
			);
		},

		async cleanup(): Promise<Response> {
			const outcome = await lifecycle.cleanup();

			return Response.json(outcome, {
				status: outcome.failures.length > 0 ? 207 : 200,
			});
		},
	};
}

function createOneDriveWorkingCopyAdapter(options: {
	converter: MarkdownDocxConverter;
	editorAuthStore: Pick<EditorAuthStore, "readTokenCache">;
	graph: GraphAppFolderBoundary;
	waitForWorkingCopyToStabilize?: () => Promise<void>;
}): OneDriveWorkingCopyAdapter {
	return {
		async create(input) {
			const tokenCache = readRequiredTokenCache(
				options.editorAuthStore,
				input.editor.id,
			);
			const content = await options.converter.convertMarkdownToDocx({
				markdown: input.markdown,
			});

			return options.graph.uploadAppFolderWorkingCopy({
				content,
				fileName: input.fileName,
				tokenCache,
			});
		},

		async read(input) {
			const tokenCache = readRequiredTokenCache(
				options.editorAuthStore,
				input.editor.id,
			);

			await options.waitForWorkingCopyToStabilize?.();
			const content = await options.graph.downloadAppFolderWorkingCopy({
				driveItemId: input.driveItemId,
				tokenCache,
			});
			const result = await options.converter.convertDocxToMarkdown({ content });

			return normalizeDocxToMarkdownResult(result);
		},

		async delete(input) {
			const tokenCache = readRequiredTokenCache(
				options.editorAuthStore,
				input.editor.id,
			);

			await options.graph.deleteAppFolderWorkingCopy({
				driveItemId: input.driveItemId,
				tokenCache,
			});
		},
	};
}

function readRequiredTokenCache(
	editorAuthStore: Pick<EditorAuthStore, "readTokenCache">,
	editorId: string,
): string {
	const tokenCache = editorAuthStore.readTokenCache(editorId);

	if (!tokenCache) {
		throw createGraphAccessRequiredError();
	}

	return tokenCache;
}

function readEditorSession(
	request: Request,
	editorAuthStore: Pick<EditorAuthStore, "readSession">,
) {
	const browserSessionId = readCookie(request, editorSessionCookieName);

	return browserSessionId
		? editorAuthStore.readSession(browserSessionId)
		: undefined;
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

function startOutcomeResponse(outcome: StartWordEditSessionOutcome): Response {
	if (outcome.kind === "graphAccessRequired") {
		return Response.json({ error: outcome.message }, { status: 401 });
	}

	return Response.json(
		{
			launchLinks: outcome.launchLinks,
			sessionId: outcome.sessionId,
			workingCopy: outcome.workingCopy,
		},
		{ status: 201 },
	);
}

function finishOutcomeResponse(
	outcome: FinishWordEditSessionOutcome,
): Response {
	if (outcome.kind === "notFound") {
		return Response.json(
			{ error: "Word編集セッション was not found" },
			{ status: 404 },
		);
	}

	if (outcome.kind === "differentEditor") {
		return Response.json(
			{ error: "Word編集セッション belongs to a different 編集者" },
			{ status: 403 },
		);
	}

	if (outcome.kind === "invalidSessionState") {
		return terminalSessionResponse(outcome);
	}

	if (outcome.kind === "importError") {
		return Response.json(
			{
				error: outcome.error,
				session: outcome.session,
			},
			{ status: 409 },
		);
	}

	return Response.json(
		outcome.notifications.length > 0
			? {
					notifications: outcome.notifications,
					webDocument: outcome.webDocument,
				}
			: { webDocument: outcome.webDocument },
	);
}

function discardOutcomeResponse(outcome: DiscardImportErrorOutcome): Response {
	if (outcome.kind === "notFound") {
		return Response.json(
			{ error: "Word編集セッション was not found" },
			{ status: 404 },
		);
	}

	if (outcome.kind === "differentEditor") {
		return Response.json(
			{ error: "Word編集セッション belongs to a different 編集者" },
			{ status: 403 },
		);
	}

	if (outcome.kind === "invalidSessionState") {
		return terminalSessionResponse(outcome);
	}

	return Response.json({
		session: outcome.session,
	});
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
	session: {
		sessionId: string;
		status: string;
	};
}): Response {
	return Response.json(
		{
			error: {
				message: input.message,
				type: "invalidSessionState",
			},
			session: {
				sessionId: input.session.sessionId,
				status: input.session.status,
			},
		},
		{ status: 409 },
	);
}
