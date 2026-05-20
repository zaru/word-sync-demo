import type { WebDocument } from "../domain/web-document-store";
import type {
	WordEditSessionState,
	WordEditSessionStore,
} from "./word-edit-session-store";

export type EditorIdentity = {
	id: string;
	displayName?: string;
	username?: string;
};

export type ImportNotification = {
	message: string;
	type: "unsupportedContentDiscarded";
};

export type DocxToMarkdownResult = {
	markdown: string;
	notifications?: ImportNotification[];
};

export type MarkdownDocxConverter = {
	convertMarkdownToDocx(input: { markdown: string }): Promise<Uint8Array>;
	convertDocxToMarkdown(input: {
		content: Uint8Array;
	}): Promise<string | DocxToMarkdownResult>;
};

type OfficeUriMetadata = {
	contentId: string;
	objectResourceId: string;
};

export type GraphAppFolderBoundary = {
	uploadAppFolderWorkingCopy(input: {
		content: Uint8Array;
		fileName: string;
		tokenCache: string;
	}): Promise<{
		driveItemId: string;
		officeUriMetadata?: OfficeUriMetadata;
		webDavUrl?: string;
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

export type OneDriveWorkingCopyAdapter = {
	create(input: {
		editor: EditorIdentity;
		fileName: string;
		markdown: string;
	}): Promise<{
		driveItemId: string;
		officeUriMetadata?: OfficeUriMetadata;
		webDavUrl?: string;
		webUrl: string;
	}>;
	read(input: {
		driveItemId: string;
		editor: EditorIdentity;
	}): Promise<Required<DocxToMarkdownResult>>;
	delete(input: {
		driveItemId: string;
		editor: Pick<EditorIdentity, "id">;
	}): Promise<void>;
};

export type StartWordEditSessionOutcome =
	| {
			kind: "started";
			launchLinks: {
				officeUri: string;
				oneDriveFallbackUrl: string;
			};
			sessionId: string;
			workingCopy: {
				driveItemId: string;
				fileName: string;
			};
	  }
	| {
			kind: "graphAccessRequired";
			message: string;
	  };

export type FinishWordEditSessionOutcome =
	| {
			kind: "finished";
			notifications: ImportNotification[];
			session: {
				sessionId: string;
				status: "finished";
			};
			webDocument: WebDocument;
	  }
	| {
			kind: "importError";
			error: {
				message: string;
				type: "importError";
			};
			session: {
				sessionId: string;
				status: "importError";
			};
	  }
	| SessionAccessOutcome
	| InvalidSessionStateOutcome;

export type DiscardImportErrorOutcome =
	| {
			kind: "discarded";
			session: {
				sessionId: string;
				status: "discarded";
			};
	  }
	| SessionAccessOutcome
	| InvalidSessionStateOutcome;

export type CleanupWordEditSessionsOutcome = {
	abandonedSessions: string[];
	deletedWorkingCopies: Array<{
		driveItemId: string;
		sessionId: string;
	}>;
	failures: Array<{
		driveItemId: string;
		message: string;
		sessionId: string;
	}>;
};

export type WordEditSessionLifecycle = {
	cleanup(): Promise<CleanupWordEditSessionsOutcome>;
	discardImportError(input: {
		editor: EditorIdentity;
		sessionId: string;
	}): Promise<DiscardImportErrorOutcome>;
	finish(input: {
		editor: EditorIdentity;
		sessionId: string;
	}): Promise<FinishWordEditSessionOutcome>;
	start(input: {
		editor: EditorIdentity;
	}): Promise<StartWordEditSessionOutcome>;
};

type SessionAccessOutcome =
	| { kind: "notFound"; sessionId: string }
	| { kind: "differentEditor"; sessionId: string };

type InvalidSessionStateOutcome = {
	kind: "invalidSessionState";
	message: string;
	session: {
		sessionId: string;
		status: WordEditSessionState;
	};
};

const abandonedSessionTimeoutMs = 2 * 60 * 60 * 1_000;
const workingCopyDeletionGraceMs = 24 * 60 * 60 * 1_000;

export function createWordEditSessionLifecycle(options: {
	createOfficeUriTimestamp?: () => number;
	createSessionId: () => string;
	oneDriveWorkingCopies: OneDriveWorkingCopyAdapter;
	webDocumentStore: {
		loadSharedDocument(): WebDocument;
		saveMarkdown(markdown: string): WebDocument;
	};
	wordEditSessionStore: WordEditSessionStore;
}): WordEditSessionLifecycle {
	return {
		async start(input) {
			const sessionId = options.createSessionId();
			const webDocument = options.webDocumentStore.loadSharedDocument();
			const workingCopyFileName = `Webドキュメント-${sessionId}.docx`;
			let workingCopy: Awaited<
				ReturnType<OneDriveWorkingCopyAdapter["create"]>
			>;

			try {
				workingCopy = await options.oneDriveWorkingCopies.create({
					editor: input.editor,
					fileName: workingCopyFileName,
					markdown: webDocument.markdown,
				});
			} catch (error) {
				if (isGraphAccessRequiredError(error)) {
					return {
						kind: "graphAccessRequired",
						message: "Graph access is required",
					};
				}

				throw error;
			}

			const wordEditSession = options.wordEditSessionStore.saveStartedSession({
				driveItemId: workingCopy.driveItemId,
				editorId: input.editor.id,
				oneDriveWebUrl: workingCopy.webUrl,
				sessionId,
				webDocumentId: webDocument.id,
				webDocumentVersion: webDocument.version,
				workingCopyFileName,
			});

			return {
				kind: "started",
				launchLinks: {
					officeUri: createWordOfficeUri(
						{
							metadata: workingCopy.officeUriMetadata,
							url: workingCopy.webDavUrl ?? wordEditSession.oneDriveWebUrl,
						},
						options.createOfficeUriTimestamp ?? Date.now,
					),
					oneDriveFallbackUrl: wordEditSession.oneDriveWebUrl,
				},
				sessionId: wordEditSession.sessionId,
				workingCopy: {
					driveItemId: wordEditSession.driveItemId,
					fileName: wordEditSession.workingCopyFileName,
				},
			};
		},

		async finish(input) {
			const accessOutcome = readOwnedSessionState({
				editor: input.editor,
				sessionId: input.sessionId,
				store: options.wordEditSessionStore,
			});

			if (accessOutcome.kind !== "owned") {
				return accessOutcome;
			}

			if (
				accessOutcome.state === "discarded" ||
				accessOutcome.state === "finished" ||
				accessOutcome.state === "abandoned"
			) {
				return invalidSessionState({
					message: "Word編集セッション is already closed.",
					sessionId: input.sessionId,
					status: accessOutcome.state,
				});
			}

			let importedWorkingCopy: Required<DocxToMarkdownResult>;
			try {
				importedWorkingCopy = await options.oneDriveWorkingCopies.read({
					driveItemId: accessOutcome.session.driveItemId,
					editor: input.editor,
				});
			} catch (error) {
				options.wordEditSessionStore.markSessionImportError(input.sessionId);

				return {
					kind: "importError",
					error: {
						message: lifecycleErrorMessage(
							error,
							"終了取り込みでOneDrive作業コピーを取り込めませんでした。",
						),
						type: "importError",
					},
					session: {
						sessionId: input.sessionId,
						status: "importError",
					},
				};
			}
			const webDocument = options.webDocumentStore.saveMarkdown(
				importedWorkingCopy.markdown,
			);
			options.wordEditSessionStore.finishSession(input.sessionId);

			return {
				kind: "finished",
				notifications: importedWorkingCopy.notifications,
				session: {
					sessionId: input.sessionId,
					status: "finished",
				},
				webDocument,
			};
		},

		async discardImportError(input) {
			const accessOutcome = readOwnedSessionState({
				editor: input.editor,
				sessionId: input.sessionId,
				store: options.wordEditSessionStore,
			});

			if (accessOutcome.kind !== "owned") {
				return accessOutcome;
			}

			if (accessOutcome.state !== "importError") {
				return invalidSessionState({
					message: "Only a 取り込みエラー Word編集セッション can be discarded.",
					sessionId: input.sessionId,
					status: accessOutcome.state,
				});
			}

			options.wordEditSessionStore.discardSession(input.sessionId);

			return {
				kind: "discarded",
				session: {
					sessionId: input.sessionId,
					status: "discarded",
				},
			};
		},

		async cleanup() {
			const abandonedSessions =
				options.wordEditSessionStore.abandonInactiveSessions(
					abandonedSessionTimeoutMs,
				);
			const deletedWorkingCopies: CleanupWordEditSessionsOutcome["deletedWorkingCopies"] =
				[];
			const failures: CleanupWordEditSessionsOutcome["failures"] = [];

			for (const candidate of options.wordEditSessionStore.listWorkingCopiesEligibleForDeletion(
				workingCopyDeletionGraceMs,
			)) {
				try {
					await options.oneDriveWorkingCopies.delete({
						driveItemId: candidate.driveItemId,
						editor: { id: candidate.editorId },
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
						message: lifecycleErrorMessage(
							error,
							"OneDrive作業コピー cleanup failed.",
						),
						sessionId: candidate.sessionId,
					});
				}
			}

			return {
				abandonedSessions,
				deletedWorkingCopies,
				failures,
			};
		},
	};
}

const graphAccessRequiredErrorName = "GraphAccessRequiredError";

export function createGraphAccessRequiredError(): Error {
	const error = new Error("Graph access is required");
	error.name = graphAccessRequiredErrorName;

	return error;
}

function isGraphAccessRequiredError(error: unknown): boolean {
	return error instanceof Error && error.name === graphAccessRequiredErrorName;
}

function readOwnedSessionState(input: {
	editor: EditorIdentity;
	sessionId: string;
	store: WordEditSessionStore;
}):
	| {
			kind: "owned";
			session: NonNullable<ReturnType<WordEditSessionStore["readSession"]>>;
			state: WordEditSessionState;
	  }
	| SessionAccessOutcome {
	const session = input.store.readSession(input.sessionId);

	if (!session) {
		return { kind: "notFound", sessionId: input.sessionId };
	}

	if (session.editorId !== input.editor.id) {
		return { kind: "differentEditor", sessionId: input.sessionId };
	}

	return {
		kind: "owned",
		session,
		state: input.store.readSessionState(input.sessionId) ?? "active",
	};
}

function invalidSessionState(input: {
	message: string;
	sessionId: string;
	status: WordEditSessionState;
}): InvalidSessionStateOutcome {
	return {
		kind: "invalidSessionState",
		message: input.message,
		session: {
			sessionId: input.sessionId,
			status: input.status,
		},
	};
}

function createWordOfficeUri(
	input: { metadata?: OfficeUriMetadata; url: string },
	createTimestamp: () => number,
): string {
	const url = new URL(input.url).href;

	if (!input.metadata) {
		return `ms-word:ofe|u|${url}`;
	}

	return [
		"ms-word:ofe",
		"or",
		input.metadata.objectResourceId,
		"cid",
		input.metadata.contentId,
		"ct",
		String(createTimestamp()),
		"u",
		url,
	].join("|");
}

function lifecycleErrorMessage(error: unknown, fallback: string): string {
	return error instanceof Error ? error.message : fallback;
}
