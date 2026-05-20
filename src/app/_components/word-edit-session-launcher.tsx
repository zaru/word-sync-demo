"use client";

import { useState } from "react";
import type { WebDocument } from "../../domain/web-document-store";

type WordEditSessionResponse = {
	sessionId: string;
	workingCopy: {
		driveItemId: string;
		fileName: string;
	};
	launchLinks: {
		officeUri: string;
		oneDriveFallbackUrl: string;
	};
};

type FinishWordEditSessionResponse = {
	notifications?: Array<{
		message: string;
		type: "unsupportedContentDiscarded";
	}>;
	webDocument: WebDocument;
};

type ImportErrorResponse = {
	error: {
		message: string;
		type: "importError";
	};
	session: {
		sessionId: string;
		status: "importError";
	};
};

type DiscardWordEditSessionResponse = {
	session: {
		sessionId: string;
		status: "discarded";
	};
};

type SessionState =
	| { status: "ready"; session: WordEditSessionResponse }
	| { status: "finishing"; session: WordEditSessionResponse }
	| { status: "discarding"; session: WordEditSessionResponse }
	| {
			status: "importError";
			session: WordEditSessionResponse;
			message: string;
	  }
	| {
			status: "finished";
			notifications: FinishWordEditSessionResponse["notifications"];
			session: WordEditSessionResponse;
			webDocument: WebDocument;
	  };

type LaunchState = {
	discardedMessage?: string;
	errorMessage?: string;
	sessions: SessionState[];
	startStatus: "idle" | "starting";
};

export function WordEditSessionLauncher() {
	const [state, setState] = useState<LaunchState>({
		sessions: [],
		startStatus: "idle",
	});
	const sessionActionIsPending = state.sessions.some(
		(session) =>
			session.status === "finishing" || session.status === "discarding",
	);

	function updateSession(
		sessionId: string,
		update: (session: SessionState) => SessionState | undefined,
	) {
		setState((current) => ({
			...current,
			sessions: current.sessions.flatMap((session) => {
				if (session.session.sessionId !== sessionId) {
					return [session];
				}

				const updatedSession = update(session);

				return updatedSession ? [updatedSession] : [];
			}),
		}));
	}

	async function startWordEditSession() {
		setState((current) => ({
			...current,
			discardedMessage: undefined,
			errorMessage: undefined,
			startStatus: "starting",
		}));

		try {
			const response = await fetch("/api/word-edit-sessions", {
				method: "POST",
			});

			if (!response.ok) {
				throw new Error(
					`Word編集セッションを開始できませんでした: ${response.status}`,
				);
			}

			const body: unknown = await response.json();

			if (!isWordEditSessionResponse(body)) {
				throw new Error("Word編集セッション response was invalid.");
			}

			setState((current) => ({
				...current,
				sessions: [...current.sessions, { status: "ready", session: body }],
				startStatus: "idle",
			}));
		} catch (error) {
			setState((current) => ({
				...current,
				errorMessage:
					error instanceof Error
						? error.message
						: "Word編集セッションを開始できませんでした。",
				startStatus: "idle",
			}));
		}
	}

	async function finishWordEditSession(session: WordEditSessionResponse) {
		updateSession(session.sessionId, () => ({ status: "finishing", session }));

		try {
			const response = await fetch(
				`/api/word-edit-sessions/${encodeURIComponent(session.sessionId)}/finish`,
				{
					method: "POST",
				},
			);

			if (response.status === 409) {
				const body: unknown = await response.json();

				if (
					isImportErrorResponse(body) &&
					body.session.sessionId === session.sessionId
				) {
					updateSession(session.sessionId, () => ({
						message: body.error.message,
						session,
						status: "importError",
					}));
					return;
				}

				throw new Error("取り込みエラー response was invalid.");
			}

			if (!response.ok) {
				throw new Error(
					`Word編集セッションを終了取り込みできませんでした: ${response.status}`,
				);
			}

			const body: unknown = await response.json();

			if (!isFinishWordEditSessionResponse(body)) {
				throw new Error("終了取り込み response was invalid.");
			}

			updateSession(session.sessionId, () => ({
				notifications: body.notifications,
				session,
				status: "finished",
				webDocument: body.webDocument,
			}));
		} catch (error) {
			updateSession(session.sessionId, () => ({ status: "ready", session }));
			setState((current) => ({
				...current,
				errorMessage:
					error instanceof Error
						? error.message
						: "Word編集セッションを終了取り込みできませんでした。",
			}));
		}
	}

	async function discardWordEditSession(session: WordEditSessionResponse) {
		updateSession(session.sessionId, () => ({ status: "discarding", session }));

		try {
			const response = await fetch(
				`/api/word-edit-sessions/${encodeURIComponent(session.sessionId)}/discard`,
				{
					method: "POST",
				},
			);

			if (!response.ok) {
				throw new Error(
					`Word編集セッションを破棄できませんでした: ${response.status}`,
				);
			}

			const body: unknown = await response.json();

			if (
				!isDiscardWordEditSessionResponse(body) ||
				body.session.sessionId !== session.sessionId
			) {
				throw new Error("Word編集セッション破棄 response was invalid.");
			}

			updateSession(session.sessionId, () => undefined);
			setState((current) => ({
				...current,
				discardedMessage: "Word編集セッションを破棄しました。",
				errorMessage: undefined,
			}));
		} catch (error) {
			updateSession(session.sessionId, () => ({
				message:
					error instanceof Error
						? error.message
						: "Word編集セッションを破棄できませんでした。",
				session,
				status: "importError",
			}));
		}
	}

	return (
		<div className="word-launcher">
			<button
				className="auth-button"
				disabled={state.startStatus === "starting" || sessionActionIsPending}
				onClick={startWordEditSession}
				type="button"
			>
				{state.startStatus === "starting"
					? "OneDrive作業コピーを作成中..."
					: "Word編集セッションを開始"}
			</button>
			{state.sessions.map((sessionState) => (
				<WordEditSessionControls
					key={sessionState.session.sessionId}
					onDiscard={discardWordEditSession}
					onFinish={finishWordEditSession}
					sessionState={sessionState}
				/>
			))}
			{state.errorMessage ? (
				<p className="editor-error">{state.errorMessage}</p>
			) : null}
			{state.discardedMessage ? (
				<p className="editor-status">{state.discardedMessage}</p>
			) : null}
		</div>
	);
}

function WordEditSessionControls(props: {
	onDiscard: (session: WordEditSessionResponse) => void;
	onFinish: (session: WordEditSessionResponse) => void;
	sessionState: SessionState;
}) {
	const { sessionState } = props;
	const { session } = sessionState;

	return (
		<div className="word-launch-links">
			<a className="auth-button" href={session.launchLinks.officeUri}>
				ローカルWordで開く
			</a>
			<a
				className="secondary-button"
				href={session.launchLinks.oneDriveFallbackUrl}
			>
				OneDriveで開く
			</a>
			<span className="editor-status">{session.workingCopy.fileName}</span>
			{sessionState.status === "finished" ? (
				<>
					<span className="editor-status">
						終了取り込みが完了しました。Version{" "}
						{sessionState.webDocument.version}
					</span>
					{sessionState.notifications?.map((notification) => (
						<p className="editor-warning" key={notification.message}>
							{notification.message}
						</p>
					))}
				</>
			) : sessionState.status === "importError" ? (
				<>
					<p className="editor-error">{sessionState.message}</p>
					<button
						className="secondary-button"
						onClick={() => props.onFinish(session)}
						type="button"
					>
						終了取り込みを再試行
					</button>
					<button
						className="secondary-button"
						onClick={() => props.onDiscard(session)}
						type="button"
					>
						Word編集セッションを破棄
					</button>
				</>
			) : sessionState.status === "discarding" ? (
				<span className="editor-status">Word編集セッションを破棄中...</span>
			) : (
				<button
					className="secondary-button"
					disabled={sessionState.status === "finishing"}
					onClick={() => props.onFinish(session)}
					type="button"
				>
					{sessionState.status === "finishing"
						? "終了取り込み中..."
						: "Word編集セッションを終了して取り込む"}
				</button>
			)}
		</div>
	);
}

function isWordEditSessionResponse(
	body: unknown,
): body is WordEditSessionResponse {
	if (typeof body !== "object" || body === null) {
		return false;
	}

	const candidate = body as Record<string, unknown>;

	return (
		typeof candidate.sessionId === "string" &&
		isWorkingCopy(candidate.workingCopy) &&
		isLaunchLinks(candidate.launchLinks)
	);
}

function isFinishWordEditSessionResponse(
	body: unknown,
): body is FinishWordEditSessionResponse {
	if (typeof body !== "object" || body === null) {
		return false;
	}

	const candidate = body as Record<string, unknown>;

	return (
		isWebDocument(candidate.webDocument) &&
		(candidate.notifications === undefined ||
			isImportNotifications(candidate.notifications))
	);
}

function isDiscardWordEditSessionResponse(
	body: unknown,
): body is DiscardWordEditSessionResponse {
	if (typeof body !== "object" || body === null) {
		return false;
	}

	return isDiscardedSession((body as Record<string, unknown>).session);
}

function isImportErrorResponse(body: unknown): body is ImportErrorResponse {
	if (typeof body !== "object" || body === null) {
		return false;
	}

	const candidate = body as Record<string, unknown>;

	return (
		isImportError(candidate.error) && isImportErrorSession(candidate.session)
	);
}

function isImportError(error: unknown): error is ImportErrorResponse["error"] {
	if (typeof error !== "object" || error === null) {
		return false;
	}

	const candidate = error as Record<string, unknown>;

	return (
		candidate.type === "importError" && typeof candidate.message === "string"
	);
}

function isImportNotifications(
	notifications: unknown,
): notifications is NonNullable<
	FinishWordEditSessionResponse["notifications"]
> {
	return (
		Array.isArray(notifications) && notifications.every(isImportNotification)
	);
}

function isImportNotification(
	notification: unknown,
): notification is NonNullable<
	FinishWordEditSessionResponse["notifications"]
>[number] {
	if (typeof notification !== "object" || notification === null) {
		return false;
	}

	const candidate = notification as Record<string, unknown>;

	return (
		candidate.type === "unsupportedContentDiscarded" &&
		typeof candidate.message === "string"
	);
}

function isImportErrorSession(
	session: unknown,
): session is ImportErrorResponse["session"] {
	if (typeof session !== "object" || session === null) {
		return false;
	}

	const candidate = session as Record<string, unknown>;

	return (
		typeof candidate.sessionId === "string" &&
		candidate.status === "importError"
	);
}

function isDiscardedSession(
	session: unknown,
): session is DiscardWordEditSessionResponse["session"] {
	if (typeof session !== "object" || session === null) {
		return false;
	}

	const candidate = session as Record<string, unknown>;

	return (
		typeof candidate.sessionId === "string" && candidate.status === "discarded"
	);
}

function isWebDocument(webDocument: unknown): webDocument is WebDocument {
	if (typeof webDocument !== "object" || webDocument === null) {
		return false;
	}

	const candidate = webDocument as Record<string, unknown>;

	return (
		candidate.id === "shared" &&
		typeof candidate.markdown === "string" &&
		typeof candidate.version === "number"
	);
}

function isWorkingCopy(
	workingCopy: unknown,
): workingCopy is WordEditSessionResponse["workingCopy"] {
	if (typeof workingCopy !== "object" || workingCopy === null) {
		return false;
	}

	const candidate = workingCopy as Record<string, unknown>;

	return (
		typeof candidate.driveItemId === "string" &&
		typeof candidate.fileName === "string"
	);
}

function isLaunchLinks(
	launchLinks: unknown,
): launchLinks is WordEditSessionResponse["launchLinks"] {
	if (typeof launchLinks !== "object" || launchLinks === null) {
		return false;
	}

	const candidate = launchLinks as Record<string, unknown>;

	return (
		typeof candidate.officeUri === "string" &&
		typeof candidate.oneDriveFallbackUrl === "string"
	);
}
