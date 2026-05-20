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

type LaunchState =
	| { status: "idle" }
	| { status: "starting" }
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
	  }
	| { status: "discarded" }
	| { status: "error"; message: string };

export function WordEditSessionLauncher() {
	const [state, setState] = useState<LaunchState>({ status: "idle" });
	const activeSession = sessionFromState(state);

	async function startWordEditSession() {
		setState({ status: "starting" });

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

			setState({ status: "ready", session: body });
		} catch (error) {
			setState({
				status: "error",
				message:
					error instanceof Error
						? error.message
						: "Word編集セッションを開始できませんでした。",
			});
		}
	}

	async function finishWordEditSession(session: WordEditSessionResponse) {
		setState({ status: "finishing", session });

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
					setState({
						message: body.error.message,
						session,
						status: "importError",
					});
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

			setState({
				notifications: body.notifications,
				session,
				status: "finished",
				webDocument: body.webDocument,
			});
		} catch (error) {
			setState({
				status: "error",
				message:
					error instanceof Error
						? error.message
						: "Word編集セッションを終了取り込みできませんでした。",
			});
		}
	}

	async function discardWordEditSession(session: WordEditSessionResponse) {
		setState({ status: "discarding", session });

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

			setState({ status: "discarded" });
		} catch (error) {
			setState({
				message:
					error instanceof Error
						? error.message
						: "Word編集セッションを破棄できませんでした。",
				session,
				status: "importError",
			});
		}
	}

	return (
		<div className="word-launcher">
			<button
				className="auth-button"
				disabled={
					state.status === "starting" ||
					state.status === "finishing" ||
					state.status === "discarding"
				}
				onClick={startWordEditSession}
				type="button"
			>
				{state.status === "starting"
					? "OneDrive作業コピーを作成中..."
					: "Word編集セッションを開始"}
			</button>
			{activeSession ? (
				<div className="word-launch-links">
					<a className="auth-button" href={activeSession.launchLinks.officeUri}>
						ローカルWordで開く
					</a>
					<a
						className="secondary-button"
						href={activeSession.launchLinks.oneDriveFallbackUrl}
					>
						OneDriveで開く
					</a>
					<span className="editor-status">
						{activeSession.workingCopy.fileName}
					</span>
					{state.status === "finished" ? (
						<>
							<span className="editor-status">
								終了取り込みが完了しました。Version {state.webDocument.version}
							</span>
							{state.notifications?.map((notification) => (
								<p className="editor-warning" key={notification.message}>
									{notification.message}
								</p>
							))}
						</>
					) : state.status === "importError" ? (
						<>
							<p className="editor-error">{state.message}</p>
							<button
								className="secondary-button"
								onClick={() => finishWordEditSession(activeSession)}
								type="button"
							>
								終了取り込みを再試行
							</button>
							<button
								className="secondary-button"
								onClick={() => discardWordEditSession(activeSession)}
								type="button"
							>
								Word編集セッションを破棄
							</button>
						</>
					) : state.status === "discarding" ? (
						<span className="editor-status">Word編集セッションを破棄中...</span>
					) : (
						<button
							className="secondary-button"
							disabled={state.status === "finishing"}
							onClick={() => finishWordEditSession(activeSession)}
							type="button"
						>
							{state.status === "finishing"
								? "終了取り込み中..."
								: "Word編集セッションを終了して取り込む"}
						</button>
					)}
				</div>
			) : null}
			{state.status === "error" ? (
				<p className="editor-error">{state.message}</p>
			) : null}
			{state.status === "discarded" ? (
				<p className="editor-status">Word編集セッションを破棄しました。</p>
			) : null}
		</div>
	);
}

function sessionFromState(
	state: LaunchState,
): WordEditSessionResponse | undefined {
	if (
		state.status === "ready" ||
		state.status === "finishing" ||
		state.status === "discarding" ||
		state.status === "importError" ||
		state.status === "finished"
	) {
		return state.session;
	}

	return undefined;
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
