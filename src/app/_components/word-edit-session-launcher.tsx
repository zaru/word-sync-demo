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
	webDocument: WebDocument;
};

type LaunchState =
	| { status: "idle" }
	| { status: "starting" }
	| { status: "ready"; session: WordEditSessionResponse }
	| { status: "finishing"; session: WordEditSessionResponse }
	| {
			status: "finished";
			session: WordEditSessionResponse;
			webDocument: WebDocument;
	  }
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

			if (!response.ok) {
				throw new Error(
					`Word編集セッションを終了取り込みできませんでした: ${response.status}`,
				);
			}

			const body: unknown = await response.json();

			if (!isFinishWordEditSessionResponse(body)) {
				throw new Error("終了取り込み response was invalid.");
			}

			setState({ status: "finished", session, webDocument: body.webDocument });
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

	return (
		<div className="word-launcher">
			<button
				className="auth-button"
				disabled={state.status === "starting" || state.status === "finishing"}
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
						<span className="editor-status">
							終了取り込みが完了しました。Version {state.webDocument.version}
						</span>
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
		</div>
	);
}

function sessionFromState(
	state: LaunchState,
): WordEditSessionResponse | undefined {
	if (
		state.status === "ready" ||
		state.status === "finishing" ||
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

	return isWebDocument((body as Record<string, unknown>).webDocument);
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
