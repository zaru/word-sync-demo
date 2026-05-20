"use client";

import { useState } from "react";

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

type LaunchState =
	| { status: "idle" }
	| { status: "starting" }
	| { status: "ready"; session: WordEditSessionResponse }
	| { status: "error"; message: string };

export function WordEditSessionLauncher() {
	const [state, setState] = useState<LaunchState>({ status: "idle" });

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

	return (
		<div className="word-launcher">
			<button
				className="auth-button"
				disabled={state.status === "starting"}
				onClick={startWordEditSession}
				type="button"
			>
				{state.status === "starting"
					? "OneDrive作業コピーを作成中..."
					: "Word編集セッションを開始"}
			</button>
			{state.status === "ready" ? (
				<div className="word-launch-links">
					<a className="auth-button" href={state.session.launchLinks.officeUri}>
						ローカルWordで開く
					</a>
					<a
						className="secondary-button"
						href={state.session.launchLinks.oneDriveFallbackUrl}
					>
						OneDriveで開く
					</a>
					<span className="editor-status">
						{state.session.workingCopy.fileName}
					</span>
				</div>
			) : null}
			{state.status === "error" ? (
				<p className="editor-error">{state.message}</p>
			) : null}
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
