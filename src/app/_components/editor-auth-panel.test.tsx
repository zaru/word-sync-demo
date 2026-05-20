// @vitest-environment happy-dom

import "@testing-library/jest-dom/vitest";
import { fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { EditorAuthPanel } from "./editor-auth-panel";

describe("EditorAuthPanel", () => {
	afterEach(() => {
		vi.unstubAllGlobals();
		document.body.replaceChildren();
	});

	it("shows Microsoft login UI before an 編集者 signs in", () => {
		render(<EditorAuthPanel session={{ signedIn: false }} />);

		expect(
			screen.getByRole("link", { name: "Microsoftでサインイン" }),
		).toHaveAttribute("href", "/auth/login");
		expect(
			screen.getByText("Word起動導線はサインイン後に利用できます。"),
		).toBeInTheDocument();
	});

	it("shows logout UI and keeps OneDrive-dependent actions behind the signed-in 編集者 session", () => {
		render(
			<EditorAuthPanel
				session={{
					signedIn: true,
					editor: {
						id: "editor-1",
						displayName: "編集者 A",
						username: "editor@example.com",
					},
				}}
			/>,
		);

		expect(screen.getByText("編集者 A としてサインイン中")).toBeInTheDocument();
		expect(
			screen.getByText(
				"OneDrive作業コピーを使う操作は、この編集者セッションで実行できます。",
			),
		).toBeInTheDocument();
		expect(
			screen.getByRole("button", { name: "ログアウト" }),
		).toBeInTheDocument();
	});

	it("starts a Word編集セッション and exposes the Word起動導線 links", async () => {
		const fetchWordEditSession = vi.fn(async () =>
			Response.json(
				{
					launchLinks: {
						officeUri:
							"ms-word:ofe|u|https://onedrive.example/Webドキュメント-word-session-1.docx",
						oneDriveFallbackUrl:
							"https://onedrive.example/Webドキュメント-word-session-1.docx",
					},
					sessionId: "word-session-1",
					workingCopy: {
						driveItemId: "drive-item-1",
						fileName: "Webドキュメント-word-session-1.docx",
					},
				},
				{ status: 201 },
			),
		);
		vi.stubGlobal("fetch", fetchWordEditSession);
		render(
			<EditorAuthPanel
				session={{
					signedIn: true,
					editor: {
						id: "editor-1",
						displayName: "編集者 A",
						username: "editor@example.com",
					},
				}}
			/>,
		);

		fireEvent.click(
			screen.getByRole("button", { name: "Word編集セッションを開始" }),
		);

		expect(
			await screen.findByRole("link", { name: "ローカルWordで開く" }),
		).toHaveAttribute(
			"href",
			"ms-word:ofe|u|https://onedrive.example/Webドキュメント-word-session-1.docx",
		);
		expect(
			screen.getByRole("link", { name: "OneDriveで開く" }),
		).toHaveAttribute(
			"href",
			"https://onedrive.example/Webドキュメント-word-session-1.docx",
		);
		expect(fetchWordEditSession).toHaveBeenCalledWith(
			"/api/word-edit-sessions",
			{ method: "POST" },
		);
	});

	it("finishes an active Word編集セッション and shows the imported Webドキュメント version", async () => {
		const fetchWordEditSession = vi.fn(
			async (url: RequestInfo | URL, init?: RequestInit) => {
				if (String(url) === "/api/word-edit-sessions") {
					return Response.json(
						{
							launchLinks: {
								officeUri:
									"ms-word:ofe|u|https://onedrive.example/Webドキュメント-word-session-1.docx",
								oneDriveFallbackUrl:
									"https://onedrive.example/Webドキュメント-word-session-1.docx",
							},
							sessionId: "word-session-1",
							workingCopy: {
								driveItemId: "drive-item-1",
								fileName: "Webドキュメント-word-session-1.docx",
							},
						},
						{ status: 201 },
					);
				}

				if (
					String(url) === "/api/word-edit-sessions/word-session-1/finish" &&
					init?.method === "POST"
				) {
					return Response.json({
						webDocument: {
							id: "shared",
							markdown: "# Imported from Word",
							version: 4,
						},
					});
				}

				throw new Error(`Unexpected fetch: ${String(url)}`);
			},
		);
		vi.stubGlobal("fetch", fetchWordEditSession);
		render(
			<EditorAuthPanel
				session={{
					signedIn: true,
					editor: {
						id: "editor-1",
						displayName: "編集者 A",
						username: "editor@example.com",
					},
				}}
			/>,
		);

		fireEvent.click(
			screen.getByRole("button", { name: "Word編集セッションを開始" }),
		);
		fireEvent.click(
			await screen.findByRole("button", {
				name: "Word編集セッションを終了して取り込む",
			}),
		);

		expect(
			await screen.findByText("終了取り込みが完了しました。Version 4"),
		).toBeInTheDocument();
		expect(fetchWordEditSession).toHaveBeenCalledWith(
			"/api/word-edit-sessions/word-session-1/finish",
			{ method: "POST" },
		);
	});

	it("shows 互換外破棄通知 after a successful 終了取り込み with discarded unsupported content", async () => {
		const notificationMessage =
			"基本Markdown要素として取り込めないWord編集を破棄しました。";
		const fetchWordEditSession = vi.fn(
			async (url: RequestInfo | URL, init?: RequestInit) => {
				if (String(url) === "/api/word-edit-sessions") {
					return Response.json(
						{
							launchLinks: {
								officeUri:
									"ms-word:ofe|u|https://onedrive.example/Webドキュメント-word-session-1.docx",
								oneDriveFallbackUrl:
									"https://onedrive.example/Webドキュメント-word-session-1.docx",
							},
							sessionId: "word-session-1",
							workingCopy: {
								driveItemId: "drive-item-1",
								fileName: "Webドキュメント-word-session-1.docx",
							},
						},
						{ status: 201 },
					);
				}

				if (
					String(url) === "/api/word-edit-sessions/word-session-1/finish" &&
					init?.method === "POST"
				) {
					return Response.json({
						notifications: [
							{
								message: notificationMessage,
								type: "unsupportedContentDiscarded",
							},
						],
						webDocument: {
							id: "shared",
							markdown: "# Imported from Word",
							version: 4,
						},
					});
				}

				throw new Error(`Unexpected fetch: ${String(url)}`);
			},
		);
		vi.stubGlobal("fetch", fetchWordEditSession);
		render(
			<EditorAuthPanel
				session={{
					signedIn: true,
					editor: {
						id: "editor-1",
						displayName: "編集者 A",
						username: "editor@example.com",
					},
				}}
			/>,
		);

		fireEvent.click(
			screen.getByRole("button", { name: "Word編集セッションを開始" }),
		);
		fireEvent.click(
			await screen.findByRole("button", {
				name: "Word編集セッションを終了して取り込む",
			}),
		);

		expect(await screen.findByText(notificationMessage)).toBeInTheDocument();
	});

	it("keeps an errored Word編集セッション visible so the 編集者 can retry 終了取り込み", async () => {
		let finishAttempts = 0;
		const fetchWordEditSession = vi.fn(
			async (url: RequestInfo | URL, init?: RequestInit) => {
				if (String(url) === "/api/word-edit-sessions") {
					return Response.json(
						{
							launchLinks: {
								officeUri:
									"ms-word:ofe|u|https://onedrive.example/Webドキュメント-word-session-1.docx",
								oneDriveFallbackUrl:
									"https://onedrive.example/Webドキュメント-word-session-1.docx",
							},
							sessionId: "word-session-1",
							workingCopy: {
								driveItemId: "drive-item-1",
								fileName: "Webドキュメント-word-session-1.docx",
							},
						},
						{ status: 201 },
					);
				}

				if (
					String(url) === "/api/word-edit-sessions/word-session-1/finish" &&
					init?.method === "POST"
				) {
					finishAttempts += 1;

					if (finishAttempts === 1) {
						return Response.json(
							{
								error: {
									message: "OneDrive作業コピー was missing or deleted.",
									type: "importError",
								},
								session: {
									sessionId: "word-session-1",
									status: "importError",
								},
							},
							{ status: 409 },
						);
					}

					return Response.json({
						webDocument: {
							id: "shared",
							markdown: "# Imported from Word",
							version: 4,
						},
					});
				}

				throw new Error(`Unexpected fetch: ${String(url)}`);
			},
		);
		vi.stubGlobal("fetch", fetchWordEditSession);
		render(
			<EditorAuthPanel
				session={{
					signedIn: true,
					editor: {
						id: "editor-1",
						displayName: "編集者 A",
						username: "editor@example.com",
					},
				}}
			/>,
		);

		fireEvent.click(
			screen.getByRole("button", { name: "Word編集セッションを開始" }),
		);
		fireEvent.click(
			await screen.findByRole("button", {
				name: "Word編集セッションを終了して取り込む",
			}),
		);

		expect(
			await screen.findByText("OneDrive作業コピー was missing or deleted."),
		).toBeInTheDocument();
		expect(
			screen.getByRole("button", { name: "終了取り込みを再試行" }),
		).toBeInTheDocument();
		expect(
			screen.getByRole("button", { name: "Word編集セッションを破棄" }),
		).toBeInTheDocument();

		fireEvent.click(
			screen.getByRole("button", { name: "終了取り込みを再試行" }),
		);

		expect(
			await screen.findByText("終了取り込みが完了しました。Version 4"),
		).toBeInTheDocument();
		expect(finishAttempts).toBe(2);
	});

	it("lets the 編集者 discard an errored Word編集セッション", async () => {
		const fetchWordEditSession = vi.fn(
			async (url: RequestInfo | URL, init?: RequestInit) => {
				if (String(url) === "/api/word-edit-sessions") {
					return Response.json(
						{
							launchLinks: {
								officeUri:
									"ms-word:ofe|u|https://onedrive.example/Webドキュメント-word-session-1.docx",
								oneDriveFallbackUrl:
									"https://onedrive.example/Webドキュメント-word-session-1.docx",
							},
							sessionId: "word-session-1",
							workingCopy: {
								driveItemId: "drive-item-1",
								fileName: "Webドキュメント-word-session-1.docx",
							},
						},
						{ status: 201 },
					);
				}

				if (
					String(url) === "/api/word-edit-sessions/word-session-1/finish" &&
					init?.method === "POST"
				) {
					return Response.json(
						{
							error: {
								message: "OneDrive作業コピー was missing or deleted.",
								type: "importError",
							},
							session: {
								sessionId: "word-session-1",
								status: "importError",
							},
						},
						{ status: 409 },
					);
				}

				if (
					String(url) === "/api/word-edit-sessions/word-session-1/discard" &&
					init?.method === "POST"
				) {
					return Response.json({
						session: {
							sessionId: "word-session-1",
							status: "discarded",
						},
					});
				}

				throw new Error(`Unexpected fetch: ${String(url)}`);
			},
		);
		vi.stubGlobal("fetch", fetchWordEditSession);
		render(
			<EditorAuthPanel
				session={{
					signedIn: true,
					editor: {
						id: "editor-1",
						displayName: "編集者 A",
						username: "editor@example.com",
					},
				}}
			/>,
		);

		fireEvent.click(
			screen.getByRole("button", { name: "Word編集セッションを開始" }),
		);
		fireEvent.click(
			await screen.findByRole("button", {
				name: "Word編集セッションを終了して取り込む",
			}),
		);
		fireEvent.click(
			await screen.findByRole("button", {
				name: "Word編集セッションを破棄",
			}),
		);

		expect(
			await screen.findByText("Word編集セッションを破棄しました。"),
		).toBeInTheDocument();
		expect(fetchWordEditSession).toHaveBeenCalledWith(
			"/api/word-edit-sessions/word-session-1/discard",
			{ method: "POST" },
		);
	});
});
