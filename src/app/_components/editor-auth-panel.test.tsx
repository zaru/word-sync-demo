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
});
