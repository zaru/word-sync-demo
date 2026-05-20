// @vitest-environment happy-dom

import "@testing-library/jest-dom/vitest";
import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";

import { EditorAuthPanel } from "./editor-auth-panel";

describe("EditorAuthPanel", () => {
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
});
