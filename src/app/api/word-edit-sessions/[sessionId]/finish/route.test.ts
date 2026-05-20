import { describe, expect, it, vi } from "vitest";

const finish = vi.hoisted(() =>
	vi.fn(async () =>
		Response.json({
			webDocument: {
				id: "shared",
				markdown: "# Imported from Word",
				version: 4,
			},
		}),
	),
);

vi.mock("../../../../../server/word-edit-session", () => ({
	getWordEditSessionHandlers() {
		return { finish };
	},
}));

describe("/api/word-edit-sessions/{sessionId}/finish", () => {
	it("finishes the requested Word編集セッション and returns the imported Webドキュメント", async () => {
		const { POST } = await import("./route");
		const request = new Request(
			"http://localhost/api/word-edit-sessions/word-session-1/finish",
			{ method: "POST" },
		);

		const response = await POST(request, {
			params: Promise.resolve({ sessionId: "word-session-1" }),
		});

		expect(finish).toHaveBeenCalledWith(request, {
			sessionId: "word-session-1",
		});
		expect(response.status).toBe(200);
		await expect(response.json()).resolves.toEqual({
			webDocument: {
				id: "shared",
				markdown: "# Imported from Word",
				version: 4,
			},
		});
	});
});
