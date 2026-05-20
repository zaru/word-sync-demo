import { describe, expect, it, vi } from "vitest";

const discard = vi.hoisted(() =>
	vi.fn(async () =>
		Response.json({
			session: {
				sessionId: "word-session-1",
				status: "discarded",
			},
		}),
	),
);

vi.mock("../../../../../server/word-edit-session", () => ({
	getWordEditSessionHandlers() {
		return { discard };
	},
}));

describe("/api/word-edit-sessions/{sessionId}/discard", () => {
	it("discards the requested errored Word編集セッション", async () => {
		const { POST } = await import("./route");
		const request = new Request(
			"http://localhost/api/word-edit-sessions/word-session-1/discard",
			{ method: "POST" },
		);

		const response = await POST(request, {
			params: Promise.resolve({ sessionId: "word-session-1" }),
		});

		expect(discard).toHaveBeenCalledWith(request, {
			sessionId: "word-session-1",
		});
		expect(response.status).toBe(200);
		await expect(response.json()).resolves.toEqual({
			session: {
				sessionId: "word-session-1",
				status: "discarded",
			},
		});
	});
});
