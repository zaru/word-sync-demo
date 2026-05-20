import { describe, expect, it, vi } from "vitest";

const cleanup = vi.hoisted(() =>
	vi.fn(async () =>
		Response.json({
			abandonedSessions: ["word-session-1"],
			deletedWorkingCopies: [
				{
					driveItemId: "drive-item-1",
					sessionId: "word-session-2",
				},
			],
			failures: [],
		}),
	),
);

vi.mock("../../../../server/word-edit-session", () => ({
	getWordEditSessionHandlers() {
		return { cleanup };
	},
}));

describe("/api/word-edit-sessions/cleanup", () => {
	it("runs Word編集セッション lifecycle cleanup", async () => {
		const { POST } = await import("./route");

		const response = await POST();

		expect(cleanup).toHaveBeenCalledWith();
		expect(response.status).toBe(200);
		await expect(response.json()).resolves.toEqual({
			abandonedSessions: ["word-session-1"],
			deletedWorkingCopies: [
				{
					driveItemId: "drive-item-1",
					sessionId: "word-session-2",
				},
			],
			failures: [],
		});
	});
});
