import { getWordEditSessionHandlers } from "../../../../../server/word-edit-session";

export const runtime = "nodejs";

export async function POST(
	request: Request,
	context: { params: Promise<{ sessionId: string }> },
): Promise<Response> {
	const params = await context.params;

	return getWordEditSessionHandlers().discard(request, {
		sessionId: params.sessionId,
	});
}
