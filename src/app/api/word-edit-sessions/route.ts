import { getWordEditSessionHandlers } from "../../../server/word-edit-session";

export const runtime = "nodejs";

export async function POST(request: Request): Promise<Response> {
	return getWordEditSessionHandlers().start(request);
}
