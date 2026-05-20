import { getWordEditSessionHandlers } from "../../../../server/word-edit-session";

export const runtime = "nodejs";

export async function POST(): Promise<Response> {
	return getWordEditSessionHandlers().cleanup();
}
