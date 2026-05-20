import { getEditorAuthHandlers } from "../../../server/editor-auth";

export const runtime = "nodejs";

export async function POST(request: Request): Promise<Response> {
	return getEditorAuthHandlers().logout(request);
}
