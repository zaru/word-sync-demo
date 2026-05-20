import { getEditorAuthHandlers } from "../../../server/editor-auth";

export const runtime = "nodejs";

export async function GET(request: Request): Promise<Response> {
	return getEditorAuthHandlers().session(request);
}
