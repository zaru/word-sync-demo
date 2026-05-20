import { getWebDocumentStore } from '../../../server/web-document-store';

export const runtime = 'nodejs';

export async function GET(): Promise<Response> {
  return Response.json(getWebDocumentStore().loadSharedDocument());
}

export async function PUT(request: Request): Promise<Response> {
  const body = await request.json();

  if (!isSaveWebDocumentRequest(body)) {
    return Response.json({ error: 'markdown must be a string' }, { status: 400 });
  }

  return Response.json(getWebDocumentStore().saveMarkdown(body.markdown));
}

function isSaveWebDocumentRequest(body: unknown): body is { markdown: string } {
  return (
    typeof body === 'object' &&
    body !== null &&
    typeof (body as Record<string, unknown>).markdown === 'string'
  );
}
