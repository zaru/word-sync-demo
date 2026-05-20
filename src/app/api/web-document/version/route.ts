import { getWebDocumentStore } from '../../../../server/web-document-store';

export const runtime = 'nodejs';

export async function GET(): Promise<Response> {
  return Response.json({ version: getWebDocumentStore().readVersion() });
}
