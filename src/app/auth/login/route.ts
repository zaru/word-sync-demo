import { getEditorAuthHandlers } from '../../../server/editor-auth';

export const runtime = 'nodejs';

export async function GET(): Promise<Response> {
  return getEditorAuthHandlers().login();
}
