import { cookies } from "next/headers";
import { readEditorSessionFromCookieValue } from "../server/editor-auth";
import { editorSessionCookieName } from "../server/editor-auth-routes";
import { getWebDocumentStore } from "../server/web-document-store";
import {
	type BrowserEditorSession,
	EditorAuthPanel,
} from "./_components/editor-auth-panel";
import { WebDocumentEditor } from "./_components/web-document-editor";

export const dynamic = "force-dynamic";

export default async function Home() {
	const document = getWebDocumentStore().loadSharedDocument();
	const cookieStore = await cookies();
	const editorSession = readEditorSessionFromCookieValue(
		cookieStore.get(editorSessionCookieName)?.value,
	);
	const browserSession: BrowserEditorSession = editorSession
		? {
				signedIn: true,
				editor: editorSession.editor,
			}
		: { signedIn: false };

	return (
		<main className="page-shell">
			<section className="hero">
				<p className="eyebrow">Single shared Webドキュメント</p>
				<h1>Markdown互換内容 editor</h1>
				<p>
					SQLiteに保存された単一のWebドキュメントをLexical上の投影として編集します。
					変更は短い待ち時間のあと自動保存され、別タブの更新はバージョンポーリングで反映されます。
				</p>
			</section>
			<EditorAuthPanel session={browserSession} />
			<WebDocumentEditor initialDocument={document} />
		</main>
	);
}
