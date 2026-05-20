import type { EditorProfile } from "../../server/editor-auth-store";
import { WordEditSessionLauncher } from "./word-edit-session-launcher";

export type BrowserEditorSession =
	| { signedIn: false }
	| {
			signedIn: true;
			editor: EditorProfile;
	  };

type EditorAuthPanelProps = {
	session: BrowserEditorSession;
};

export function EditorAuthPanel({ session }: EditorAuthPanelProps) {
	if (!session.signedIn) {
		return (
			<section className="auth-card" aria-label="Microsoft sign-in">
				<div>
					<p className="eyebrow">Microsoft account</p>
					<h2>編集者としてサインイン</h2>
					<p>Word起動導線はサインイン後に利用できます。</p>
				</div>
				<a className="auth-button" href="/auth/login">
					Microsoftでサインイン
				</a>
			</section>
		);
	}

	const label =
		session.editor.displayName ?? session.editor.username ?? "編集者";

	return (
		<section className="auth-card" aria-label="Signed-in editor">
			<div>
				<p className="eyebrow">Microsoft account</p>
				<h2>{label} としてサインイン中</h2>
				{session.editor.username ? <p>{session.editor.username}</p> : null}
				<p>
					OneDrive作業コピーを使う操作は、この編集者セッションで実行できます。
				</p>
			</div>
			<div className="auth-actions">
				<WordEditSessionLauncher />
				<form action="/auth/logout" method="post">
					<button className="secondary-button" type="submit">
						ログアウト
					</button>
				</form>
			</div>
		</section>
	);
}
