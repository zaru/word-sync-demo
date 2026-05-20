import { WebDocumentEditor } from './web-document-editor';
import { getWebDocumentStore } from '../server/web-document-store';

export const dynamic = 'force-dynamic';

export default function Home() {
  const document = getWebDocumentStore().loadSharedDocument();

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
      <WebDocumentEditor initialDocument={document} />
    </main>
  );
}
