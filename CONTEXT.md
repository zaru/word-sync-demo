# Word Sync Demo

Web上の編集体験とローカルWord編集体験をつなぐデモの文脈。Markdownで表現できる範囲だけを共有可能な文書内容として扱う。

## Language

**Webドキュメント**:
Webアプリが保持する正本の文書。
_Avoid_: ドキュメントデータ, 元データ, 正式版

**Markdown互換内容**:
Markdownで表現できる範囲に制限された共有可能な文書内容。
_Avoid_: 本文, コンテンツ, Word内容

**基本Markdown要素**:
段落、見出し、強調、リンク、リスト、引用、インラインコード、コードブロックで構成される文書要素。
_Avoid_: Markdown全部, GFM, リッチテキスト

**編集者**:
WebドキュメントをWordで編集するサインイン済み利用者。
_Avoid_: ユーザー, アカウント, Microsoftユーザー

**OneDrive作業コピー**:
Wordで編集するために正本から作られる一時的な文書コピー。
_Avoid_: OneDriveドキュメント, 同期ファイル, Wordファイル

**編集者OneDrive領域**:
Wordで編集する本人が所有するOneDrive上の作業場所。
_Avoid_: OneDrive領域, 保存先, クラウド

**Word編集セッション**:
ユーザーがWebドキュメントをWordで開いてから、変更がWebドキュメントへ取り込まれるまでの一連の編集。
_Avoid_: Word連携, 外部編集, 開く処理

**セッション終了**:
編集者の終了操作または無更新時間によってWord編集セッションを閉じたものとして扱う状態。
_Avoid_: Word終了, ファイルクローズ, 同期完了

**放置終了**:
無更新時間によってWebドキュメントを更新せずにWord編集セッションを閉じる状態。
_Avoid_: 自動取り込み, タイムアウト同期, 期限切れ保存

**最後取り込み優先**:
最後に終了取り込みが成功したWord編集セッションのMarkdown互換内容でWebドキュメントを置き換える競合解決。
_Avoid_: 最後保存優先, 自動マージ, 上書き

**終了取り込み**:
編集者の終了操作後にOneDrive作業コピーの変更が落ち着いてから最新内容をWebドキュメントへ反映する取り込み。
_Avoid_: 即時同期, 保存反映, 自動取り込み

**取り込みエラー**:
OneDrive作業コピーを取得または変換できないためWebドキュメントを更新しないWord編集セッションの状態。
_Avoid_: 同期失敗, 変換失敗, 自動破棄

**互換外破棄通知**:
基本Markdown要素として取り込めないWord編集が破棄されたことを編集者へ知らせる通知。
_Avoid_: エラー, 変換警告, 失敗

**作業コピー削除猶予**:
セッション終了後もしばらくOneDrive作業コピーを残す削除待ち期間。
_Avoid_: 即時削除, アーカイブ, 保存期間

**Word編集環境**:
編集者本人のOneDriveにある作業コピーをローカルWordで編集できる端末状態。
_Avoid_: ローカル環境, クライアント, Word環境

**Word起動導線**:
編集者がOneDrive作業コピーをローカルWordで開くためのWebアプリ上の操作。
_Avoid_: Wordで開くボタン, 起動URL, 外部リンク

## Relationships

- A **Webドキュメント** may have zero or more active **Word編集セッション**
- A **Webドキュメント** contains **Markdown互換内容**
- A **Webドキュメント** is shared by zero or more **編集者**
- **Markdown互換内容** is limited to **基本Markdown要素**
- A **編集者** owns exactly one **編集者OneDrive領域**
- A **Word編集セッション** uses exactly one **OneDrive作業コピー**
- A **Word編集セッション** requires a **Word編集環境**
- A **Word編集セッション** starts from a **Word起動導線**
- A **Word編集セッション** is reconciled by **終了取り込み**
- A **Word編集セッション** ends with **セッション終了**
- A **Word編集セッション** may end as **放置終了**
- A **Word編集セッション** may enter **取り込みエラー**
- **最後取り込み優先** resolves concurrent **Word編集セッション** imports for one **Webドキュメント**
- **互換外破棄通知** may be shown after **終了取り込み**
- A **OneDrive作業コピー** is derived from exactly one **Webドキュメント**
- A **OneDrive作業コピー** belongs to exactly one **Word編集セッション**
- A **OneDrive作業コピー** belongs to exactly one **編集者OneDrive領域**
- A **OneDrive作業コピー** enters **作業コピー削除猶予** after **セッション終了**

## Example dialogue

> **Dev:** "When the user edits in Word, does the **OneDrive作業コピー** become the source of truth?"
> **Domain expert:** "No — it is only an editable copy during the **Word編集セッション**. The **Webドキュメント** remains the source of truth."

## Flagged ambiguities

- "ドキュメントデータ" was used ambiguously for both the Web-side source and the Word-editable file — resolved: **Webドキュメント** is the source of truth, and **OneDrive作業コピー** is temporary.
- "Wordデータ" was used broadly, but this demo only imports **Markdown互換内容** from Word edits.
- "Markdownで表現できるもの" was narrowed to **基本Markdown要素**, excluding tables, images, footnotes, and embedded HTML.
- Unsupported Word content is not an import failure — resolved: show **互換外破棄通知** and keep the import successful.
- A missing, inaccessible, or unconvertible working copy is an **取り込みエラー**, not an update to the **Webドキュメント**.
- "Wordを閉じる" is not an observable domain event — resolved: **セッション終了** is explicit or timeout-based.
- Timeout does not import Word changes — resolved: **放置終了** leaves the **Webドキュメント** unchanged.
- Changes to a working copy after **セッション終了** are ignored by the **Webドキュメント**.
- Concurrent Word edits are not merged — resolved: **最後取り込み優先** defines the conflict behavior.
- Concurrent Web edits are not protected by confirmation — resolved: **最後取り込み優先** may replace them.
- "保存したら反映" was narrowed to **終了取り込み**, not push notification or polling.
