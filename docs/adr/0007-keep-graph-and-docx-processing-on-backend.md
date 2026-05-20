# Keep Graph and DOCX processing on the backend

Graph access, OAuth token handling, DOCX export, and **終了取り込み** will run on the backend rather than in the browser. The app will request `offline_access` so the backend can refresh Graph access during long **Word編集セッション**, keeping Microsoft tokens and document conversion behavior in one controlled boundary while the Lexical UI remains a projection over the canonical Markdown **Webドキュメント**.
