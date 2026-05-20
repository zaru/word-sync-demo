# Store Markdown as the canonical web document content

The **Webドキュメント** will store canonical content as Markdown, while Lexical state is treated as an editor projection that autosaves Markdown changes with a short debounce. This keeps the DOCX import/export boundary aligned with **Markdown互換内容** and makes it explicit that Word-only structures are discarded instead of being preserved in hidden Lexical or DOCX state.
