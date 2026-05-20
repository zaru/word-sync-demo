import { join } from "node:path";

import { WebDocumentStore } from "../domain/web-document-store";

export const defaultSeedMarkdown = [
	"# Webドキュメント",
	"",
	"この共有WebドキュメントはMarkdown互換内容として保存されます。",
	"",
	"- 基本Markdown要素を編集できます",
	"- 変更は短い待ち時間のあと自動保存されます",
].join("\n");

let cachedStore: { databasePath: string; store: WebDocumentStore } | undefined;

export function getWebDocumentStore(): WebDocumentStore {
	const databasePath =
		process.env.WEB_DOCUMENT_DB_PATH ??
		join(process.cwd(), ".data", "web-document.sqlite");

	if (cachedStore?.databasePath !== databasePath) {
		cachedStore?.store.close();
		cachedStore = {
			databasePath,
			store: new WebDocumentStore({
				databasePath,
				seedMarkdown: defaultSeedMarkdown,
			}),
		};
	}

	return cachedStore.store;
}
