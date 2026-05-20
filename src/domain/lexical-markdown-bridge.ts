import { CodeHighlightNode, CodeNode } from "@lexical/code";
import { createHeadlessEditor } from "@lexical/headless";
import { LinkNode } from "@lexical/link";
import { ListItemNode, ListNode } from "@lexical/list";
import {
	$convertFromMarkdownString,
	$convertToMarkdownString,
	TRANSFORMERS,
	type Transformer,
} from "@lexical/markdown";
import { HeadingNode, QuoteNode } from "@lexical/rich-text";
import type { Klass, LexicalNode } from "lexical";

export const lexicalMarkdownNodes: Array<Klass<LexicalNode>> = [
	HeadingNode,
	QuoteNode,
	ListNode,
	ListItemNode,
	LinkNode,
	CodeNode,
	CodeHighlightNode,
];

export const markdownTransformers: Transformer[] = TRANSFORMERS;

export async function markdownToLexicalState(
	markdown: string,
): Promise<string> {
	const editor = createBridgeEditor();

	editor.update(
		() => {
			$convertFromMarkdownString(markdown, markdownTransformers);
		},
		{ discrete: true },
	);

	return JSON.stringify(editor.getEditorState().toJSON());
}

export async function lexicalStateToMarkdown(
	editorStateJson: string,
): Promise<string> {
	const editor = createBridgeEditor();
	const editorState = editor.parseEditorState(editorStateJson);

	editor.setEditorState(editorState);

	let markdown = "";
	editor.getEditorState().read(() => {
		markdown = $convertToMarkdownString(markdownTransformers);
	});

	return markdown.trimEnd();
}

function createBridgeEditor() {
	return createHeadlessEditor({
		namespace: "LexicalMarkdownBridge",
		nodes: lexicalMarkdownNodes,
		onError(error) {
			throw error;
		},
	});
}
