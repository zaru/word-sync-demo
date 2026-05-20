"use client";

import {
	$convertFromMarkdownString,
	$convertToMarkdownString,
} from "@lexical/markdown";
import { AutoFocusPlugin } from "@lexical/react/LexicalAutoFocusPlugin";
import { LexicalComposer } from "@lexical/react/LexicalComposer";
import { useLexicalComposerContext } from "@lexical/react/LexicalComposerContext";
import { ContentEditable } from "@lexical/react/LexicalContentEditable";
import { LexicalErrorBoundary } from "@lexical/react/LexicalErrorBoundary";
import { HistoryPlugin } from "@lexical/react/LexicalHistoryPlugin";
import { LinkPlugin } from "@lexical/react/LexicalLinkPlugin";
import { ListPlugin } from "@lexical/react/LexicalListPlugin";
import { MarkdownShortcutPlugin } from "@lexical/react/LexicalMarkdownShortcutPlugin";
import { OnChangePlugin } from "@lexical/react/LexicalOnChangePlugin";
import { RichTextPlugin } from "@lexical/react/LexicalRichTextPlugin";
import type { EditorState } from "lexical";
import {
	type MutableRefObject,
	useEffect,
	useMemo,
	useRef,
	useState,
} from "react";
import {
	lexicalMarkdownNodes,
	markdownTransformers,
} from "../../domain/lexical-markdown-bridge";
import type { WebDocument } from "../../domain/web-document-store";
import {
	createDebouncedAutosave,
	loadChangedWebDocument,
} from "../_lib/web-document-client";

const autosaveDelayMs = 700;
const pollingIntervalMs = 3_000;

type SaveStatus = "saved" | "saving" | "reloaded" | "error";

type WebDocumentEditorProps = {
	initialDocument: WebDocument;
};

export function WebDocumentEditor({ initialDocument }: WebDocumentEditorProps) {
	const [currentVersion, setCurrentVersion] = useState(initialDocument.version);
	const [status, setStatus] = useState<SaveStatus>("saved");
	const [errorMessage, setErrorMessage] = useState<string | undefined>();
	const [isEditorMounted, setIsEditorMounted] = useState(false);
	const lastMarkdownRef = useRef(initialDocument.markdown);

	const autosave = useMemo(
		() =>
			createDebouncedAutosave({
				delayMs: autosaveDelayMs,
				saveMarkdown,
				onSaved(savedDocument) {
					lastMarkdownRef.current = savedDocument.markdown;
					setCurrentVersion(savedDocument.version);
					setStatus("saved");
					setErrorMessage(undefined);
				},
				onError(error) {
					setStatus("error");
					setErrorMessage(
						error instanceof Error ? error.message : "Autosave failed.",
					);
				},
			}),
		[],
	);

	useEffect(() => {
		setIsEditorMounted(true);
	}, []);

	useEffect(() => autosave.cancel, [autosave]);

	const initialConfig = useMemo(
		() => ({
			namespace: "SharedWebDocumentEditor",
			nodes: lexicalMarkdownNodes,
			theme: editorTheme,
			onError(error: Error) {
				throw error;
			},
			editorState() {
				$convertFromMarkdownString(
					initialDocument.markdown,
					markdownTransformers,
				);
			},
		}),
		[initialDocument.markdown],
	);

	return (
		<section className="editor-card" aria-label="Webドキュメント editor">
			<div className="editor-toolbar">
				<strong>Version {currentVersion}</strong>
				<span className={status === "error" ? "editor-error" : "editor-status"}>
					{statusLabel(status, errorMessage)}
				</span>
			</div>
			{isEditorMounted ? (
				<LexicalComposer initialConfig={initialConfig}>
					<div className="editor-shell">
						<RichEditor />
						<HistoryPlugin />
						<ListPlugin />
						<LinkPlugin />
						<MarkdownShortcutPlugin transformers={markdownTransformers} />
						<AutoFocusPlugin />
						<OnChangePlugin
							ignoreSelectionChange
							onChange={(editorState) => {
								handleEditorChange(
									editorState,
									lastMarkdownRef,
									autosave,
									setStatus,
								);
							}}
						/>
						<VersionPollingPlugin
							autosave={autosave}
							currentVersion={currentVersion}
							lastMarkdownRef={lastMarkdownRef}
							onReload={(document) => {
								setCurrentVersion(document.version);
								setStatus("reloaded");
								setErrorMessage(undefined);
							}}
							onError={(error) => {
								setStatus("error");
								setErrorMessage(
									error instanceof Error ? error.message : "Polling failed.",
								);
							}}
						/>
					</div>
				</LexicalComposer>
			) : (
				<div className="editor-shell" aria-hidden="true" />
			)}
		</section>
	);
}

function RichEditor() {
	return (
		<RichTextPlugin
			contentEditable={
				<ContentEditable
					aria-label="Markdown互換内容"
					className="editor-input"
				/>
			}
			placeholder={
				<div className="editor-placeholder">Markdown互換内容を入力...</div>
			}
			ErrorBoundary={LexicalErrorBoundary}
		/>
	);
}

function VersionPollingPlugin(props: {
	autosave: ReturnType<typeof createDebouncedAutosave>;
	currentVersion: number;
	lastMarkdownRef: MutableRefObject<string>;
	onReload: (document: WebDocument) => void;
	onError: (error: unknown) => void;
}) {
	const [editor] = useLexicalComposerContext();
	const currentVersionRef = useRef(props.currentVersion);

	useEffect(() => {
		currentVersionRef.current = props.currentVersion;
	}, [props.currentVersion]);

	useEffect(() => {
		const interval = setInterval(() => {
			loadChangedWebDocument({
				currentVersion: currentVersionRef.current,
				fetchVersion,
				fetchDocument,
			}).then((changedDocument) => {
				if (!changedDocument) {
					return;
				}

				props.autosave.cancel();
				props.lastMarkdownRef.current = changedDocument.markdown;
				currentVersionRef.current = changedDocument.version;
				editor.update(() => {
					$convertFromMarkdownString(
						changedDocument.markdown,
						markdownTransformers,
					);
				});
				props.onReload(changedDocument);
			}, props.onError);
		}, pollingIntervalMs);

		return () => {
			clearInterval(interval);
		};
	}, [editor, props]);

	return null;
}

function handleEditorChange(
	editorState: EditorState,
	lastMarkdownRef: MutableRefObject<string>,
	autosave: ReturnType<typeof createDebouncedAutosave>,
	setStatus: (status: SaveStatus) => void,
) {
	editorState.read(() => {
		const markdown = $convertToMarkdownString(markdownTransformers);

		if (markdown === lastMarkdownRef.current) {
			return;
		}

		lastMarkdownRef.current = markdown;
		setStatus("saving");
		autosave.schedule(markdown);
	});
}

async function saveMarkdown(markdown: string): Promise<WebDocument> {
	const response = await fetch("/api/web-document", {
		method: "PUT",
		headers: {
			"content-type": "application/json",
		},
		body: JSON.stringify({ markdown }),
	});

	return readWebDocumentResponse(response);
}

async function fetchDocument(): Promise<WebDocument> {
	const response = await fetch("/api/web-document", { cache: "no-store" });

	return readWebDocumentResponse(response);
}

async function fetchVersion(): Promise<Pick<WebDocument, "version">> {
	const response = await fetch("/api/web-document/version", {
		cache: "no-store",
	});

	if (!response.ok) {
		throw new Error(
			`Could not read Webドキュメント version: ${response.status}`,
		);
	}

	const body: unknown = await response.json();

	if (!isVersionResponse(body)) {
		throw new Error("Webドキュメント version response was invalid.");
	}

	return body;
}

async function readWebDocumentResponse(
	response: Response,
): Promise<WebDocument> {
	if (!response.ok) {
		throw new Error(
			`Could not save or load Webドキュメント: ${response.status}`,
		);
	}

	const body: unknown = await response.json();

	if (!isWebDocument(body)) {
		throw new Error("Webドキュメント response was invalid.");
	}

	return body;
}

function isWebDocument(body: unknown): body is WebDocument {
	if (typeof body !== "object" || body === null) {
		return false;
	}

	const candidate = body as Record<string, unknown>;

	return (
		candidate.id === "shared" &&
		typeof candidate.markdown === "string" &&
		typeof candidate.version === "number"
	);
}

function isVersionResponse(
	body: unknown,
): body is Pick<WebDocument, "version"> {
	return (
		typeof body === "object" &&
		body !== null &&
		typeof (body as Record<string, unknown>).version === "number"
	);
}

function statusLabel(
	status: SaveStatus,
	errorMessage: string | undefined,
): string {
	if (status === "saving") {
		return "Autosaving...";
	}

	if (status === "reloaded") {
		return "Reloaded from a newer version";
	}

	if (status === "error") {
		return errorMessage ?? "An error occurred";
	}

	return "Saved";
}

const editorTheme = {
	code: "editor-code",
	heading: {
		h1: "editor-heading-h1",
		h2: "editor-heading-h2",
		h3: "editor-heading-h3",
	},
	list: {
		listitem: "editor-list-item",
		nested: {
			listitem: "editor-nested-list-item",
		},
		ol: "editor-list-ol",
		ul: "editor-list-ul",
	},
	paragraph: "editor-paragraph",
	quote: "editor-quote",
	text: {
		bold: "editor-text-bold",
		code: "editor-text-code",
		italic: "editor-text-italic",
	},
};
