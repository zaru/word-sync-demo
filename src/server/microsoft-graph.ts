import { ConfidentialClientApplication } from "@azure/msal-node";

import type { GraphAppFolderBoundary } from "./word-edit-session-routes";

const microsoftAuthority = "https://login.microsoftonline.com/common";
const graphAppFolderUploadScope = "Files.ReadWrite.AppFolder";
const docxContentType =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export function createMicrosoftGraphAppFolderBoundary(options: {
	clientId: string;
	clientSecret: string;
}): GraphAppFolderBoundary {
	return {
		async uploadAppFolderWorkingCopy(input) {
			const client = new ConfidentialClientApplication({
				auth: {
					authority: microsoftAuthority,
					clientId: options.clientId,
					clientSecret: options.clientSecret,
				},
			});
			const tokenCache = client.getTokenCache();
			tokenCache.deserialize(input.tokenCache);

			const [account] = await tokenCache.getAllAccounts();

			if (!account) {
				throw new Error("Signed-in 編集者 Graph token cache has no account.");
			}

			const authResult = await client.acquireTokenSilent({
				account,
				scopes: [graphAppFolderUploadScope],
			});

			if (!authResult?.accessToken) {
				throw new Error(
					"Could not acquire Graph access for OneDrive作業コピー.",
				);
			}

			const response = await fetch(createAppFolderUploadUrl(input.fileName), {
				body: new Blob([toArrayBuffer(input.content)], {
					type: docxContentType,
				}),
				headers: {
					authorization: `Bearer ${authResult.accessToken}`,
					"content-type": docxContentType,
				},
				method: "PUT",
			});

			if (!response.ok) {
				throw new Error(
					`Could not upload OneDrive作業コピー to App Folder: ${response.status}`,
				);
			}

			const body: unknown = await response.json();

			if (!isDriveItemResponse(body)) {
				throw new Error("Graph App Folder upload response was invalid.");
			}

			return {
				driveItemId: body.id,
				webUrl: body.webUrl,
			};
		},
	};
}

function createAppFolderUploadUrl(fileName: string): string {
	return `https://graph.microsoft.com/v1.0/me/drive/special/approot:/${encodeURIComponent(fileName)}:/content`;
}

function toArrayBuffer(content: Uint8Array): ArrayBuffer {
	const buffer = new ArrayBuffer(content.byteLength);
	new Uint8Array(buffer).set(content);

	return buffer;
}

function isDriveItemResponse(
	body: unknown,
): body is { id: string; webUrl: string } {
	if (typeof body !== "object" || body === null) {
		return false;
	}

	const candidate = body as Record<string, unknown>;

	return (
		typeof candidate.id === "string" && typeof candidate.webUrl === "string"
	);
}
