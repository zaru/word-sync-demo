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
			const accessToken = await acquireGraphAccessToken({
				clientId: options.clientId,
				clientSecret: options.clientSecret,
				tokenCache: input.tokenCache,
			});

			const response = await fetch(createAppFolderUploadUrl(input.fileName), {
				body: new Blob([toArrayBuffer(input.content)], {
					type: docxContentType,
				}),
				headers: {
					authorization: `Bearer ${accessToken}`,
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

		async downloadAppFolderWorkingCopy(input) {
			const accessToken = await acquireGraphAccessToken({
				clientId: options.clientId,
				clientSecret: options.clientSecret,
				tokenCache: input.tokenCache,
			});

			const response = await fetch(
				createDriveItemContentUrl(input.driveItemId),
				{
					headers: {
						authorization: `Bearer ${accessToken}`,
					},
					method: "GET",
				},
			);

			if (!response.ok) {
				throw new Error(
					`Could not download OneDrive作業コピー from App Folder: ${response.status}`,
				);
			}

			return new Uint8Array(await response.arrayBuffer());
		},

		async deleteAppFolderWorkingCopy(input) {
			const accessToken = await acquireGraphAccessToken({
				clientId: options.clientId,
				clientSecret: options.clientSecret,
				tokenCache: input.tokenCache,
			});

			const response = await fetch(createDriveItemUrl(input.driveItemId), {
				headers: {
					authorization: `Bearer ${accessToken}`,
				},
				method: "DELETE",
			});

			if (!response.ok) {
				throw new Error(
					`Could not delete OneDrive作業コピー from App Folder: ${response.status}`,
				);
			}
		},
	};
}

function createAppFolderUploadUrl(fileName: string): string {
	return `https://graph.microsoft.com/v1.0/me/drive/special/approot:/${encodeURIComponent(fileName)}:/content`;
}

function createDriveItemContentUrl(driveItemId: string): string {
	return `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(driveItemId)}/content`;
}

function createDriveItemUrl(driveItemId: string): string {
	return `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(driveItemId)}`;
}

async function acquireGraphAccessToken(options: {
	clientId: string;
	clientSecret: string;
	tokenCache: string;
}): Promise<string> {
	const client = new ConfidentialClientApplication({
		auth: {
			authority: microsoftAuthority,
			clientId: options.clientId,
			clientSecret: options.clientSecret,
		},
	});
	const tokenCache = client.getTokenCache();
	tokenCache.deserialize(options.tokenCache);

	const [account] = await tokenCache.getAllAccounts();

	if (!account) {
		throw new Error("Signed-in 編集者 Graph token cache has no account.");
	}

	const authResult = await client.acquireTokenSilent({
		account,
		scopes: [graphAppFolderUploadScope],
	});

	if (!authResult?.accessToken) {
		throw new Error("Could not acquire Graph access for OneDrive作業コピー.");
	}

	return authResult.accessToken;
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
