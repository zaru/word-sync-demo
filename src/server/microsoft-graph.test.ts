import { afterEach, describe, expect, it, vi } from "vitest";

const msalCalls = vi.hoisted(() => ({
	configs: [] as unknown[],
	deserializedTokenCaches: [] as string[],
	silentTokenRequests: [] as unknown[],
}));

vi.mock("@azure/msal-node", () => ({
	ConfidentialClientApplication: class {
		constructor(config: unknown) {
			msalCalls.configs.push(config);
		}

		getTokenCache() {
			return {
				deserialize(tokenCache: string) {
					msalCalls.deserializedTokenCaches.push(tokenCache);
				},
				async getAllAccounts() {
					return [{ homeAccountId: "home-account-1" }];
				},
			};
		}

		async acquireTokenSilent(request: unknown) {
			msalCalls.silentTokenRequests.push(request);

			return { accessToken: "graph-access-token" };
		}
	},
}));

describe("Microsoft Graph App Folder boundary", () => {
	afterEach(() => {
		msalCalls.configs.length = 0;
		msalCalls.deserializedTokenCaches.length = 0;
		msalCalls.silentTokenRequests.length = 0;
		vi.unstubAllGlobals();
	});

	it("uploads the OneDrive作業コピー to the App Folder with signed-in editor Graph access", async () => {
		const fetchGraph = vi.fn(
			async (url: RequestInfo | URL, init?: RequestInit) => {
				expect(String(url)).toBe(
					"https://graph.microsoft.com/v1.0/me/drive/special/approot:/Web%E3%83%89%E3%82%AD%E3%83%A5%E3%83%A1%E3%83%B3%E3%83%88%201.docx:/content",
				);
				expect(init?.method).toBe("PUT");
				expect(init?.headers).toEqual({
					authorization: "Bearer graph-access-token",
					"content-type":
						"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
				});

				if (!(init?.body instanceof Blob)) {
					throw new Error("Graph upload body was not a Blob.");
				}

				expect(new Uint8Array(await init.body.arrayBuffer())).toEqual(
					new Uint8Array([1, 2, 3]),
				);

				return Response.json({
					id: "drive-item-1",
					webUrl: "https://onedrive.example/Webドキュメント 1.docx",
				});
			},
		);
		vi.stubGlobal("fetch", fetchGraph);
		const { createMicrosoftGraphAppFolderBoundary } = await import(
			"./microsoft-graph"
		);

		const graph = createMicrosoftGraphAppFolderBoundary({
			clientId: "client-id",
			clientSecret: "client-secret",
		});

		await expect(
			graph.uploadAppFolderWorkingCopy({
				content: new Uint8Array([1, 2, 3]),
				fileName: "Webドキュメント 1.docx",
				tokenCache: '{"RefreshToken":{"cached":true}}',
			}),
		).resolves.toEqual({
			driveItemId: "drive-item-1",
			webUrl: "https://onedrive.example/Webドキュメント 1.docx",
		});
		expect(msalCalls.configs).toEqual([
			{
				auth: {
					authority: "https://login.microsoftonline.com/common",
					clientId: "client-id",
					clientSecret: "client-secret",
				},
			},
		]);
		expect(msalCalls.deserializedTokenCaches).toEqual([
			'{"RefreshToken":{"cached":true}}',
		]);
		expect(msalCalls.silentTokenRequests).toEqual([
			{
				account: { homeAccountId: "home-account-1" },
				scopes: ["Files.ReadWrite.AppFolder"],
			},
		]);
		expect(fetchGraph).toHaveBeenCalledTimes(1);
	});

	it("deletes the OneDrive作業コピー from the App Folder by DriveItem ID", async () => {
		const fetchGraph = vi.fn(
			async (url: RequestInfo | URL, init?: RequestInit) => {
				expect(String(url)).toBe(
					"https://graph.microsoft.com/v1.0/me/drive/items/drive-item-1",
				);
				expect(init?.method).toBe("DELETE");
				expect(init?.headers).toEqual({
					authorization: "Bearer graph-access-token",
				});

				return new Response(null, { status: 204 });
			},
		);
		vi.stubGlobal("fetch", fetchGraph);
		const { createMicrosoftGraphAppFolderBoundary } = await import(
			"./microsoft-graph"
		);

		const graph = createMicrosoftGraphAppFolderBoundary({
			clientId: "client-id",
			clientSecret: "client-secret",
		});

		await expect(
			graph.deleteAppFolderWorkingCopy({
				driveItemId: "drive-item-1",
				tokenCache: '{"RefreshToken":{"cached":true}}',
			}),
		).resolves.toBeUndefined();
		expect(msalCalls.deserializedTokenCaches).toEqual([
			'{"RefreshToken":{"cached":true}}',
		]);
		expect(msalCalls.silentTokenRequests).toEqual([
			{
				account: { homeAccountId: "home-account-1" },
				scopes: ["Files.ReadWrite.AppFolder"],
			},
		]);
		expect(fetchGraph).toHaveBeenCalledTimes(1);
	});
});
