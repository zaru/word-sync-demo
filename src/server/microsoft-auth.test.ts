import { afterEach, describe, expect, it, vi } from "vitest";

const msalCalls = vi.hoisted(() => ({
	authCodeUrlRequests: [] as unknown[],
	configs: [] as unknown[],
	tokenRequests: [] as unknown[],
}));

vi.mock("@azure/msal-node", () => ({
	ConfidentialClientApplication: class {
		constructor(config: unknown) {
			msalCalls.configs.push(config);
		}

		async getAuthCodeUrl(request: unknown) {
			msalCalls.authCodeUrlRequests.push(request);

			return "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
		}

		async acquireTokenByCode(request: unknown) {
			msalCalls.tokenRequests.push(request);

			return {
				account: {
					homeAccountId: "home-account-1",
					name: "編集者 A",
					username: "editor@example.com",
				},
			};
		}

		getTokenCache() {
			return {
				serialize() {
					return '{"AccessToken":{},"RefreshToken":{"cached":true}}';
				},
			};
		}
	},
}));

describe("MicrosoftAuthBoundary", () => {
	afterEach(() => {
		msalCalls.authCodeUrlRequests.length = 0;
		msalCalls.configs.length = 0;
		msalCalls.tokenRequests.length = 0;
	});

	it("uses MSAL Node common authority and requests App Folder plus offline access scopes", async () => {
		const { createMicrosoftAuthBoundary } = await import("./microsoft-auth");
		const auth = createMicrosoftAuthBoundary({
			clientId: "client-id",
			clientSecret: "client-secret",
		});

		await expect(
			auth.createAuthorizationUrl({
				redirectUri: "http://localhost/auth/callback",
				scopes: [
					"openid",
					"profile",
					"email",
					"offline_access",
					"Files.ReadWrite.AppFolder",
				],
				state: "state-1",
			}),
		).resolves.toBe(
			"https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
		);
		await expect(
			auth.completeAuthorizationCode({
				code: "auth-code",
				redirectUri: "http://localhost/auth/callback",
				scopes: [
					"openid",
					"profile",
					"email",
					"offline_access",
					"Files.ReadWrite.AppFolder",
				],
			}),
		).resolves.toEqual({
			editorId: "home-account-1",
			displayName: "編集者 A",
			username: "editor@example.com",
			tokenCache: '{"AccessToken":{},"RefreshToken":{"cached":true}}',
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
		expect(msalCalls.authCodeUrlRequests).toEqual([
			{
				redirectUri: "http://localhost/auth/callback",
				scopes: [
					"openid",
					"profile",
					"email",
					"offline_access",
					"Files.ReadWrite.AppFolder",
				],
				state: "state-1",
			},
		]);
		expect(msalCalls.tokenRequests).toEqual([
			{
				code: "auth-code",
				redirectUri: "http://localhost/auth/callback",
				scopes: [
					"openid",
					"profile",
					"email",
					"offline_access",
					"Files.ReadWrite.AppFolder",
				],
			},
		]);
	});
});
