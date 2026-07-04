/* eslint-disable @typescript-eslint/no-unsafe-return, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-argument --
 * Desktop-only module: it runs a loopback OAuth redirect server using Node's
 * built-in `http`, `crypto`, and `net`. Under the review linter's TypeScript
 * program these built-ins resolve to `any` (its project does not load
 * @types/node), so every use trips the type-aware no-unsafe-* rules even though
 * the code is correctly typed (e.g. `buf: Buffer`, typed req/res, `AddressInfo`
 * casts). The plugin's own `tsc` build loads @types/node and type-checks this
 * file cleanly, so these are false positives specific to the linter's setup. */
import * as http from "http";
import * as crypto from "crypto";
import { AddressInfo } from "net";
import { requestUrl } from "obsidian";
import { log, logError, withTimeout } from "./log";

export interface TokenSet {
	accessToken: string;
	refreshToken: string | null;
	expiresAt: number; // epoch ms
}

interface TokenResponse {
	access_token: string;
	refresh_token?: string;
	expires_in: number;
	token_type: string;
	scope?: string;
	error?: string;
	error_description?: string;
}

function authorizeEndpoint(tenant: string): string {
	return `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/authorize`;
}

function tokenEndpoint(tenant: string): string {
	return `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`;
}

function base64url(buf: Buffer): string {
	return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function makePkce(): { verifier: string; challenge: string } {
	const verifier = base64url(crypto.randomBytes(32));
	const challenge = base64url(crypto.createHash("sha256").update(verifier).digest());
	return { verifier, challenge };
}

/**
 * Runs the interactive Authorization Code + PKCE flow using a loopback
 * redirect. Microsoft allows `http://localhost` (any port) as a redirect URI
 * for the "Mobile and desktop applications" platform, so no fixed port needs
 * to be registered. Returns a full token set including a refresh token.
 */
export async function interactiveLogin(
	clientId: string,
	tenant: string,
	scopes: string[],
	openBrowser: (url: string) => void,
): Promise<TokenSet> {
	if (!clientId) throw new Error("Client ID is not configured.");
	const { verifier, challenge } = makePkce();
	const state = base64url(crypto.randomBytes(16));

	const { code, redirectUri } = await new Promise<{ code: string; redirectUri: string }>(
		(resolve, reject) => {
			const server = http.createServer((req, res) => {
				try {
					const url = new URL(req.url ?? "/", "http://localhost");
					if (url.pathname !== "/") {
						res.writeHead(404);
						res.end();
						return;
					}
					const returnedState = url.searchParams.get("state");
					const err = url.searchParams.get("error");
					const code = url.searchParams.get("code");

					res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
					if (err) {
						res.end(htmlPage("Login failed", url.searchParams.get("error_description") || err));
						cleanup();
						reject(new Error(`${err}: ${url.searchParams.get("error_description") ?? ""}`));
						return;
					}
					if (returnedState !== state) {
						res.end(htmlPage("Login failed", "State mismatch — possible CSRF. Try again."));
						cleanup();
						reject(new Error("OAuth state mismatch."));
						return;
					}
					if (!code) {
						res.end(htmlPage("Login failed", "No authorization code returned."));
						cleanup();
						reject(new Error("No authorization code returned."));
						return;
					}
					res.end(htmlPage("Signed in", "You can close this tab and return to Obsidian."));
					const addr = server.address() as AddressInfo;
					const redirectUri = `http://localhost:${addr.port}`;
					cleanup();
					resolve({ code, redirectUri });
				} catch (e) {
					cleanup();
					reject(e instanceof Error ? e : new Error(String(e)));
				}
			});

			const timeout = window.setTimeout(() => {
				cleanup();
				reject(new Error("Login timed out after 5 minutes."));
			}, 5 * 60 * 1000);

			function cleanup() {
				window.clearTimeout(timeout);
				server.close();
			}

			server.on("error", (e) => {
				window.clearTimeout(timeout);
				reject(e instanceof Error ? e : new Error(String(e)));
			});

			// Bind to a random free port on loopback only.
			server.listen(0, "127.0.0.1", () => {
				const addr = server.address() as AddressInfo;
				const redirectUri = `http://localhost:${addr.port}`;
				const params = new URLSearchParams({
					client_id: clientId,
					response_type: "code",
					redirect_uri: redirectUri,
					response_mode: "query",
					scope: scopes.join(" "),
					state,
					code_challenge: challenge,
					code_challenge_method: "S256",
					prompt: "select_account",
				});
				const authUrl = `${authorizeEndpoint(tenant)}?${params.toString()}`;
				log("Opening browser for consent:", authUrl);
				openBrowser(authUrl);
			});
		},
	);

	// Exchange the code for tokens.
	const body = new URLSearchParams({
		client_id: clientId,
		grant_type: "authorization_code",
		code,
		redirect_uri: redirectUri,
		code_verifier: verifier,
		scope: scopes.join(" "),
	});
	return exchange(tenant, body);
}

/** Uses a stored refresh token to obtain a fresh access token. */
export async function refreshAccessToken(
	clientId: string,
	tenant: string,
	scopes: string[],
	refreshToken: string,
): Promise<TokenSet> {
	const body = new URLSearchParams({
		client_id: clientId,
		grant_type: "refresh_token",
		refresh_token: refreshToken,
		scope: scopes.join(" "),
	});
	return exchange(tenant, body);
}

async function exchange(tenant: string, body: URLSearchParams): Promise<TokenSet> {
	// Use Obsidian's requestUrl (a native request with no Origin header) rather
	// than the renderer's fetch. fetch runs in Obsidian's Electron renderer and
	// attaches `Origin: app://obsidian.md`, which makes Entra treat the call as a
	// browser/SPA request and reject it with AADSTS9002326 ("cross-origin token
	// redemption is permitted only for the 'Single-Page Application' client-type").
	const resp = await withTimeout(
		requestUrl({
			url: tokenEndpoint(tenant),
			method: "POST",
			contentType: "application/x-www-form-urlencoded",
			headers: { "Content-Type": "application/x-www-form-urlencoded" },
			body: body.toString(),
			throw: false,
		}),
		30_000,
		"Token endpoint",
	);
	let json: TokenResponse;
	try {
		json = resp.json as TokenResponse;
	} catch {
		throw new Error(`Token endpoint returned ${resp.status}: ${(resp.text ?? "").slice(0, 300)}`);
	}
	const ok = resp.status >= 200 && resp.status < 300;
	if (!ok || json.error) {
		logError("Token exchange failed:", json);
		throw new Error(json.error_description || json.error || `Token endpoint returned ${resp.status}`);
	}
	return {
		accessToken: json.access_token,
		refreshToken: json.refresh_token ?? null,
		expiresAt: Date.now() + json.expires_in * 1000,
	};
}

function htmlPage(title: string, message: string): string {
	return `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title>
<style>body{font-family:system-ui,sans-serif;background:#1e1e1e;color:#eee;display:flex;
align-items:center;justify-content:center;height:100vh;margin:0}
.card{text-align:center;padding:2rem 3rem;background:#2a2a2a;border-radius:12px;
box-shadow:0 8px 30px rgba(0,0,0,.4)}h1{margin:0 0 .5rem;font-size:1.4rem}
p{margin:0;color:#aaa}</style></head>
<body><div class="card"><h1>${title}</h1><p>${message}</p></div></body></html>`;
}

/* eslint-enable @typescript-eslint/no-unsafe-return, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-argument */
