import { requestUrl, RequestUrlResponse } from "obsidian";
import { interactiveLogin, refreshAccessToken, TokenSet } from "./auth";
import {
	GraphAttachment,
	GraphEvent,
	GraphMailFolder,
	GraphMessage,
	PluginSettings,
	TeamsMessage,
	scopesFor,
} from "./types";
import { info, log, logError, withTimeout, SyncCancelledError } from "./log";

const REQUEST_TIMEOUT_MS = 60_000;

/** Error carrying the HTTP status so callers can special-case (e.g. 410). */
export class GraphError extends Error {
	constructor(message: string, readonly status: number) {
		super(message);
		this.name = "GraphError";
	}
}

interface GraphInit {
	method?: string;
	headers?: Record<string, string>;
	body?: string;
}

/** Case-insensitive header lookup for requestUrl's response headers. */
function headerValue(headers: Record<string, string>, name: string): string | undefined {
	const lower = name.toLowerCase();
	for (const key of Object.keys(headers)) {
		if (key.toLowerCase() === lower) return headers[key];
	}
	return undefined;
}

const GRAPH = "https://graph.microsoft.com/v1.0";

// Fields we request for each message. Keep tight to reduce payload.
const MESSAGE_SELECT = [
	"id",
	"conversationId",
	"subject",
	"bodyPreview",
	"body",
	"from",
	"sender",
	"toRecipients",
	"ccRecipients",
	"receivedDateTime",
	"sentDateTime",
	"hasAttachments",
	"isRead",
	"importance",
	"webLink",
	"categories",
].join(",");

export interface DeltaResult {
	messages: GraphMessage[];
	/** Persist this and pass it next time for incremental sync. */
	deltaLink: string | null;
}

/**
 * Thin Microsoft Graph client. Owns the in-memory access token and refreshes
 * it transparently. Persistence of the refresh token is delegated back to the
 * plugin via `onTokens`.
 */
export class GraphClient {
	private token: TokenSet | null = null;

	constructor(
		private settings: PluginSettings,
		private openBrowser: (url: string) => void,
		private onTokens: (refreshToken: string | null) => Promise<void>,
	) {}

	get isAuthenticated(): boolean {
		return !!this.settings.refreshToken;
	}

	/** Interactive sign-in; stores the refresh token. */
	async login(): Promise<void> {
		const tokens = await interactiveLogin(
			this.settings.clientId,
			this.settings.tenant,
			scopesFor(this.settings),
			this.openBrowser,
		);
		this.token = tokens;
		this.settings.refreshToken = tokens.refreshToken;
		await this.onTokens(tokens.refreshToken);
		log("Interactive login complete.");
	}

	async logout(): Promise<void> {
		this.token = null;
		this.settings.refreshToken = null;
		await this.onTokens(null);
	}

	private async accessToken(): Promise<string> {
		// Reuse a still-valid token (60s safety margin).
		if (this.token && this.token.expiresAt - 60_000 > Date.now()) {
			return this.token.accessToken;
		}
		if (!this.settings.refreshToken) {
			throw new Error("Not signed in. Open the plugin settings and connect your account.");
		}
		log("Refreshing access token.");
		const tokens = await refreshAccessToken(
			this.settings.clientId,
			this.settings.tenant,
			scopesFor(this.settings),
			this.settings.refreshToken,
		);
		this.token = tokens;
		// Microsoft rotates refresh tokens — persist the new one.
		if (tokens.refreshToken && tokens.refreshToken !== this.settings.refreshToken) {
			this.settings.refreshToken = tokens.refreshToken;
			await this.onTokens(tokens.refreshToken);
		}
		return tokens.accessToken;
	}

	// Uses Obsidian's requestUrl rather than fetch: fetch runs in the renderer
	// and adds `Origin: app://obsidian.md`, which trips Entra's cross-origin/SPA
	// checks (AADSTS9002326). requestUrl is a native request with no Origin and
	// bypasses CORS.
	private async graphRequest(url: string, init: GraphInit = {}, retry = true): Promise<RequestUrlResponse> {
		const token = await this.accessToken();
		const headers: Record<string, string> = { ...(init.headers ?? {}), Authorization: `Bearer ${token}` };
		if (init.method && init.method !== "GET" && !("Content-Type" in headers)) {
			headers["Content-Type"] = "application/json";
		}
		const resp = await withTimeout(
			requestUrl({
				url,
				method: init.method ?? "GET",
				headers,
				body: init.body,
				throw: false,
			}),
			REQUEST_TIMEOUT_MS,
			`Graph ${init.method ?? "GET"} request`,
		);

		// Transparent single retry on 401 (stale token) and on 429/503 (throttle).
		if (resp.status === 401 && retry) {
			this.token = null;
			return this.graphRequest(url, init, false);
		}
		if ((resp.status === 429 || resp.status === 503) && retry) {
			const wait = Number(headerValue(resp.headers, "retry-after") ?? "2") * 1000;
			log(`Throttled (${resp.status}); waiting ${wait}ms.`);
			await sleep(wait);
			return this.graphRequest(url, init, false);
		}
		return resp;
	}

	private async graphJson<T>(url: string, init: GraphInit = {}): Promise<T> {
		const resp = await this.graphRequest(url, init);
		if (resp.status < 200 || resp.status >= 300) {
			const text = resp.text ?? "";
			logError(`Graph ${init.method ?? "GET"} ${url} -> ${resp.status}`, text);
			throw new GraphError(`Graph request failed (${resp.status}): ${text.slice(0, 300)}`, resp.status);
		}
		return resp.json as T;
	}

	/** Returns the signed-in user's UPN/email, for display. */
	async me(): Promise<{ userPrincipalName?: string; mail?: string; displayName?: string }> {
		return this.graphJson(`${GRAPH}/me?$select=displayName,mail,userPrincipalName`);
	}

	/**
	 * Resolves an Outlook folder path to its folder id. Single-segment paths try
	 * the well-known name first (inbox, archive, sentitems, …); otherwise the
	 * path is walked segment by segment matching `displayName`.
	 */
	async resolveFolderId(path: string): Promise<string> {
		const segments = path.split("/").map((s) => s.trim()).filter(Boolean);
		if (segments.length === 0) throw new Error("Empty folder path.");

		let parentId: string | null = null;

		// Well-known shortcut for a single, recognised segment.
		if (segments.length === 1 && WELL_KNOWN.has(segments[0].toLowerCase())) {
			return segments[0].toLowerCase();
		}

		for (const segment of segments) {
			const base = parentId
				? `${GRAPH}/me/mailFolders/${parentId}/childFolders`
				: `${GRAPH}/me/mailFolders`;
			// Graph does not allow $filter on childFolders reliably across tenants,
			// so page through and match display name case-insensitively.
			const found = await this.findChildByName(base, segment);
			if (!found) {
				throw new Error(`Outlook folder not found: "${segment}" (in path "${path}").`);
			}
			parentId = found.id;
		}
		return parentId as string;
	}

	private async findChildByName(base: string, name: string): Promise<GraphMailFolder | null> {
		let url: string | null = `${base}?$select=id,displayName,childFolderCount&$top=100`;
		const target = name.toLowerCase();
		while (url) {
			const page: { value: GraphMailFolder[]; "@odata.nextLink"?: string } =
				await this.graphJson(url);
			const hit = page.value.find((f) => f.displayName?.toLowerCase() === target);
			if (hit) return hit;
			url = page["@odata.nextLink"] ?? null;
		}
		return null;
	}

	/**
	 * Delta sync for a folder. If `deltaLink` is provided, only changes since the
	 * last call are returned; otherwise a full initial enumeration runs. Follows
	 * `@odata.nextLink` paging and returns the final `@odata.deltaLink`.
	 */
	async deltaMessages(
		folderId: string,
		deltaLink: string | null,
		onProgress?: (fetched: number, page: number) => void,
		shouldCancel?: () => boolean,
	): Promise<DeltaResult> {
		let url: string | null =
			deltaLink ??
			`${GRAPH}/me/mailFolders/${encodeURIComponent(folderId)}/messages/delta?$select=${MESSAGE_SELECT}`;

		const messages: GraphMessage[] = [];
		let finalDeltaLink: string | null = null;
		let pageNum = 0;

		info(`Folder ${folderId}: ${deltaLink ? "fetching changes (delta)" : "full enumeration"}…`);
		while (url) {
			if (shouldCancel?.()) {
				info(`Folder ${folderId}: fetch stopped by user after ${messages.length} message(s).`);
				throw new SyncCancelledError();
			}
			const page: {
				value: GraphMessage[];
				"@odata.nextLink"?: string;
				"@odata.deltaLink"?: string;
			} = await this.graphJson(url, {
				// Ask for HTML bodies; turndown converts them.
				headers: { Prefer: 'outlook.body-content-type="html"' },
			});
			messages.push(...(page.value ?? []));
			pageNum++;
			info(`Folder ${folderId}: page ${pageNum}, ${messages.length} message(s) fetched so far.`);
			onProgress?.(messages.length, pageNum);
			if (page["@odata.nextLink"]) {
				url = page["@odata.nextLink"];
			} else {
				finalDeltaLink = page["@odata.deltaLink"] ?? null;
				url = null;
			}
		}
		info(`Folder ${folderId}: ${messages.length} changed message(s) total.`);
		return { messages, deltaLink: finalDeltaLink };
	}

	/** Lists file attachments for a message (metadata + base64 bytes). */
	async listAttachments(messageId: string): Promise<GraphAttachment[]> {
		const resp = await this.graphJson<{ value: GraphAttachment[] }>(
			`${GRAPH}/me/messages/${encodeURIComponent(messageId)}/attachments`,
		);
		return resp.value ?? [];
	}

	/**
	 * Fetches calendar events between two instants (inclusive) via calendarView,
	 * which expands recurring series into individual occurrences. Times are
	 * requested in UTC (see Prefer header) and returned without a 'Z' suffix.
	 */
	async calendarView(
		startIso: string,
		endIso: string,
		onProgress?: (fetched: number) => void,
		shouldCancel?: () => boolean,
	): Promise<GraphEvent[]> {
		const params = new URLSearchParams({
			startDateTime: startIso,
			endDateTime: endIso,
			$select: EVENT_SELECT,
			$orderby: "start/dateTime",
			$top: "50",
		});
		let url: string | null = `${GRAPH}/me/calendarView?${params.toString()}`;
		const events: GraphEvent[] = [];

		info(`Calendar: fetching ${startIso.slice(0, 10)} → ${endIso.slice(0, 10)}…`);
		while (url) {
			if (shouldCancel?.()) {
				info(`Calendar: fetch stopped by user after ${events.length} event(s).`);
				throw new SyncCancelledError();
			}
			const page: { value: GraphEvent[]; "@odata.nextLink"?: string } = await this.graphJson(url, {
				headers: { Prefer: 'outlook.timezone="UTC"' },
			});
			events.push(...(page.value ?? []));
			onProgress?.(events.length);
			url = page["@odata.nextLink"] ?? null;
		}
		info(`Calendar: ${events.length} event(s) in window.`);
		return events;
	}

	// ---- Teams ----

	/** Lists the user's chats with expanded members (for titling). */
	async listChats(): Promise<{ id: string; topic: string; chatType: string; members: string[] }[]> {
		const out: { id: string; topic: string; chatType: string; members: string[] }[] = [];
		let url: string | null = `${GRAPH}/me/chats?$expand=members&$top=50`;
		info("Teams: listing chats…");
		while (url) {
			const page: { value: RawChat[]; "@odata.nextLink"?: string } = await this.graphJson(url);
			for (const c of page.value ?? []) {
				out.push({
					id: c.id,
					topic: c.topic ?? "",
					chatType: c.chatType ?? "",
					members: (c.members ?? []).map((m) => m.displayName ?? "").filter(Boolean),
				});
			}
			info(`Teams: ${out.length} chat(s) listed so far…`);
			url = page["@odata.nextLink"] ?? null;
		}
		info(`Teams: ${out.length} chat(s) total.`);
		return out;
	}

	/** Teams the user has joined. */
	async listJoinedTeams(): Promise<{ id: string; displayName: string }[]> {
		const resp = await this.graphJson<{ value: { id: string; displayName: string }[] }>(
			`${GRAPH}/me/joinedTeams`,
		);
		return resp.value ?? [];
	}

	/** Channels within a team. */
	async listChannels(teamId: string): Promise<{ id: string; displayName: string }[]> {
		const resp = await this.graphJson<{ value: { id: string; displayName: string }[] }>(
			`${GRAPH}/teams/${encodeURIComponent(teamId)}/channels`,
		);
		return resp.value ?? [];
	}

	/**
	 * Lists a chat's messages, newest-first, paging until older than `sinceIso`.
	 *
	 * Graph does **not** support delta / change-tracking on chat messages
	 * (`/chats/{id}/messages/delta` → 400 "Change tracking is not supported
	 * against 'microsoft.graph.chatMessage'"), so chats use plain listing bounded
	 * by the sync window. Incremental behaviour comes from the caller's
	 * per-conversation `lastWritten` timestamp (dedup), not a server cursor —
	 * hence `deltaLink` is always null here.
	 */
	async chatMessagesList(
		chatId: string,
		sinceIso: string,
		onProgress?: (fetched: number) => void,
		shouldCancel?: () => boolean,
	): Promise<{ messages: TeamsMessage[]; deltaLink: string | null }> {
		let url: string | null = `${GRAPH}/chats/${encodeURIComponent(chatId)}/messages?$top=50`;
		const messages: TeamsMessage[] = [];
		let pageNum = 0;
		while (url) {
			if (shouldCancel?.()) throw new SyncCancelledError();
			const page: { value: RawChatMessage[]; "@odata.nextLink"?: string } =
				await this.graphJson(url);
			const batch = (page.value ?? []).map(parseTeamsMessage);
			messages.push(...batch);
			pageNum++;
			info(`Teams: chat page ${pageNum}, ${messages.length} message(s) fetched so far.`);
			onProgress?.(messages.length);
			// Messages arrive newest-first. Once any message on this page predates
			// the window floor, every later page is older too — stop paging.
			if (sinceIso && batch.some((m) => m.createdIso && m.createdIso < sinceIso)) break;
			url = page["@odata.nextLink"] ?? null;
		}
		return { messages, deltaLink: null };
	}

	/** Delta of top-level messages in a channel. */
	async channelMessagesDelta(
		teamId: string,
		channelId: string,
		deltaLink: string | null,
		onProgress?: (fetched: number) => void,
		shouldCancel?: () => boolean,
	): Promise<{ messages: TeamsMessage[]; deltaLink: string | null }> {
		const start = `${GRAPH}/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/delta`;
		return this.messagesDelta(deltaLink ?? start, onProgress, shouldCancel);
	}

	/** Replies to a channel message (no delta; paged in full). */
	async listReplies(teamId: string, channelId: string, messageId: string): Promise<TeamsMessage[]> {
		const out: TeamsMessage[] = [];
		let url: string | null =
			`${GRAPH}/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/${encodeURIComponent(messageId)}/replies`;
		while (url) {
			const page: { value: RawChatMessage[]; "@odata.nextLink"?: string } = await this.graphJson(url);
			for (const m of page.value ?? []) out.push(parseTeamsMessage(m));
			url = page["@odata.nextLink"] ?? null;
		}
		return out;
	}

	/** Shared paginator for the chat/channel message delta endpoints. */
	private async messagesDelta(
		startUrl: string,
		onProgress?: (fetched: number) => void,
		shouldCancel?: () => boolean,
	): Promise<{ messages: TeamsMessage[]; deltaLink: string | null }> {
		let url: string | null = startUrl;
		const messages: TeamsMessage[] = [];
		let finalDeltaLink: string | null = null;
		let pageNum = 0;
		while (url) {
			if (shouldCancel?.()) throw new SyncCancelledError();
			const page: {
				value: RawChatMessage[];
				"@odata.nextLink"?: string;
				"@odata.deltaLink"?: string;
			} = await this.graphJson(url);
			for (const m of page.value ?? []) messages.push(parseTeamsMessage(m));
			pageNum++;
			info(`Teams: page ${pageNum}, ${messages.length} message(s) fetched so far.`);
			onProgress?.(messages.length);
			if (page["@odata.nextLink"]) {
				url = page["@odata.nextLink"];
			} else {
				finalDeltaLink = page["@odata.deltaLink"] ?? null;
				url = null;
			}
		}
		return { messages, deltaLink: finalDeltaLink };
	}
}

interface RawChat {
	id: string;
	topic?: string;
	chatType?: string;
	members?: { displayName?: string }[];
}

interface RawChatMessage {
	id: string;
	messageType?: string;
	createdDateTime?: string;
	deletedDateTime?: string | null;
	from?: { user?: { displayName?: string }; application?: { displayName?: string } };
	body?: { contentType?: "html" | "text"; content?: string };
	attachments?: { name?: string; contentType?: string }[];
}

function parseTeamsMessage(m: RawChatMessage): TeamsMessage {
	const html = m.body?.contentType === "html";
	const content = m.body?.content ?? "";
	return {
		id: m.id,
		createdIso: m.createdDateTime ?? "",
		from: m.from?.user?.displayName || m.from?.application?.displayName || "Unknown",
		bodyHtml: html ? content : undefined,
		bodyText: html ? undefined : content,
		attachments: (m.attachments ?? []).map((a) => a.name ?? "").filter(Boolean),
		deleted: !!m.deletedDateTime,
		system: (m.messageType ?? "message") !== "message",
	};
}

const EVENT_SELECT = [
	"id",
	"subject",
	"bodyPreview",
	"body",
	"start",
	"end",
	"isAllDay",
	"isCancelled",
	"location",
	"organizer",
	"attendees",
	"onlineMeeting",
	"onlineMeetingUrl",
	"webLink",
	"seriesMasterId",
	"type",
	"categories",
].join(",");

const WELL_KNOWN = new Set([
	"inbox",
	"archive",
	"drafts",
	"sentitems",
	"deleteditems",
	"junkemail",
	"outbox",
	"clutter",
	"conflicts",
	"conversationhistory",
	"localfailures",
	"recoverableitemsdeletions",
	"scheduled",
	"searchfolders",
	"serverfailures",
	"syncissues",
]);

function sleep(ms: number): Promise<void> {
	return new Promise((r) => window.setTimeout(r, ms));
}
