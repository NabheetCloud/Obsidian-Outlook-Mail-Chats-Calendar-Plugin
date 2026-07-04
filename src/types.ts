// Shared types for the Outlook Teams and Calendar plugin.

/** A single mail folder the user wants mirrored into the vault. */
export interface FolderMapping {
	/**
	 * The Outlook folder path. Either a well-known name (case-insensitive) such
	 * as "inbox", "archive", "sentitems", or a display-name path with "/" as the
	 * nesting separator, e.g. "Projects/ClientX".
	 */
	outlookPath: string;
	/**
	 * Vault-relative subfolder under `targetFolder` where notes for this folder
	 * land. If empty, a slug of `outlookPath` is used.
	 */
	vaultSubfolder: string;
	/** Whether this mapping participates in sync. */
	enabled: boolean;
}

export interface PluginSettings {
	// --- Azure AD app registration ---
	clientId: string;
	/** "common" | "organizations" | "consumers" | a specific tenant GUID. */
	tenant: string;

	// --- Vault layout ---
	/** Root vault folder for all mail, e.g. "10-Mailbox". */
	targetFolder: string;
	folders: FolderMapping[];

	// --- Sync behaviour ---
	/**
	 * Only write mail received on/after this date (YYYY-MM-DD). Empty = all mail.
	 * Applied as a write-filter on every sync; older messages are skipped.
	 */
	syncSince: string;
	/** Auto-sync interval in minutes. 0 disables the timer. */
	syncIntervalMinutes: number;
	/** Sync automatically when Obsidian starts. */
	syncOnStartup: boolean;
	/** Save attachments alongside notes under an `_attachments` subfolder. */
	downloadAttachments: boolean;
	/** Max attachment size to download, in MB (0 = no limit). */
	maxAttachmentMB: number;
	/** Verbose console logging under the `[obs-mail]` prefix. */
	debugLogging: boolean;

	// --- Calendar ---
	/** Fetch calendar events and show the Upcoming sidebar. */
	syncCalendar: boolean;
	/** Rolling window size in days for upcoming events. */
	calendarDaysAhead: number;
	/** Vault subfolder (under targetFolder) for meeting notes. */
	calendarSubfolder: string;
	/** Cross-link meeting notes to related email notes. */
	linkRelatedEmails: boolean;

	// --- Teams ---
	/** Sync Teams messages (chats and/or channels) as per-conversation transcripts. */
	syncTeams: boolean;
	/** Sync 1:1 and group chats (Chat.Read). */
	teamsChats: boolean;
	/** Sync team channel messages (ChannelMessage.Read.All — usually needs admin consent). */
	teamsChannels: boolean;
	/** Vault subfolder (under targetFolder) for conversation transcripts. */
	teamsSubfolder: string;
	/** Only sync messages from the last N days (0 = all, still bounded by syncSince). */
	teamsDaysBack: number;

	// --- Persisted runtime state (not user-editable in the UI) ---
	/** Encrypted-at-rest is NOT applied; see README security note. */
	refreshToken: string | null;
	/** Delta links keyed by resolved folder id. */
	deltaLinks: Record<string, string>;
	/** Thread index keyed by `${vaultSubfolder}::${conversationId}`. */
	threads: Record<string, ThreadEntry>;
	/** Event id → note path, for reconciling removed/cancelled meetings. */
	calendarNotes: Record<string, string>;
	/** Cached upcoming events for the sidebar (rebuilt each calendar sync). */
	upcomingCache: UpcomingEvent[];
	/** Teams delta links keyed by conversation key ("chat::id" / "channel::team::chan"). */
	teamsDelta: Record<string, string>;
	/** Teams conversation transcripts state, keyed by conversation key. */
	teamsConvos: Record<string, TeamsConversation>;
	/** ISO timestamp of last successful sync. */
	lastSync: string | null;
}

/** Per-conversation transcript bookkeeping. The message bodies live in the
 * note file; we persist only enough to append incrementally and rebuild the index. */
export interface TeamsConversation {
	key: string;
	kind: "chat" | "channel";
	title: string;
	subfolder: string;
	notePath: string;
	participants: string[];
	/** ISO createdDateTime of the most recent message appended to the note. */
	lastWrittenTs: string;
	/** ISO of last activity, for index sorting. */
	lastActivity: string;
	messageCount: number;
}

/** Normalized Teams chat/channel message (parsed from a Graph chatMessage). */
export interface TeamsMessage {
	id: string;
	createdIso: string;
	from: string;
	bodyHtml?: string;
	bodyText?: string;
	attachments: string[];
	deleted: boolean;
	/** True for join/leave/system event messages we skip. */
	system: boolean;
}

export interface ThreadEntry {
	conversationId: string;
	vaultSubfolder: string;
	subject: string;
	messages: ThreadMessageRef[];
}

export interface ThreadMessageRef {
	id: string;
	notePath: string;
	from: string;
	receivedIso: string;
	subject: string;
	/** Lowercased participant emails (from + to + cc), for calendar linking. */
	people?: string[];
}

/** Subset of a Microsoft Graph `message` resource we consume. */
export interface GraphMessage {
	id: string;
	conversationId?: string;
	subject?: string;
	bodyPreview?: string;
	body?: { contentType: "html" | "text"; content: string };
	from?: GraphRecipient;
	sender?: GraphRecipient;
	toRecipients?: GraphRecipient[];
	ccRecipients?: GraphRecipient[];
	receivedDateTime?: string;
	sentDateTime?: string;
	hasAttachments?: boolean;
	isRead?: boolean;
	importance?: string;
	webLink?: string;
	categories?: string[];
	/** Present on delta responses when a message was removed. */
	"@removed"?: { reason: string };
}

export interface GraphRecipient {
	emailAddress?: { name?: string; address?: string };
}

export interface GraphMailFolder {
	id: string;
	displayName: string;
	childFolderCount?: number;
}

export interface GraphAttachment {
	id: string;
	name: string;
	contentType: string;
	size: number;
	"@odata.type": string;
	/** base64 for #microsoft.graph.fileAttachment */
	contentBytes?: string;
}

/** Subset of a Microsoft Graph `event` resource we consume. */
export interface GraphEvent {
	id: string;
	subject?: string;
	bodyPreview?: string;
	body?: { contentType: "html" | "text"; content: string };
	start?: GraphDateTime;
	end?: GraphDateTime;
	isAllDay?: boolean;
	isCancelled?: boolean;
	location?: { displayName?: string };
	organizer?: GraphRecipient;
	attendees?: GraphAttendee[];
	onlineMeeting?: { joinUrl?: string };
	onlineMeetingUrl?: string;
	webLink?: string;
	seriesMasterId?: string;
	type?: string;
	categories?: string[];
}

export interface GraphDateTime {
	dateTime: string;
	timeZone: string;
}

export interface GraphAttendee {
	emailAddress?: { name?: string; address?: string };
	type?: string;
	status?: { response?: string };
}

/** Lightweight summary persisted for the Upcoming sidebar. */
export interface UpcomingEvent {
	id: string;
	subject: string;
	startIso: string; // UTC ISO
	endIso: string; // UTC ISO
	isAllDay: boolean;
	location: string;
	notePath: string | null;
	onlineUrl: string;
	webLink: string;
}

export const DEFAULT_SETTINGS: PluginSettings = {
	clientId: "",
	tenant: "common",
	targetFolder: "10-Mailbox",
	folders: [
		{ outlookPath: "inbox", vaultSubfolder: "Inbox", enabled: true },
	],
	syncIntervalMinutes: 15,
	syncOnStartup: true,
	syncSince: "",
	downloadAttachments: false,
	maxAttachmentMB: 10,
	debugLogging: false,
	syncCalendar: false,
	calendarDaysAhead: 14,
	calendarSubfolder: "Calendar",
	linkRelatedEmails: true,
	syncTeams: false,
	teamsChats: true,
	teamsChannels: false,
	teamsSubfolder: "Teams",
	teamsDaysBack: 30,
	refreshToken: null,
	deltaLinks: {},
	threads: {},
	calendarNotes: {},
	upcomingCache: [],
	teamsDelta: {},
	teamsConvos: {},
	lastSync: null,
};

/**
 * The delegated scopes to request. Teams scopes are only requested when the
 * matching feature is enabled, so users who never touch Teams aren't prompted
 * for the channel-message permission (which usually needs admin consent).
 */
export function scopesFor(s: PluginSettings): string[] {
	const scopes = ["offline_access", "User.Read", "Mail.Read", "Calendars.Read"];
	if (s.syncTeams && s.teamsChats) scopes.push("Chat.Read");
	if (s.syncTeams && s.teamsChannels) {
		scopes.push("Team.ReadBasic.All", "Channel.ReadBasic.All", "ChannelMessage.Read.All");
	}
	return scopes;
}
