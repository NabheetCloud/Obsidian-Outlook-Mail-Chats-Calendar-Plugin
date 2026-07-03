import { App, normalizePath, TFile, TFolder } from "obsidian";
import TurndownService from "turndown";
import {
	GraphAttendee,
	GraphEvent,
	GraphMessage,
	GraphRecipient,
	PluginSettings,
	TeamsConversation,
	TeamsMessage,
	ThreadEntry,
	ThreadMessageRef,
} from "./types";
import { log } from "./log";

const turndown = new TurndownService({
	headingStyle: "atx",
	codeBlockStyle: "fenced",
	bulletListMarker: "-",
});
// Drop noisy Outlook signature/style cruft.
turndown.remove(["style", "script"]);

export class NoteWriter {
	constructor(private app: App, private settings: PluginSettings) {}

	private get root(): string {
		return normalizePath(this.settings.targetFolder);
	}

	/** Ensures every folder along a vault-relative path exists. */
	private async ensureFolder(path: string): Promise<void> {
		const norm = normalizePath(path);
		if (norm === "" || norm === "/") return;
		const existing = this.app.vault.getAbstractFileByPath(norm);
		if (existing instanceof TFolder) return;
		// createFolder throws if any ancestor is missing, so build up.
		const parts = norm.split("/");
		let cur = "";
		for (const part of parts) {
			cur = cur ? `${cur}/${part}` : part;
			if (!this.app.vault.getAbstractFileByPath(cur)) {
				try {
					await this.app.vault.createFolder(cur);
				} catch (e) {
					// Race: another writer may have just created it.
					if (!this.app.vault.getAbstractFileByPath(cur)) throw e;
				}
			}
		}
	}

	/**
	 * Creates or overwrites a note, tolerating an "already exists" race (the
	 * vault index can briefly lag a just-created file). Never lets a single
	 * note failure crash the caller's loop.
	 */
	private async writeFile(path: string, content: string): Promise<void> {
		const existing = this.app.vault.getAbstractFileByPath(path);
		if (existing instanceof TFile) {
			await this.app.vault.modify(existing, content);
			return;
		}
		try {
			await this.app.vault.create(path, content);
		} catch (e) {
			const again = this.app.vault.getAbstractFileByPath(path);
			if (again instanceof TFile) await this.app.vault.modify(again, content);
			else throw e;
		}
	}

	/**
	 * Writes (or overwrites, on delta update) the note for a single message.
	 * Returns the vault-relative note path and a thread reference.
	 */
	async writeMessage(
		msg: GraphMessage,
		vaultSubfolder: string,
	): Promise<{ notePath: string; ref: ThreadMessageRef }> {
		const folder = normalizePath(`${this.root}/${vaultSubfolder}`);
		await this.ensureFolder(folder);

		const received = msg.receivedDateTime ?? msg.sentDateTime ?? "";
		const datePrefix = received ? received.slice(0, 10) : "no-date";
		const subject = msg.subject?.trim() || "(no subject)";
		const fileName = `${datePrefix} ${slug(subject)} ${shortId(msg.id)}.md`;
		const notePath = normalizePath(`${folder}/${fileName}`);

		const content = this.renderMarkdown(msg, vaultSubfolder);
		await this.writeFile(notePath, content);

		const people = [msg.from ?? msg.sender, ...(msg.toRecipients ?? []), ...(msg.ccRecipients ?? [])]
			.map(recipientEmail)
			.filter(Boolean)
			.map((e) => e.toLowerCase());
		const ref: ThreadMessageRef = {
			id: msg.id,
			notePath,
			from: recipientName(msg.from ?? msg.sender),
			receivedIso: received,
			subject,
			people: Array.from(new Set(people)),
		};
		return { notePath, ref };
	}

	/** Deletes the note backing a removed message, if present. */
	async deleteMessageNote(notePath: string): Promise<void> {
		const f = this.app.vault.getAbstractFileByPath(normalizePath(notePath));
		if (f instanceof TFile) {
			await this.app.fileManager.trashFile(f);
			log("Trashed removed message note:", notePath);
		}
	}

	/**
	 * Writes (or overwrites) a meeting note. `relatedNotes` are vault-relative
	 * paths of email notes to cross-link. Returns the note path.
	 */
	async writeEvent(
		ev: GraphEvent,
		startIso: string,
		endIso: string,
		vaultSubfolder: string,
		relatedNotes: string[],
	): Promise<string> {
		const folder = normalizePath(`${this.root}/${vaultSubfolder}`);
		await this.ensureFolder(folder);

		const datePrefix = startIso.slice(0, 10);
		const subject = ev.subject?.trim() || "(no title)";
		const fileName = `${datePrefix} ${slug(subject)} ${shortId(ev.id)}.md`;
		const notePath = normalizePath(`${folder}/${fileName}`);

		const content = this.renderEvent(ev, startIso, endIso, relatedNotes);
		await this.writeFile(notePath, content);
		return notePath;
	}

	async deleteEventNote(notePath: string): Promise<void> {
		const f = this.app.vault.getAbstractFileByPath(normalizePath(notePath));
		if (f instanceof TFile) {
			await this.app.fileManager.trashFile(f);
			log("Trashed removed/cancelled event note:", notePath);
		}
	}

	private renderEvent(
		ev: GraphEvent,
		startIso: string,
		endIso: string,
		relatedNotes: string[],
	): string {
		const organizer = ev.organizer;
		const attendees = ev.attendees ?? [];
		const join = ev.onlineMeeting?.joinUrl || ev.onlineMeetingUrl || "";
		const loc = ev.location?.displayName || "";

		const fm: Record<string, unknown> = {
			source: "outlook",
			type: "meeting",
			event_id: ev.id,
			title: ev.subject ?? "",
			start: startIso,
			end: endIso,
			all_day: ev.isAllDay ?? false,
			organizer: recipientEmail(organizer),
			attendees: attendees.map((a) => a.emailAddress?.address ?? "").filter(Boolean),
			location: loc,
			online_url: join,
			web_link: ev.webLink ?? "",
			categories: ev.categories ?? [],
		};

		const when = ev.isAllDay
			? `${startIso.slice(0, 10)} (all day)`
			: `${fmtLocal(startIso)} – ${fmtLocal(endIso)}`;

		let bodyMd = "";
		if (ev.body?.contentType === "html") bodyMd = turndown.turndown(ev.body.content || "");
		else bodyMd = (ev.body?.content || ev.bodyPreview || "").trim();

		const lines: string[] = [
			yaml(fm),
			"",
			`# ${escapeMd(ev.subject?.trim() || "(no title)")}`,
			"",
			`**When:** ${when}`,
			loc ? `**Where:** ${escapeMd(loc)}` : "",
			`**Organizer:** ${escapeMd(recipientDisplay(organizer))}`,
			attendees.length
				? `**Attendees:** ${attendees.map((a) => escapeMd(recipientDisplay(a))).join(", ")}`
				: "",
			join ? `**[Join online](${join})**` : "",
			ev.webLink ? `**[Open in Outlook](${ev.webLink})**` : "",
		].filter(Boolean);

		if (relatedNotes.length) {
			lines.push("", "## Related emails");
			for (const p of relatedNotes) lines.push(`- ${vaultLink(p)}`);
		}

		lines.push("", "---", "", bodyMd, "");
		return lines.join("\n");
	}

	private renderMarkdown(msg: GraphMessage, vaultSubfolder: string): string {
		const received = msg.receivedDateTime ?? msg.sentDateTime ?? "";
		const from = msg.from ?? msg.sender;
		const fm: Record<string, unknown> = {
			source: "outlook",
			message_id: msg.id,
			conversation_id: msg.conversationId ?? "",
			subject: msg.subject ?? "",
			from: recipientEmail(from),
			from_name: recipientName(from),
			to: (msg.toRecipients ?? []).map(recipientEmail),
			cc: (msg.ccRecipients ?? []).map(recipientEmail),
			received: received,
			folder: vaultSubfolder,
			is_read: msg.isRead ?? false,
			importance: msg.importance ?? "normal",
			has_attachments: msg.hasAttachments ?? false,
			categories: msg.categories ?? [],
			web_link: msg.webLink ?? "",
		};

		let bodyMd = "";
		if (msg.body?.contentType === "html") {
			bodyMd = turndown.turndown(msg.body.content || "");
		} else {
			bodyMd = (msg.body?.content || msg.bodyPreview || "").trim();
		}

		const toLine = (msg.toRecipients ?? []).map(recipientDisplay).join(", ") || "—";
		const ccLine = (msg.ccRecipients ?? []).map(recipientDisplay).join(", ");

		const header =
			`# ${escapeMd(msg.subject?.trim() || "(no subject)")}\n\n` +
			`**From:** ${escapeMd(recipientDisplay(from))}\n` +
			`**To:** ${escapeMd(toLine)}\n` +
			(ccLine ? `**Cc:** ${escapeMd(ccLine)}\n` : "") +
			`**Date:** ${received}\n` +
			(msg.webLink ? `**[Open in Outlook](${msg.webLink})**\n` : "") +
			`\n---\n\n`;

		return `${yaml(fm)}\n${header}${bodyMd}\n`;
	}

	/**
	 * Regenerates a folder's thread-index note from the in-memory thread map.
	 * Threads sorted by most recent activity; messages within a thread oldest→newest.
	 */
	async writeThreadIndex(vaultSubfolder: string, threads: ThreadEntry[]): Promise<void> {
		const folder = normalizePath(`${this.root}/${vaultSubfolder}`);
		await this.ensureFolder(folder);
		const indexPath = normalizePath(`${folder}/_Thread Index.md`);

		const sorted = threads
			.map((t) => ({
				t,
				latest: t.messages.reduce((m, x) => (x.receivedIso > m ? x.receivedIso : m), ""),
			}))
			.sort((a, b) => (a.latest < b.latest ? 1 : -1));

		const lines: string[] = [
			"---",
			"source: outlook",
			"type: thread-index",
			`folder: ${vaultSubfolder}`,
			`thread_count: ${threads.length}`,
			"---",
			"",
			`# Thread Index — ${vaultSubfolder}`,
			"",
			`> ${threads.length} conversation(s). Auto-generated by the Outlook Mailbox plugin.`,
			"",
		];

		for (const { t, latest } of sorted) {
			const msgs = [...t.messages].sort((a, b) => (a.receivedIso < b.receivedIso ? -1 : 1));
			lines.push(`## ${escapeMd(t.subject || "(no subject)")}`);
			lines.push(`*${msgs.length} message(s) · last activity ${latest.slice(0, 10) || "—"}*`);
			lines.push("");
			for (const m of msgs) {
				const link = vaultLink(m.notePath);
				lines.push(
					`- ${m.receivedIso.slice(0, 10) || "—"} — **${escapeMd(m.from)}** — ${link}`,
				);
			}
			lines.push("");
		}

		await this.writeFile(indexPath, lines.join("\n"));
	}

	/** Saves file attachments under `<subfolder>/_attachments/<shortId>/`. */
	async writeAttachment(
		vaultSubfolder: string,
		messageId: string,
		name: string,
		bytes: ArrayBuffer,
	): Promise<string> {
		const dir = normalizePath(`${this.root}/${vaultSubfolder}/_attachments/${shortId(messageId)}`);
		await this.ensureFolder(dir);
		const path = normalizePath(`${dir}/${sanitizeFileName(name)}`);
		if (!this.app.vault.getAbstractFileByPath(path)) {
			await this.app.vault.createBinary(path, bytes);
		}
		return path;
	}

	// ---- Teams transcripts ----

	/** Builds the vault-relative note path for a conversation transcript. */
	conversationNotePath(subfolder: string, kind: "chat" | "channel", title: string, key: string): string {
		const prefix = kind === "chat" ? "Chat" : "Channel";
		const fileName = `${prefix} - ${slug(title) || "Untitled"} ${shortId(key)}.md`;
		return normalizePath(`${this.root}/${subfolder}/${fileName}`);
	}

	/** Creates the transcript note with a header if it does not yet exist. */
	async ensureConversationNote(convo: TeamsConversation): Promise<void> {
		const existing = this.app.vault.getAbstractFileByPath(convo.notePath);
		if (existing instanceof TFile) return;
		await this.ensureFolder(parentPath(convo.notePath));
		const fm = yaml({
			source: "teams",
			type: "conversation",
			kind: convo.kind,
			title: convo.title,
			participants: convo.participants,
			conversation_key: convo.key,
		});
		const header =
			`${fm}\n\n# ${escapeMd(convo.title || "(untitled)")}\n\n` +
			`> Teams ${convo.kind} transcript · ${escapeMd(convo.participants.join(", "))}\n`;
		await this.writeFile(convo.notePath, header);
	}

	/**
	 * Appends rendered message blocks to a conversation's transcript (creating it
	 * first if needed). Bodies live in the note, not in settings, so state stays small.
	 */
	async appendTeamsMessages(convo: TeamsConversation, messages: TeamsMessage[]): Promise<void> {
		if (messages.length === 0) return;
		await this.ensureConversationNote(convo);
		const f = this.app.vault.getAbstractFileByPath(convo.notePath);
		if (!(f instanceof TFile)) return;
		const blocks = messages.map(renderTeamsMessage).join("\n");
		await this.app.vault.append(f, `\n${blocks}`);
	}

	/** Regenerates the conversation index for a Teams subfolder. */
	async writeConversationIndex(subfolder: string, convos: TeamsConversation[]): Promise<void> {
		const folder = normalizePath(`${this.root}/${subfolder}`);
		await this.ensureFolder(folder);
		const indexPath = normalizePath(`${folder}/_Conversation Index.md`);
		const sorted = [...convos].sort((a, b) => (a.lastActivity < b.lastActivity ? 1 : -1));

		const lines: string[] = [
			"---",
			"source: teams",
			"type: conversation-index",
			`conversation_count: ${convos.length}`,
			"---",
			"",
			"# Teams Conversations",
			"",
			`> ${convos.length} conversation(s). Auto-generated by the Outlook Mailbox plugin.`,
			"",
		];
		for (const c of sorted) {
			const link = vaultLink(c.notePath);
			const kind = c.kind === "chat" ? "chat" : "channel";
			lines.push(
				`- ${c.lastActivity.slice(0, 10) || "—"} — **${escapeMd(c.title || "(untitled)")}** ` +
					`(${kind}, ${c.messageCount} msg) — ${link}`,
			);
		}
		await this.writeFile(indexPath, lines.join("\n"));
	}
}

// ---- helpers ----

function renderTeamsMessage(m: TeamsMessage): string {
	const d = new Date(m.createdIso);
	const when = isNaN(d.getTime())
		? m.createdIso.slice(0, 16)
		: d.toLocaleString([], { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });
	let body = m.bodyHtml ? turndown.turndown(m.bodyHtml) : (m.bodyText || "").trim();
	body = body.trim() || "*(no text)*";
	const quoted = body
		.split("\n")
		.map((l) => `> ${l}`)
		.join("\n");
	const att = m.attachments.length ? `\n> 📎 ${m.attachments.map(escapeMd).join(", ")}` : "";
	return `**${escapeMd(m.from)}** · ${when}\n${quoted}${att}\n`;
}

function parentPath(path: string): string {
	const idx = path.lastIndexOf("/");
	return idx <= 0 ? "" : path.slice(0, idx);
}

function recipientName(r?: GraphRecipient): string {
	return r?.emailAddress?.name || r?.emailAddress?.address || "Unknown";
}
function recipientEmail(r?: GraphRecipient): string {
	return r?.emailAddress?.address || "";
}
function recipientDisplay(r?: GraphRecipient | GraphAttendee): string {
	const name = r?.emailAddress?.name;
	const addr = r?.emailAddress?.address;
	if (name && addr && name !== addr) return `${name} <${addr}>`;
	return addr || name || "Unknown";
}

/** Formats a UTC ISO string as a local, human-readable datetime. */
function fmtLocal(iso: string): string {
	const d = new Date(iso);
	if (isNaN(d.getTime())) return iso;
	return d.toLocaleString([], {
		weekday: "short",
		month: "short",
		day: "numeric",
		hour: "2-digit",
		minute: "2-digit",
	});
}

function slug(s: string): string {
	return s
		.replace(/[\r\n]+/g, " ")
		.replace(/[\\/:*?"<>|#^[\]]/g, "")
		.replace(/\s+/g, " ")
		.trim()
		.slice(0, 80)
		.trim();
}

function sanitizeFileName(s: string): string {
	return s.replace(/[\\/:*?"<>|]/g, "_").slice(0, 120) || "attachment";
}

function shortId(id: string): string {
	// FNV-1a 32-bit hash of the FULL id → base36. Graph ids often share the same
	// trailing characters, so a suffix slice collides across distinct messages;
	// hashing the whole id yields a stable, collision-resistant filename token.
	let h = 0x811c9dc5;
	for (let i = 0; i < id.length; i++) {
		h ^= id.charCodeAt(i);
		h = Math.imul(h, 0x01000193);
	}
	return (h >>> 0).toString(36);
}

function vaultLink(notePath: string): string {
	// Wiki-link by basename without extension for nice display.
	const base = notePath.split("/").pop() ?? notePath;
	const noExt = base.replace(/\.md$/, "");
	return `[[${noExt}]]`;
}

function escapeMd(s: string): string {
	return s.replace(/([\\`*_{}\[\]])/g, "\\$1");
}

function yaml(obj: Record<string, unknown>): string {
	const lines = ["---"];
	for (const [k, v] of Object.entries(obj)) {
		if (Array.isArray(v)) {
			if (v.length === 0) {
				lines.push(`${k}: []`);
			} else {
				lines.push(`${k}:`);
				for (const item of v) lines.push(`  - ${yamlScalar(item)}`);
			}
		} else {
			lines.push(`${k}: ${yamlScalar(v)}`);
		}
	}
	lines.push("---");
	return lines.join("\n");
}

function yamlScalar(v: unknown): string {
	if (typeof v === "boolean" || typeof v === "number") return String(v);
	const s = String(v ?? "");
	if (s === "") return '""';
	// Quote if it contains YAML-significant characters.
	if (/[:#\[\]{}",&*!|>'%@`]/.test(s) || /^\s|\s$/.test(s)) {
		return `"${s.replace(/"/g, '\\"')}"`;
	}
	return s;
}
