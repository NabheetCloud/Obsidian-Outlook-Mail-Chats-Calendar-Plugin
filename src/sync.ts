import { GraphClient, GraphError } from "./graph";
import { NoteWriter } from "./notes";
import {
	FolderMapping,
	GraphDateTime,
	GraphEvent,
	GraphMessage,
	PluginSettings,
	TeamsConversation,
	TeamsMessage,
	ThreadEntry,
	ThreadMessageRef,
	UpcomingEvent,
} from "./types";

interface EmailRef {
	notePath: string;
	receivedIso: string;
}
interface EmailIndex {
	byPerson: Map<string, EmailRef[]>;
	bySubject: Map<string, EmailRef[]>;
}
import { info, log, warn, SyncCancelledError } from "./log";

export interface SyncReport {
	added: number;
	updated: number;
	removed: number;
	folders: number;
	events: number;
	teamsMessages: number;
	errors: string[];
	/** True when the user stopped the sync before it finished. */
	cancelled: boolean;
}

export interface SyncProgress {
	folder: string;
	phase: "fetching" | "writing";
	processed: number;
	/** Total known once fetch completes; null while still fetching. */
	total: number | null;
}

function threadKey(vaultSubfolder: string, conversationId: string): string {
	return `${vaultSubfolder}::${conversationId || "no-conversation"}`;
}

function vaultSubfolderFor(m: FolderMapping): string {
	return (m.vaultSubfolder || m.outlookPath).replace(/[\\/:*?"<>|]/g, "_").trim() || "Mail";
}

export class SyncEngine {
	private running = false;
	private cancelRequested = false;

	constructor(
		private settings: PluginSettings,
		private graph: GraphClient,
		private notes: NoteWriter,
		private save: () => Promise<void>,
		private onProgress: (p: SyncProgress) => void = () => {},
	) {}

	get isRunning(): boolean {
		return this.running;
	}

	get isCancelling(): boolean {
		return this.running && this.cancelRequested;
	}

	/**
	 * Cooperative stop. Sets a flag the fetch-paging and write loops check
	 * between items/pages, so the current sync unwinds cleanly at the next
	 * checkpoint (no half-written note, and the interrupted folder's delta
	 * token is left untouched so the next sync re-fetches it). No-op if idle.
	 */
	requestCancel(): void {
		if (this.running) {
			this.cancelRequested = true;
			info("Stop requested; sync will halt at the next checkpoint.");
		}
	}

	/** Syncs every enabled folder. Returns aggregate counts. */
	async syncAll(): Promise<SyncReport> {
		if (this.running) throw new Error("A sync is already in progress.");
		if (!this.graph.isAuthenticated) throw new Error("Not signed in.");
		this.running = true;
		this.cancelRequested = false;
		const report: SyncReport = {
			added: 0,
			updated: 0,
			removed: 0,
			folders: 0,
			events: 0,
			teamsMessages: 0,
			errors: [],
			cancelled: false,
		};
		try {
			const enabled = this.settings.folders.filter((f) => f.enabled && f.outlookPath.trim());
			info(`Sync started: ${enabled.length} folder(s).`);
			for (const mapping of enabled) {
				if (this.cancelRequested) {
					report.cancelled = true;
					break;
				}
				try {
					await this.syncFolder(mapping, report);
					report.folders++;
				} catch (e) {
					if (e instanceof SyncCancelledError) {
						report.cancelled = true;
						break;
					}
					const msg = `Folder "${mapping.outlookPath}": ${(e as Error).message}`;
					warn(msg);
					report.errors.push(msg);
				}
			}

			if (!report.cancelled && this.settings.syncCalendar) {
				try {
					await this.syncCalendarStep(report);
				} catch (e) {
					if (e instanceof SyncCancelledError) {
						report.cancelled = true;
					} else {
						const msg = `Calendar: ${(e as Error).message}`;
						warn(msg);
						report.errors.push(msg);
					}
				}
			}

			if (!report.cancelled && this.settings.syncTeams) {
				try {
					await this.syncTeamsStep(report);
				} catch (e) {
					if (e instanceof SyncCancelledError) {
						report.cancelled = true;
					} else {
						const msg = `Teams: ${(e as Error).message}`;
						warn(msg);
						report.errors.push(msg);
					}
				}
			}

			// Persist whatever we managed to write (thread state, delta links for
			// folders that finished). lastSync is only stamped on a full run so a
			// stopped sync doesn't masquerade as an up-to-date one.
			if (!report.cancelled) this.settings.lastSync = new Date().toISOString();
			await this.save();
			info(
				(report.cancelled ? "Sync stopped: " : "Sync complete: ") +
					`+${report.added} new, ${report.updated} updated, ` +
					`${report.removed} removed across ${report.folders} folder(s); ` +
					`${report.events} event(s); ${report.teamsMessages} Teams msg(s); ` +
					`${report.errors.length} error(s).`,
			);
			return report;
		} finally {
			this.running = false;
			this.cancelRequested = false;
		}
	}

	private async syncFolder(mapping: FolderMapping, report: SyncReport): Promise<void> {
		const vaultSubfolder = vaultSubfolderFor(mapping);
		info(`Syncing "${mapping.outlookPath}" → ${vaultSubfolder}`);
		const folderId = await this.graph.resolveFolderId(mapping.outlookPath);

		const onFetch = (fetched: number) =>
			this.onProgress({ folder: vaultSubfolder, phase: "fetching", processed: fetched, total: null });

		const shouldCancel = () => this.cancelRequested;
		const priorDelta = this.settings.deltaLinks[folderId] ?? null;
		let messages: GraphMessage[];
		let deltaLink: string | null;
		try {
			({ messages, deltaLink } = await this.graph.deltaMessages(folderId, priorDelta, onFetch, shouldCancel));
		} catch (e) {
			// A stored delta token expires after ~30 days of inactivity → Graph
			// returns 410 Gone. This is expected: drop the token and re-enumerate
			// the folder from scratch. Not surfaced as a user error.
			if (e instanceof GraphError && e.status === 410 && priorDelta) {
				info(`Delta token expired for "${vaultSubfolder}"; re-enumerating from scratch.`);
				delete this.settings.deltaLinks[folderId];
				({ messages, deltaLink } = await this.graph.deltaMessages(folderId, null, onFetch, shouldCancel));
			} else {
				throw e;
			}
		}

		const cutoff = validCutoff(this.settings.syncSince);
		const touchedThreadKeys = new Set<string>();
		const total = messages.length;
		let done = 0;
		let skipped = 0;

		let cancelled = false;
		for (const msg of messages) {
			if (this.cancelRequested) {
				cancelled = true;
				info(`"${vaultSubfolder}": stopped by user after writing ${done}/${total}.`);
				break;
			}
			if (msg["@removed"]) {
				const removed = this.removeMessage(msg.id, touchedThreadKeys);
				if (removed) report.removed++;
			} else if (cutoff && messageDate(msg) < cutoff) {
				// Older than the configured "sync from" date — skip writing it.
				skipped++;
			} else {
				try {
					const isUpdate = this.messageExists(msg.id);
					const { ref } = await this.notes.writeMessage(msg, vaultSubfolder);
					this.upsertThread(vaultSubfolder, msg, ref, touchedThreadKeys);

					if (this.settings.downloadAttachments && msg.hasAttachments) {
						await this.saveAttachments(vaultSubfolder, msg, report);
					}
					if (isUpdate) report.updated++;
					else report.added++;
				} catch (e) {
					// One bad message must not abort the whole folder.
					const msgErr = `Message "${msg.subject ?? msg.id}": ${(e as Error).message}`;
					warn(msgErr);
					report.errors.push(msgErr);
				}
			}
			done++;
			if (done % 10 === 0 || done === total) {
				this.onProgress({ folder: vaultSubfolder, phase: "writing", processed: done, total });
				if (done % 100 === 0 || done === total) {
					info(`${vaultSubfolder}: wrote ${done}/${total}`);
				}
			}
		}

		// Persist the new delta link so next run is incremental — but ONLY if the
		// folder finished. On a stopped sync the remaining messages weren't
		// written, so keeping the old (or no) token forces a re-fetch next time
		// rather than silently skipping them.
		if (deltaLink && !cancelled) this.settings.deltaLinks[folderId] = deltaLink;

		info(
			`"${vaultSubfolder}" ${cancelled ? "stopped" : "done"}: ${done}/${total} processed` +
				(skipped ? `, ${skipped} skipped (older than ${cutoff})` : ""),
		);

		// Regenerate the thread index for any thread that changed in this folder,
		// so the notes we did write are indexed even on a partial run.
		if (touchedThreadKeys.size > 0) {
			const threads = this.threadsForSubfolder(vaultSubfolder);
			await this.notes.writeThreadIndex(vaultSubfolder, threads);
		}

		if (cancelled) throw new SyncCancelledError();
	}

	/**
	 * Fetches the upcoming calendar window, writes one note per event (with
	 * related-email links), reconciles cancelled/expired events, and refreshes
	 * the Upcoming sidebar cache. No delta — the window is small and refetched
	 * whole each sync.
	 */
	private async syncCalendarStep(report: SyncReport): Promise<void> {
		const subfolder =
			(this.settings.calendarSubfolder || "Calendar").replace(/[\\/:*?"<>|]/g, "_").trim() ||
			"Calendar";
		const days = Math.max(1, this.settings.calendarDaysAhead || 14);
		const now = new Date();
		const startIso = now.toISOString();
		const endIso = new Date(now.getTime() + days * 86_400_000).toISOString();

		const events = await this.graph.calendarView(
			startIso,
			endIso,
			(n) => this.onProgress({ folder: "Calendar", phase: "fetching", processed: n, total: null }),
			() => this.cancelRequested,
		);

		const index = this.settings.linkRelatedEmails ? this.buildEmailIndex() : null;
		const currentIds = new Set<string>();
		const upcoming: UpcomingEvent[] = [];
		const total = events.length;
		let done = 0;

		for (const ev of events) {
			if (this.cancelRequested) {
				info(`Calendar: stopped by user after ${done}/${total}.`);
				throw new SyncCancelledError();
			}
			done++;
			if (done % 10 === 0 || done === total) {
				this.onProgress({ folder: "Calendar", phase: "writing", processed: done, total });
			}
			if (ev.isCancelled) continue; // left out of currentIds → note reconciled away

			const evStart = normalizeEventTime(ev.start);
			const evEnd = normalizeEventTime(ev.end);
			const related = index ? this.matchRelatedEmails(ev, evStart, index) : [];
			const notePath = await this.notes.writeEvent(ev, evStart, evEnd, subfolder, related);
			this.settings.calendarNotes[ev.id] = notePath;
			currentIds.add(ev.id);
			report.events++;

			upcoming.push({
				id: ev.id,
				subject: ev.subject?.trim() || "(no title)",
				startIso: evStart,
				endIso: evEnd,
				isAllDay: ev.isAllDay ?? false,
				location: ev.location?.displayName || "",
				notePath,
				onlineUrl: ev.onlineMeeting?.joinUrl || ev.onlineMeetingUrl || "",
				webLink: ev.webLink || "",
				organiser: ev.organizer?.emailAddress
					? {
						name: ev.organizer.emailAddress.name || "",
						email: ev.organizer.emailAddress.address || "",
					}
					: null,
				attendees: (ev.attendees ?? [])
					.map((a) => ({
						name: a.emailAddress?.name || "",
						email: a.emailAddress?.address || "",
						type: a.type,
						response: a.status?.response,
					}))
					.filter((a) => a.name || a.email),
			});
		}

		// Reconcile: any previously-written event no longer in the window
		// (cancelled, moved out, past) gets its note trashed.
		for (const [id, notePath] of Object.entries(this.settings.calendarNotes)) {
			if (!currentIds.has(id)) {
				await this.notes.deleteEventNote(notePath);
				delete this.settings.calendarNotes[id];
			}
		}

		upcoming.sort((a, b) => (a.startIso < b.startIso ? -1 : 1));
		this.settings.upcomingCache = upcoming;
		info(`Calendar done: ${upcoming.length} upcoming event(s).`);
	}

	/**
	 * Syncs Teams chats and/or channels as per-conversation transcript notes.
	 * Messages are appended to the note; only compact state (delta link + last
	 * timestamp) is persisted, so an active account doesn't bloat data.json.
	 */
	private async syncTeamsStep(report: SyncReport): Promise<void> {
		const subfolder =
			(this.settings.teamsSubfolder || "Teams").replace(/[\\/:*?"<>|]/g, "_").trim() || "Teams";
		const cutoffIso = this.teamsCutoffIso();
		info(`Teams sync: messages since ${cutoffIso || "the beginning"}.`);

		if (this.settings.teamsChats) {
			const chats = await this.gatherChats();
			let i = 0;
			for (const c of chats) {
				if (this.cancelRequested) throw new SyncCancelledError();
				i++;
				info(`Teams chat ${i}/${chats.length}: "${c.title}"…`);
				this.onProgress({ folder: "Teams chats", phase: "fetching", processed: i, total: chats.length });
				try {
					await this.syncConversation(
						c.key,
						"chat",
						c.title,
						c.participants,
						subfolder,
						cutoffIso,
						(_deltaLink, sinceIso) =>
							this.graph.chatMessagesList(
								c.chatId,
								sinceIso,
								(n) => this.onProgress({ folder: c.title, phase: "fetching", processed: n, total: null }),
								() => this.cancelRequested,
							),
						null,
						report,
					);
				} catch (e) {
					if (e instanceof SyncCancelledError) throw e;
					const msg = `Teams chat "${c.title}": ${(e as Error).message}`;
					warn(msg);
					report.errors.push(msg);
				}
			}
		}

		if (this.settings.teamsChannels) {
			info("Teams: listing channels…");
			const channels = await this.gatherChannels();
			info(`Teams: ${channels.length} channel(s) total.`);
			let i = 0;
			for (const ch of channels) {
				if (this.cancelRequested) throw new SyncCancelledError();
				i++;
				info(`Teams channel ${i}/${channels.length}: "${ch.title}"…`);
				this.onProgress({
					folder: "Teams channels",
					phase: "fetching",
					processed: i,
					total: channels.length,
				});
				try {
					await this.syncConversation(
						ch.key,
						"channel",
						ch.title,
						[],
						subfolder,
						cutoffIso,
						(deltaLink, _sinceIso) =>
							this.graph.channelMessagesDelta(
								ch.teamId,
								ch.channelId,
								deltaLink,
								(n) => this.onProgress({ folder: ch.title, phase: "fetching", processed: n, total: null }),
								() => this.cancelRequested,
							),
						(msgId) => this.graph.listReplies(ch.teamId, ch.channelId, msgId),
						report,
					);
				} catch (e) {
					if (e instanceof SyncCancelledError) throw e;
					const msg = `Teams channel "${ch.title}": ${(e as Error).message}`;
					warn(msg);
					report.errors.push(msg);
				}
			}
		}

		const convos = Object.values(this.settings.teamsConvos).filter((c) => c.subfolder === subfolder);
		if (convos.length) await this.notes.writeConversationIndex(subfolder, convos);
		info(`Teams done: ${report.teamsMessages} message(s) across ${convos.length} conversation(s).`);
	}

	/** Enumerates chats into conversation descriptors. */
	private async gatherChats(): Promise<
		{ key: string; chatId: string; title: string; participants: string[] }[]
	> {
		const chats = await this.graph.listChats();
		return chats.map((c) => {
			const title = c.topic?.trim() || c.members.join(", ") || c.id;
			return { key: `chat::${c.id}`, chatId: c.id, title, participants: c.members };
		});
	}

	/** Enumerates joined teams' channels into conversation descriptors. */
	private async gatherChannels(): Promise<
		{ key: string; teamId: string; channelId: string; title: string }[]
	> {
		const out: { key: string; teamId: string; channelId: string; title: string }[] = [];
		const teams = await this.graph.listJoinedTeams();
		for (const t of teams) {
			if (this.cancelRequested) throw new SyncCancelledError();
			const channels = await this.graph.listChannels(t.id);
			for (const ch of channels) {
				out.push({
					key: `channel::${t.id}::${ch.id}`,
					teamId: t.id,
					channelId: ch.id,
					title: `${t.displayName} / ${ch.displayName}`,
				});
			}
		}
		return out;
	}

	/**
	 * Fetches a conversation's message delta (with stale-token recovery), filters
	 * to new, in-window, non-system messages, appends them to the transcript, and
	 * persists the delta cursor. For channels, replies to freshly-seen posts are
	 * fetched and merged chronologically.
	 */
	private async syncConversation(
		key: string,
		kind: "chat" | "channel",
		title: string,
		participants: string[],
		subfolder: string,
		cutoffIso: string,
		fetchDelta: (
			deltaLink: string | null,
			sinceIso: string,
		) => Promise<{ messages: TeamsMessage[]; deltaLink: string | null }>,
		fetchReplies: ((messageId: string) => Promise<TeamsMessage[]>) | null,
		report: SyncReport,
	): Promise<void> {
		const existing = this.settings.teamsConvos[key];
		const lastWritten = existing?.lastWrittenTs ?? "";
		// Window floor for list-based fetchers (chats): never page below the sync
		// window, and never below what we've already written. Delta-based fetchers
		// (channels) ignore this and rely on the server cursor.
		const sinceIso = lastWritten > cutoffIso ? lastWritten : cutoffIso;

		const prior = this.settings.teamsDelta[key] ?? null;
		let result: { messages: TeamsMessage[]; deltaLink: string | null };
		try {
			result = await fetchDelta(prior, sinceIso);
		} catch (e) {
			// A Teams delta token expires → Graph returns 404/410. Re-enumerate.
			if (e instanceof GraphError && (e.status === 404 || e.status === 410) && prior) {
				info(`Teams delta expired for "${title}"; re-enumerating from scratch.`);
				delete this.settings.teamsDelta[key];
				result = await fetchDelta(null, sinceIso);
			} else {
				throw e;
			}
		}

		// Keep top-level posts worth writing; for channels, also pull their replies.
		const keep = (m: TeamsMessage) =>
			!m.deleted && !m.system && m.createdIso && (!cutoffIso || m.createdIso >= cutoffIso);

		let all: TeamsMessage[] = result.messages.filter(keep);
		if (fetchReplies) {
			const withReplies: TeamsMessage[] = [];
			for (const top of all) {
				withReplies.push(top);
				try {
					const replies = (await fetchReplies(top.id)).filter(keep);
					withReplies.push(...replies);
				} catch (e) {
					report.errors.push(`Replies for "${title}": ${(e as Error).message}`);
				}
			}
			all = withReplies;
		}

		// Only messages newer than what we've already written; chronological.
		const fresh = all
			.filter((m) => m.createdIso > lastWritten)
			.sort((a, b) => (a.createdIso < b.createdIso ? -1 : 1));

		// Persist the delta cursor regardless (so incremental stays cheap), unless stopped.
		if (result.deltaLink && !this.cancelRequested) this.settings.teamsDelta[key] = result.deltaLink;

		if (fresh.length === 0) return;

		const convo: TeamsConversation = existing ?? {
			key,
			kind,
			title,
			subfolder,
			notePath: this.notes.conversationNotePath(subfolder, kind, title, key),
			participants,
			lastWrittenTs: "",
			lastActivity: "",
			messageCount: 0,
		};
		convo.title = title;
		if (participants.length) convo.participants = participants;

		this.onProgress({ folder: title, phase: "writing", processed: fresh.length, total: fresh.length });
		await this.notes.appendTeamsMessages(convo, fresh);

		const newest = fresh[fresh.length - 1].createdIso;
		convo.lastWrittenTs = newest;
		convo.lastActivity = newest;
		convo.messageCount += fresh.length;
		this.settings.teamsConvos[key] = convo;
		report.teamsMessages += fresh.length;
		info(`Teams "${title}": +${fresh.length} message(s).`);
	}

	/** Earliest message timestamp to sync, from teamsDaysBack and syncSince. */
	private teamsCutoffIso(): string {
		let floor = 0;
		const since = validCutoff(this.settings.syncSince);
		if (since) floor = Date.parse(`${since}T00:00:00Z`);
		if (this.settings.teamsDaysBack > 0) {
			floor = Math.max(floor, Date.now() - this.settings.teamsDaysBack * 86_400_000);
		}
		return floor ? new Date(floor).toISOString() : "";
	}

	/** Indexes email notes by participant email and normalized subject. */
	private buildEmailIndex(): EmailIndex {
		const byPerson = new Map<string, EmailRef[]>();
		const bySubject = new Map<string, EmailRef[]>();
		const push = (map: Map<string, EmailRef[]>, key: string, ref: EmailRef) => {
			const arr = map.get(key);
			if (arr) arr.push(ref);
			else map.set(key, [ref]);
		};
		for (const t of Object.values(this.settings.threads)) {
			for (const m of t.messages) {
				const ref: EmailRef = { notePath: m.notePath, receivedIso: m.receivedIso };
				const subj = normalizeSubject(m.subject).toLowerCase();
				if (subj) push(bySubject, subj, ref);
				for (const p of m.people ?? []) push(byPerson, p, ref);
			}
		}
		return { byPerson, bySubject };
	}

	/**
	 * Related email notes for an event: exact normalized-subject matches (strong),
	 * plus emails with a shared participant received within 21 days before the
	 * meeting (to avoid over-linking every mail from a frequent contact). Capped.
	 */
	private matchRelatedEmails(ev: GraphEvent, evStartIso: string, index: EmailIndex): string[] {
		const score = new Map<string, number>();
		const bump = (notePath: string, s: number) =>
			score.set(notePath, Math.max(score.get(notePath) ?? 0, s));

		const subj = normalizeSubject(ev.subject).toLowerCase();
		if (subj) for (const r of index.bySubject.get(subj) ?? []) bump(r.notePath, 2);

		const emails = [ev.organizer, ...(ev.attendees ?? [])]
			.map((a) => a?.emailAddress?.address?.toLowerCase())
			.filter((e): e is string => !!e);
		const windowStart = new Date(
			(isNaN(Date.parse(evStartIso)) ? Date.now() : Date.parse(evStartIso)) - 21 * 86_400_000,
		).toISOString();
		for (const e of emails) {
			for (const r of index.byPerson.get(e) ?? []) {
				if (r.receivedIso >= windowStart) bump(r.notePath, 1);
			}
		}

		return [...score.entries()]
			.sort((a, b) => b[1] - a[1])
			.slice(0, 12)
			.map(([notePath]) => notePath);
	}

	private async saveAttachments(
		vaultSubfolder: string,
		msg: GraphMessage,
		report: SyncReport,
	): Promise<void> {
		try {
			const atts = await this.graph.listAttachments(msg.id);
			const limit = this.settings.maxAttachmentMB * 1024 * 1024;
			for (const a of atts) {
				if (a["@odata.type"] !== "#microsoft.graph.fileAttachment" || !a.contentBytes) continue;
				if (limit > 0 && a.size > limit) {
					log(`Skipping oversized attachment ${a.name} (${a.size} bytes).`);
					continue;
				}
				const bytes = base64ToArrayBuffer(a.contentBytes);
				await this.notes.writeAttachment(vaultSubfolder, msg.id, a.name, bytes);
			}
		} catch (e) {
			report.errors.push(`Attachments for "${msg.subject ?? msg.id}": ${(e as Error).message}`);
		}
	}

	// ---- thread bookkeeping ----

	private messageExists(id: string): boolean {
		for (const t of Object.values(this.settings.threads)) {
			if (t.messages.some((m) => m.id === id)) return true;
		}
		return false;
	}

	private upsertThread(
		vaultSubfolder: string,
		msg: GraphMessage,
		ref: ThreadMessageRef,
		touched: Set<string>,
	): void {
		const key = threadKey(vaultSubfolder, msg.conversationId ?? "");
		let entry: ThreadEntry | undefined = this.settings.threads[key];
		if (!entry) {
			entry = {
				conversationId: msg.conversationId ?? "",
				vaultSubfolder,
				subject: normalizeSubject(msg.subject),
				messages: [],
			};
			this.settings.threads[key] = entry;
		}
		const existingIdx = entry.messages.findIndex((m) => m.id === ref.id);
		if (existingIdx >= 0) entry.messages[existingIdx] = ref;
		else entry.messages.push(ref);
		// Keep the shortest (root) subject as the thread label.
		const norm = normalizeSubject(msg.subject);
		if (norm && (!entry.subject || norm.length < entry.subject.length)) entry.subject = norm;
		touched.add(key);
	}

	private removeMessage(id: string, touched: Set<string>): boolean {
		for (const [key, t] of Object.entries(this.settings.threads)) {
			const idx = t.messages.findIndex((m) => m.id === id);
			if (idx >= 0) {
				const [ref] = t.messages.splice(idx, 1);
				void this.notes.deleteMessageNote(ref.notePath);
				touched.add(key);
				if (t.messages.length === 0) delete this.settings.threads[key];
				return true;
			}
		}
		return false;
	}

	private threadsForSubfolder(vaultSubfolder: string): ThreadEntry[] {
		return Object.values(this.settings.threads).filter(
			(t) => t.vaultSubfolder === vaultSubfolder,
		);
	}

	/** Clears delta + thread state so the next sync re-enumerates from scratch. */
	resetState(): void {
		this.settings.deltaLinks = {};
		this.settings.threads = {};
		this.settings.calendarNotes = {};
		this.settings.upcomingCache = [];
		this.settings.teamsDelta = {};
		this.settings.teamsConvos = {};
		this.settings.lastSync = null;
	}
}

/** Returns a valid YYYY-MM-DD cutoff, or null if unset/malformed. */
function validCutoff(since: string): string | null {
	return /^\d{4}-\d{2}-\d{2}$/.test(since.trim()) ? since.trim() : null;
}

/** Message date as YYYY-MM-DD for comparison against the cutoff. */
function messageDate(msg: GraphMessage): string {
	return (msg.receivedDateTime ?? msg.sentDateTime ?? "").slice(0, 10);
}

/**
 * Graph returns event times in the requested UTC zone but without a suffix
 * (e.g. "2026-07-05T10:00:00.0000000"). Normalize to a real UTC ISO string.
 */
function normalizeEventTime(dt?: GraphDateTime): string {
	const raw = dt?.dateTime;
	if (!raw) return "";
	const withZone = /[zZ]|[+-]\d{2}:?\d{2}$/.test(raw) ? raw : `${raw}Z`;
	const d = new Date(withZone);
	return isNaN(d.getTime()) ? raw : d.toISOString();
}

function normalizeSubject(subject?: string): string {
	return (subject ?? "")
		.replace(/^(\s*(re|fw|fwd|aw|wg)\s*:\s*)+/i, "")
		.trim() || (subject ?? "").trim();
}

function base64ToArrayBuffer(b64: string): ArrayBuffer {
	const binary = atob(b64);
	const len = binary.length;
	const bytes = new Uint8Array(len);
	for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
	return bytes.buffer;
}
