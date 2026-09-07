import { normalizePath, Notice, Plugin, TFile, TFolder } from "obsidian";
import { DEFAULT_SETTINGS, PluginSettings, UpcomingEvent, UpcomingPerson } from "./types";
import { GraphClient } from "./graph";
import { NoteWriter } from "./notes";
import { SyncEngine, SyncProgress } from "./sync";
import { OutlookMailboxSettingTab } from "./settings";
import { UpcomingView, VIEW_TYPE_UPCOMING } from "./views/upcoming";
import { setDebug, log, logError } from "./log";

export default class OutlookMailboxPlugin extends Plugin {
	settings!: PluginSettings;
	graph!: GraphClient;
	notes!: NoteWriter;
	sync!: SyncEngine;
	connectedAs: string | null = null;

	private statusEl: HTMLElement | null = null;
	private timer: number | null = null;

	async onload(): Promise<void> {
		await this.loadSettings();
		this.applyDebug();

		this.graph = new GraphClient(
			this.settings,
			(url) => window.open(url, "_blank"),
			async (refreshToken) => {
				this.settings.refreshToken = refreshToken;
				await this.saveSettings();
			},
		);
		this.notes = new NoteWriter(this.app, this.settings);
		this.sync = new SyncEngine(
			this.settings,
			this.graph,
			this.notes,
			() => this.saveSettings(),
			(p) => this.onSyncProgress(p),
		);

		this.addSettingTab(new OutlookMailboxSettingTab(this.app, this));

		this.registerView(VIEW_TYPE_UPCOMING, (leaf) => new UpcomingView(leaf, this));

		this.addRibbonIcon("mail", "Sync Outlook mailbox", () => this.runSync());
		this.addRibbonIcon("calendar-clock", "Outlook: upcoming meetings", () =>
			this.activateUpcomingView(),
		);

		this.addCommand({
			id: "outlook-sync-now",
			name: "Sync now",
			callback: () => this.runSync(),
		});
		this.addCommand({
			id: "outlook-stop-sync",
			name: "Stop sync",
			checkCallback: (checking) => {
				if (!this.sync.isRunning) return false;
				if (!checking) this.stopSync();
				return true;
			},
		});
		this.addCommand({
			id: "outlook-connect",
			name: "Connect account",
			callback: () => {
				this.connect().catch((e: unknown) =>
					new Notice(`Connect failed: ${e instanceof Error ? e.message : String(e)}`),
				);
			},
		});
		this.addCommand({
			id: "outlook-open-upcoming",
			name: "Open upcoming meetings",
			callback: () => this.activateUpcomingView(),
		});

		this.statusEl = this.addStatusBarItem();
		this.updateStatus(this.settings.lastSync ? "idle" : "not-synced");

		this.rescheduleTimer();

		if (this.settings.syncOnStartup && this.graph.isAuthenticated) {
			// Defer so startup isn't blocked.
			this.app.workspace.onLayoutReady(() => {
				window.setTimeout(() => this.runSync(true), 3000);
			});
		}

		log("Plugin loaded.");
	}

	onunload(): void {
		if (this.timer !== null) window.clearInterval(this.timer);
	}

	async connect(): Promise<void> {
		await this.graph.login();
		try {
			const me = await this.graph.me();
			this.connectedAs = me.mail || me.userPrincipalName || me.displayName || null;
		} catch {
			this.connectedAs = null;
		}
		this.updateStatus("idle");
	}

	async runSync(silent = false): Promise<void> {
		if (!this.graph.isAuthenticated) {
			if (!silent) new Notice("Outlook Teams and Calendar: not connected. Open settings to connect.");
			return;
		}
		if (this.sync.isRunning) {
			if (!silent) new Notice("Outlook Teams and Calendar: a sync is already running.");
			return;
		}
		this.updateStatus("syncing");
		try {
			const report = await this.sync.syncAll();
			this.updateStatus("idle");
			const teamsPart = report.teamsMessages ? `, ${report.teamsMessages} Teams msg` : "";
			const summary =
				`${report.cancelled ? "Outlook (stopped): +" : "Outlook: +"}${report.added} new, ` +
				`${report.updated} updated, ${report.removed} removed across ` +
				`${report.folders} folder(s)${teamsPart}.`;
			if (
				report.cancelled ||
				!silent ||
				report.added ||
				report.updated ||
				report.removed ||
				report.teamsMessages
			) {
				new Notice(summary);
			}
			if (report.errors.length) {
				new Notice(`Outlook: ${report.errors.length} error(s). See console.`);
				report.errors.forEach((e) => logError(e));
			}
			this.refreshUpcomingView();
		} catch (e) {
			this.updateStatus("error");
			logError("Sync failed:", e);
			if (!silent) new Notice(`Outlook sync failed: ${(e as Error).message}`);
		}
	}

	/** Cooperatively stops an in-flight sync. Safe to call when idle. */
	stopSync(): void {
		if (!this.sync.isRunning) {
			new Notice("Outlook Teams and Calendar: no sync is running.");
			return;
		}
		this.sync.requestCancel();
		new Notice("Outlook Teams and Calendar: stopping sync…");
		this.updateStatus("stopping");
	}

	/** Opens (or reveals) the Upcoming meetings sidebar view. */
	async activateUpcomingView(): Promise<void> {
		const { workspace } = this.app;
		let leaf = workspace.getLeavesOfType(VIEW_TYPE_UPCOMING)[0];
		if (!leaf) {
			leaf = workspace.getRightLeaf(false) ?? workspace.getLeaf(true);
			await leaf.setViewState({ type: VIEW_TYPE_UPCOMING, active: true });
		}
		await workspace.revealLeaf(leaf);
	}

	/** Re-renders any open Upcoming views from the refreshed cache. */
	refreshUpcomingView(): void {
		for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_UPCOMING)) {
			const view = leaf.view;
			if (view instanceof UpcomingView) view.render();
		}
	}

	/** Create a durable meeting note from an Outlook calendar event, or open the existing one. */
	async createOrOpenMeetingNote(ev: UpcomingEvent): Promise<void> {
		const rawTemplatePath = this.settings.meetingTemplatePath.trim();
		const rawMeetingFolder = this.settings.meetingNotesFolder.trim();
		if (!rawTemplatePath) {
			new Notice("Outlook: set Permanent meeting template in plugin settings first.");
			return;
		}
		if (!rawMeetingFolder) {
			new Notice("Outlook: set Permanent meeting notes folder in plugin settings first.");
			return;
		}

		const templatePath = normalizePath(rawTemplatePath);
		const meetingFolder = normalizePath(rawMeetingFolder);
		const existing = this.findPermanentMeetingNote(ev.id, meetingFolder);
		if (existing) {
			await this.updatePermanentMeetingFrontmatter(existing, ev);
			await this.app.workspace.getLeaf(false).openFile(existing);
			return;
		}

		const template = this.app.vault.getAbstractFileByPath(templatePath);
		if (!(template instanceof TFile)) {
			new Notice(`Outlook: meeting template not found: ${templatePath}`);
			return;
		}

		await this.ensureVaultFolder(meetingFolder);
		const templateBody = neutralizeLegacyMeetingTemplate(await this.app.vault.read(template));
		const path = this.uniqueMeetingPath(meetingFolder, ev);
		const file = await this.app.vault.create(path, templateBody);
		await this.updatePermanentMeetingFrontmatter(file, ev);
		await this.app.workspace.getLeaf(false).openFile(file);
	}

	private findPermanentMeetingNote(eventId: string, meetingFolder: string): TFile | null {
		const prefix = `${meetingFolder}/`;
		for (const file of this.app.vault.getMarkdownFiles()) {
			if (!file.path.startsWith(prefix)) continue;
			const fm = this.app.metadataCache.getFileCache(file)?.frontmatter;
			if (String(fm?.outlook_event_id ?? "") === eventId) return file;
		}
		return null;
	}

	private async updatePermanentMeetingFrontmatter(file: TFile, ev: UpcomingEvent): Promise<void> {
		await this.app.fileManager.processFrontMatter(file, (fm) => {
			const start = new Date(ev.startIso);
			const end = new Date(ev.endIso);
			const date = localDate(start);
			const startDateTime = localDateTime(start);
			const endDateTime = localDateTime(end);
			const startTime = ev.isAllDay ? "" : localTime(start);

			fm.type = fm.type || "meeting";
			fm.status = fm.status || "open";
			fm.meeting_date = date;
			fm.start = startDateTime;
			fm.end = endDateTime;
			fm.organiser = formatPerson(ev.organiser ?? null);
			fm.Attendees = (ev.attendees ?? []).map(formatPerson).filter(Boolean);
			fm.outlook_event_id = ev.id;
			fm.outlook_link = ev.webLink || "";
			fm.outlook_subject = ev.subject;
			fm.location = ev.location || "";
			fm.online_url = ev.onlineUrl || "";

			// Retain compatibility with fields from the previous template attempt.
			fm["outlook date"] = date;
			fm.outlook_time = startTime;
			fm.calendar_event = ev.webLink || "";
		});
	}

	private async ensureVaultFolder(folder: string): Promise<void> {
		const parts = normalizePath(folder).split("/").filter(Boolean);
		let current = "";
		for (const part of parts) {
			current = current ? `${current}/${part}` : part;
			const existing = this.app.vault.getAbstractFileByPath(current);
			if (existing instanceof TFolder) continue;
			if (existing) throw new Error(`Cannot create meeting folder; a file exists at ${current}`);
			await this.app.vault.createFolder(current);
		}
	}

	private uniqueMeetingPath(folder: string, ev: UpcomingEvent): string {
		const start = new Date(ev.startIso);
		const stamp = meetingFileStamp(start);
		const subject = shortenMeetingSubject(sanitizeFileName(ev.subject) || "Meeting", 60);
		const base = `${subject}  [${stamp}]`;
		let path = normalizePath(`${folder}/${base}.md`);
		if (!this.app.vault.getAbstractFileByPath(path)) return path;

		// Extremely unlikely fallback: two different events at the exact same minute
		// with the same subject. Permanent-note deduplication still uses outlook_event_id.
		const suffix = ev.id.replace(/[^a-zA-Z0-9]/g, "").slice(-8) || Date.now().toString(36);
		return normalizePath(`${folder}/${base} ${suffix}.md`);
	}
	rescheduleTimer(): void {
		if (this.timer !== null) {
			window.clearInterval(this.timer);
			this.timer = null;
		}
		const mins = this.settings.syncIntervalMinutes;
		if (mins > 0) {
			this.timer = window.setInterval(() => this.runSync(true), mins * 60 * 1000);
			this.registerInterval(this.timer);
			log(`Auto-sync scheduled every ${mins} min.`);
		}
	}

	applyDebug(): void {
		setDebug(this.settings.debugLogging);
	}

	private onSyncProgress(p: SyncProgress): void {
		if (!this.statusEl) return;
		const count = p.total != null ? `${p.processed}/${p.total}` : `${p.processed}`;
		const verb = p.phase === "fetching" ? "fetching" : "writing";
		this.statusEl.setText(`📬 ${p.folder}: ${verb} ${count}…`);
	}

	private updateStatus(state: "idle" | "syncing" | "stopping" | "error" | "not-synced"): void {
		if (!this.statusEl) return;
		const last = this.settings.lastSync
			? new Date(this.settings.lastSync).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
			: "never";
		const label: Record<typeof state, string> = {
			idle: `📬 Outlook · ${last}`,
			syncing: "📬 Outlook · syncing…",
			stopping: "📪 Outlook · stopping…",
			error: "📭 Outlook · error",
			"not-synced": "📭 Outlook · not synced",
		};
		this.statusEl.setText(label[state]);
		this.statusEl.title = "Click the ribbon mail icon to sync";
	}

	async loadSettings(): Promise<void> {
		const data = ((await this.loadData()) as Partial<PluginSettings> | null) ?? {};
		this.settings = Object.assign({}, DEFAULT_SETTINGS, data);
		// Ensure nested objects exist even on partial old data.
		this.settings.deltaLinks = this.settings.deltaLinks ?? {};
		this.settings.threads = this.settings.threads ?? {};
		this.settings.calendarNotes = this.settings.calendarNotes ?? {};
		this.settings.upcomingCache = this.settings.upcomingCache ?? [];
		this.settings.teamsDelta = this.settings.teamsDelta ?? {};
		this.settings.teamsConvos = this.settings.teamsConvos ?? {};
		this.settings.folders = this.settings.folders?.length
			? this.settings.folders
			: DEFAULT_SETTINGS.folders.map((f) => ({ ...f }));
	}

	async saveSettings(): Promise<void> {
		await this.saveData(this.settings);
	}
}
function formatPerson(person: UpcomingPerson | null): string {
	if (!person) return "";
	const name = person.name.trim();
	const email = person.email.trim();
	if (name && email) return `${name} <${email}>`;
	return name || email;
}

function localDate(d: Date): string {
	if (isNaN(d.getTime())) return "";
	const y = d.getFullYear();
	const m = String(d.getMonth() + 1).padStart(2, "0");
	const day = String(d.getDate()).padStart(2, "0");
	return `${y}-${m}-${day}`;
}

function localTime(d: Date): string {
	if (isNaN(d.getTime())) return "";
	return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

function localDateTime(d: Date): string {
	const date = localDate(d);
	const time = localTime(d);
	return date && time ? `${date}T${time}` : "";
}

function sanitizeFileName(s: string): string {
	return s.replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, " ").trim();
}

function meetingFileStamp(d: Date): string {
	if (isNaN(d.getTime())) return "0000000000";
	const yy = String(d.getFullYear()).slice(-2);
	const mm = String(d.getMonth() + 1).padStart(2, "0");
	const dd = String(d.getDate()).padStart(2, "0");
	const hh = String(d.getHours()).padStart(2, "0");
	const min = String(d.getMinutes()).padStart(2, "0");
	return `${yy}${mm}${dd}${hh}${min}`;
}

function shortenMeetingSubject(subject: string, maxLength: number): string {
	if (subject.length <= maxLength) return subject;
	const clipped = subject.slice(0, maxLength).trimEnd();
	const lastSpace = clipped.lastIndexOf(" ");
	return (lastSpace >= Math.floor(maxLength * 0.65) ? clipped.slice(0, lastSpace) : clipped).trimEnd();
}

function neutralizeLegacyMeetingTemplate(body: string): string {
	return body.replace(
		/^(outlook date|outlook_time|outlook_subject|calendar_event|location):.*$/gm,
		"$1:",
	);
}
