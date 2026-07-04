import { Notice, Plugin } from "obsidian";
import { DEFAULT_SETTINGS, PluginSettings } from "./types";
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
