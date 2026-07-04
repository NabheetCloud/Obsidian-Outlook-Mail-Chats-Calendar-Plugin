import { App, PluginSettingTab, Setting, Notice } from "obsidian";
import type OutlookMailboxPlugin from "./main";
import { FolderMapping } from "./types";

export class OutlookMailboxSettingTab extends PluginSettingTab {
	constructor(app: App, private plugin: OutlookMailboxPlugin) {
		super(app, plugin);
	}

	display(): void {
		const { containerEl } = this;
		containerEl.empty();
		const s = this.plugin.settings;

		// --- Account / auth ---
		new Setting(containerEl).setName("Account").setHeading();

		const status = this.plugin.graph.isAuthenticated
			? `Connected${this.plugin.connectedAs ? ` as ${this.plugin.connectedAs}` : ""}`
			: "Not connected";
		new Setting(containerEl)
			.setName("Connection status")
			.setDesc(status)
			.addButton((b) =>
				b
					.setButtonText(this.plugin.graph.isAuthenticated ? "Reconnect" : "Connect")
					.setCta()
					.onClick(async () => {
						b.setDisabled(true);
						try {
							await this.plugin.connect();
							new Notice("Connected to Microsoft 365.");
						} catch (e) {
							new Notice(`Connection failed: ${(e as Error).message}`);
						} finally {
							b.setDisabled(false);
							this.display();
						}
					}),
			)
			.addExtraButton((b) =>
				b
					.setIcon("log-out")
					.setTooltip("Disconnect")
					.onClick(async () => {
						await this.plugin.graph.logout();
						new Notice("Disconnected.");
						this.display();
					}),
			);

		new Setting(containerEl)
			.setName("Application (client) ID")
			.setDesc("From your Azure AD app registration. See README for one-time setup.")
			.addText((t) =>
				t
					.setPlaceholder("00000000-0000-0000-0000-000000000000")
					.setValue(s.clientId)
					.onChange(async (v) => {
						s.clientId = v.trim();
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Tenant")
			.setDesc('"common", "organizations", or your tenant GUID. Use "organizations" for work/school.')
			.addText((t) =>
				t
					.setPlaceholder("organizations")
					.setValue(s.tenant)
					.onChange(async (v) => {
						s.tenant = v.trim() || "common";
						await this.plugin.saveSettings();
					}),
			);

		// --- Vault layout ---
		new Setting(containerEl).setName("Vault layout").setHeading();

		new Setting(containerEl)
			.setName("Target folder")
			.setDesc("Root vault folder that all mail notes live under.")
			.addText((t) =>
				t
					.setPlaceholder("10-Mailbox")
					.setValue(s.targetFolder)
					.onChange(async (v) => {
						s.targetFolder = v.trim() || "10-Mailbox";
						await this.plugin.saveSettings();
					}),
			);

		// --- Folder mappings ---
		new Setting(containerEl).setName("Folders to sync").setHeading();
		containerEl.createEl("p", {
			text: "Map Outlook folders to vault subfolders. Use a well-known name (inbox, archive, sentitems) or a display-name path like Projects/ClientX.",
			cls: "setting-item-description",
		});

		s.folders.forEach((mapping, idx) => {
			this.renderFolderRow(containerEl, mapping, idx);
		});

		new Setting(containerEl).addButton((b) =>
			b
				.setButtonText("Add folder")
				.setIcon("plus")
				.onClick(async () => {
					s.folders.push({ outlookPath: "", vaultSubfolder: "", enabled: true });
					await this.plugin.saveSettings();
					this.display();
				}),
		);

		// --- Sync behaviour ---
		new Setting(containerEl).setName("Sync").setHeading();

		new Setting(containerEl)
			.setName("Sync mail newer than")
			.setDesc(
				"Only import mail received on/after this date (YYYY-MM-DD). Leave empty for all mail. Older messages are skipped on every sync.",
			)
			.addText((t) =>
				t
					.setPlaceholder("2025-01-01")
					.setValue(s.syncSince)
					.onChange(async (v) => {
						const val = v.trim();
						if (val && !/^\d{4}-\d{2}-\d{2}$/.test(val)) {
							t.inputEl.addClass("outlook-mailbox-invalid");
							return; // don't persist malformed input
						}
						t.inputEl.removeClass("outlook-mailbox-invalid");
						s.syncSince = val;
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Auto-sync interval (minutes)")
			.setDesc("0 disables the timer. Sync manually via the ribbon icon or command.")
			.addText((t) =>
				t.setValue(String(s.syncIntervalMinutes)).onChange(async (v) => {
					const n = Math.max(0, Math.floor(Number(v) || 0));
					s.syncIntervalMinutes = n;
					await this.plugin.saveSettings();
					this.plugin.rescheduleTimer();
				}),
			);

		new Setting(containerEl)
			.setName("Sync on startup")
			.addToggle((t) =>
				t.setValue(s.syncOnStartup).onChange(async (v) => {
					s.syncOnStartup = v;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl)
			.setName("Download attachments")
			.setDesc("Save file attachments under each subfolder's _attachments directory.")
			.addToggle((t) =>
				t.setValue(s.downloadAttachments).onChange(async (v) => {
					s.downloadAttachments = v;
					await this.plugin.saveSettings();
					this.display();
				}),
			);

		if (s.downloadAttachments) {
			new Setting(containerEl)
				.setName("Max attachment size (MB)")
				.setDesc("0 = no limit.")
				.addText((t) =>
					t.setValue(String(s.maxAttachmentMB)).onChange(async (v) => {
						s.maxAttachmentMB = Math.max(0, Math.floor(Number(v) || 0));
						await this.plugin.saveSettings();
					}),
				);
		}

		new Setting(containerEl)
			.setName("Debug logging")
			.setDesc("Verbose console output under the [obs-mail] prefix.")
			.addToggle((t) =>
				t.setValue(s.debugLogging).onChange(async (v) => {
					s.debugLogging = v;
					await this.plugin.saveSettings();
					this.plugin.applyDebug();
				}),
			);

		// --- Calendar ---
		new Setting(containerEl).setName("Calendar").setHeading();
		containerEl.createEl("p", {
			text: "Requires the Calendars.Read permission. Add it to your Azure app registration and click Reconnect above after enabling.",
			cls: "setting-item-description",
		});

		new Setting(containerEl)
			.setName("Sync calendar")
			.setDesc("Fetch upcoming events, write meeting notes, and show the Upcoming sidebar.")
			.addToggle((t) =>
				t.setValue(s.syncCalendar).onChange(async (v) => {
					s.syncCalendar = v;
					await this.plugin.saveSettings();
					this.display();
				}),
			);

		if (s.syncCalendar) {
			new Setting(containerEl)
				.setName("Days ahead")
				.setDesc("Rolling window of upcoming days to keep.")
				.addText((t) =>
					t.setValue(String(s.calendarDaysAhead)).onChange(async (v) => {
						s.calendarDaysAhead = Math.max(1, Math.floor(Number(v) || 14));
						await this.plugin.saveSettings();
					}),
				);

			new Setting(containerEl)
				.setName("Calendar subfolder")
				.setDesc("Vault subfolder (under the target folder) for meeting notes.")
				.addText((t) =>
					t
						.setPlaceholder("Calendar")
						.setValue(s.calendarSubfolder)
						.onChange(async (v) => {
							s.calendarSubfolder = v.trim() || "Calendar";
							await this.plugin.saveSettings();
						}),
				);

			new Setting(containerEl)
				.setName("Link related emails")
				.setDesc("Cross-link meeting notes to email notes sharing attendees or subject.")
				.addToggle((t) =>
					t.setValue(s.linkRelatedEmails).onChange(async (v) => {
						s.linkRelatedEmails = v;
						await this.plugin.saveSettings();
					}),
				);

			new Setting(containerEl)
				.setName("Upcoming panel")
				.addButton((b) =>
					b
						.setButtonText("Open sidebar")
						.onClick(() => this.plugin.activateUpcomingView()),
				);
		}

		// --- Teams ---
		new Setting(containerEl).setName("Teams").setHeading();
		containerEl.createEl("p", {
			text:
				"Sync Teams messages as one transcript note per conversation. Chats use Chat.Read (user consent). " +
				"Channels need ChannelMessage.Read.All + Team/Channel.ReadBasic.All, which usually requires tenant admin consent. " +
				"After changing these toggles, click Reconnect above to re-consent to the new scopes.",
			cls: "setting-item-description",
		});

		new Setting(containerEl)
			.setName("Sync Teams")
			.setDesc("Fetch Teams messages into per-conversation transcript notes.")
			.addToggle((t) =>
				t.setValue(s.syncTeams).onChange(async (v) => {
					s.syncTeams = v;
					await this.plugin.saveSettings();
					this.display();
				}),
			);

		if (s.syncTeams) {
			new Setting(containerEl)
				.setName("Sync chats")
				.setDesc("1:1 and group chats (Chat.Read).")
				.addToggle((t) =>
					t.setValue(s.teamsChats).onChange(async (v) => {
						s.teamsChats = v;
						await this.plugin.saveSettings();
					}),
				);

			new Setting(containerEl)
				.setName("Sync channels")
				.setDesc("Team channel posts + replies. Needs admin consent for ChannelMessage.Read.All.")
				.addToggle((t) =>
					t.setValue(s.teamsChannels).onChange(async (v) => {
						s.teamsChannels = v;
						await this.plugin.saveSettings();
					}),
				);

			new Setting(containerEl)
				.setName("Teams subfolder")
				.setDesc("Vault subfolder (under the target folder) for conversation transcripts.")
				.addText((t) =>
					t
						.setPlaceholder("Teams")
						.setValue(s.teamsSubfolder)
						.onChange(async (v) => {
							s.teamsSubfolder = v.trim() || "Teams";
							await this.plugin.saveSettings();
						}),
				);

			new Setting(containerEl)
				.setName("Sync messages from the last N days")
				.setDesc("0 = all history (still bounded by the mail 'sync newer than' date if set).")
				.addText((t) =>
					t.setValue(String(s.teamsDaysBack)).onChange(async (v) => {
						s.teamsDaysBack = Math.max(0, Math.floor(Number(v) || 0));
						await this.plugin.saveSettings();
					}),
				);
		}

		// --- Maintenance ---
		new Setting(containerEl).setName("Maintenance").setHeading();
		new Setting(containerEl)
			.setName("Sync now")
			.setDesc("Stop halts the running sync at the next message/page; already-written notes are kept.")
			.addButton((b) =>
				b
					.setButtonText("Sync now")
					.setCta()
					.setDisabled(this.plugin.sync.isRunning)
					.onClick(() => {
						// Fire without awaiting so the panel can re-render immediately
						// (enabling Stop); re-render again when the run settles.
						void this.plugin.runSync().finally(() => this.display());
						this.display();
					}),
			)
			.addButton((b) =>
				b
					.setButtonText("Stop")
					.setWarning()
					.setDisabled(!this.plugin.sync.isRunning)
					.onClick(() => {
						this.plugin.stopSync();
						this.display();
					}),
			);

		new Setting(containerEl)
			.setName("Reset sync state")
			.setDesc(
				"Clears delta tokens and the thread index so the next sync re-enumerates all messages. Existing notes are kept.",
			)
			.addButton((b) =>
				b
					.setButtonText("Reset")
					.setWarning()
					.onClick(async () => {
						this.plugin.sync.resetState();
						await this.plugin.saveSettings();
						new Notice("Sync state reset. Next sync will do a full enumeration.");
					}),
			);

		if (s.lastSync) {
			containerEl.createEl("p", {
				text: `Last sync: ${new Date(s.lastSync).toLocaleString()}`,
				cls: "setting-item-description",
			});
		}
	}

	private renderFolderRow(containerEl: HTMLElement, mapping: FolderMapping, idx: number): void {
		const s = this.plugin.settings;
		const row = new Setting(containerEl)
			.addText((t) =>
				t
					.setPlaceholder("Outlook path (e.g. inbox)")
					.setValue(mapping.outlookPath)
					.onChange(async (v) => {
						mapping.outlookPath = v.trim();
						await this.plugin.saveSettings();
					}),
			)
			.addText((t) =>
				t
					.setPlaceholder("Vault subfolder (e.g. Inbox)")
					.setValue(mapping.vaultSubfolder)
					.onChange(async (v) => {
						mapping.vaultSubfolder = v.trim();
						await this.plugin.saveSettings();
					}),
			)
			.addToggle((t) =>
				t
					.setTooltip("Enabled")
					.setValue(mapping.enabled)
					.onChange(async (v) => {
						mapping.enabled = v;
						await this.plugin.saveSettings();
					}),
			)
			.addExtraButton((b) =>
				b
					.setIcon("trash")
					.setTooltip("Remove")
					.onClick(async () => {
						s.folders.splice(idx, 1);
						await this.plugin.saveSettings();
						this.display();
					}),
			);
		row.infoEl.remove();
	}
}
