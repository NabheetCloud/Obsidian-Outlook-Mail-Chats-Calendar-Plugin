import { ItemView, WorkspaceLeaf } from "obsidian";
import type OutlookMailboxPlugin from "../main";
import { UpcomingEvent } from "../types";

export const VIEW_TYPE_UPCOMING = "outlook-upcoming-view";

export class UpcomingView extends ItemView {
	constructor(leaf: WorkspaceLeaf, private plugin: OutlookMailboxPlugin) {
		super(leaf);
	}

	getViewType(): string {
		return VIEW_TYPE_UPCOMING;
	}
	getDisplayText(): string {
		return "Upcoming meetings";
	}
	getIcon(): string {
		return "calendar-clock";
	}

	async onOpen(): Promise<void> {
		this.render();
	}
	async onClose(): Promise<void> {}

	/** Re-renders the list from the persisted cache. Called after each sync. */
	render(): void {
		const root = this.contentEl;
		root.empty();
		root.addClass("outlook-upcoming");

		const header = root.createDiv({ cls: "outlook-upcoming-header" });
		header.createEl("h4", { text: "Upcoming" });
		const refresh = header.createEl("button", { text: "Sync" });
		refresh.onclick = () => this.plugin.runSync();

		const events = [...this.plugin.settings.upcomingCache].sort((a, b) =>
			a.startIso < b.startIso ? -1 : 1,
		);

		if (!this.plugin.settings.syncCalendar) {
			root.createEl("p", {
				text: "Calendar sync is off. Enable it in plugin settings.",
				cls: "outlook-upcoming-empty",
			});
			return;
		}
		if (events.length === 0) {
			root.createEl("p", {
				text: "No upcoming meetings. Run a sync.",
				cls: "outlook-upcoming-empty",
			});
			return;
		}

		let lastDay = "";
		for (const ev of events) {
			const dayLabel = dayHeading(ev.startIso);
			if (dayLabel !== lastDay) {
				root.createEl("div", { text: dayLabel, cls: "outlook-upcoming-day" });
				lastDay = dayLabel;
			}
			this.renderRow(root, ev);
		}
	}

	private renderRow(root: HTMLElement, ev: UpcomingEvent): void {
		const row = root.createDiv({ cls: "outlook-upcoming-row" });

		const time = row.createDiv({ cls: "outlook-upcoming-time" });
		time.setText(ev.isAllDay ? "all day" : fmtTime(ev.startIso));

		const body = row.createDiv({ cls: "outlook-upcoming-body" });
		const title = body.createDiv({ cls: "outlook-upcoming-title", text: ev.subject });
		if (ev.location) {
			body.createDiv({ cls: "outlook-upcoming-loc", text: ev.location });
		}

		// Click title → open the meeting note (or the Outlook web link as fallback).
		title.onclick = () => {
			if (ev.notePath) {
				this.app.workspace.openLinkText(ev.notePath, "", false);
			} else if (ev.webLink) {
				window.open(ev.webLink, "_blank");
			}
		};

		if (ev.onlineUrl) {
			const join = row.createEl("a", { cls: "outlook-upcoming-join", text: "Join" });
			join.onclick = (e) => {
				e.preventDefault();
				window.open(ev.onlineUrl, "_blank");
			};
		}
	}
}

// ---- date helpers (local time) ----

function startOfDay(d: Date): number {
	return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

function dayHeading(iso: string): string {
	const d = new Date(iso);
	if (isNaN(d.getTime())) return iso.slice(0, 10);
	const today = startOfDay(new Date());
	const day = startOfDay(d);
	const diff = Math.round((day - today) / 86_400_000);
	if (diff === 0) return "Today";
	if (diff === 1) return "Tomorrow";
	return d.toLocaleDateString([], { weekday: "long", month: "short", day: "numeric" });
}

function fmtTime(iso: string): string {
	const d = new Date(iso);
	if (isNaN(d.getTime())) return "";
	return d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}
