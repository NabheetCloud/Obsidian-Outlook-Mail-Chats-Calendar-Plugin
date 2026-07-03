// Tiny logging shim so every module logs under a consistent prefix and the
// verbose channel can be toggled from settings.

let debug = false;

export function setDebug(on: boolean): void {
	debug = on;
}

/**
 * Thrown to unwind a sync the user asked to stop. Defined here (a leaf module
 * with no project imports) so both graph.ts and sync.ts can throw/catch it
 * without a circular dependency.
 */
export class SyncCancelledError extends Error {
	constructor() {
		super("Sync stopped by user.");
		this.name = "SyncCancelledError";
	}
}

export function log(...args: unknown[]): void {
	if (debug) console.log("[obs-mail]", ...args);
}

/** Always-on baseline progress — prints regardless of the debug toggle. */
export function info(...args: unknown[]): void {
	console.log("[obs-mail]", ...args);
}

/**
 * Rejects if `p` doesn't settle within `ms`. The underlying request can't be
 * cancelled, but this guarantees the await resolves so callers can surface the
 * error and clear their running flag instead of hanging forever.
 */
export function withTimeout<T>(p: Promise<T>, ms: number, label: string): Promise<T> {
	return new Promise<T>((resolve, reject) => {
		const timer = setTimeout(
			() => reject(new Error(`${label} timed out after ${Math.round(ms / 1000)}s`)),
			ms,
		);
		p.then(
			(v) => {
				clearTimeout(timer);
				resolve(v);
			},
			(e) => {
				clearTimeout(timer);
				reject(e);
			},
		);
	});
}

export function warn(...args: unknown[]): void {
	console.warn("[obs-mail]", ...args);
}

export function logError(...args: unknown[]): void {
	console.error("[obs-mail]", ...args);
}
