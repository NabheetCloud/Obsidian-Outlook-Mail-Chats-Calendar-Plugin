// Copies build artifacts into an Obsidian vault's plugin directory.
// Usage: VAULT="/path/to/vault" node scripts/install.mjs
import { cpSync, existsSync, mkdirSync } from "fs";
import { join } from "path";
import process from "process";

const vault = process.env.VAULT;
if (!vault) {
	console.error('Set VAULT to your vault path, e.g. VAULT="/path/to/vault" npm run install:vault');
	process.exit(1);
}

const pluginDir = join(vault, ".obsidian", "plugins", "outlook-teams-calendar");
mkdirSync(pluginDir, { recursive: true });

for (const f of ["main.js", "manifest.json", "styles.css"]) {
	if (!existsSync(f)) {
		console.error(`Missing ${f}. Run "npm run build" first.`);
		process.exit(1);
	}
	cpSync(f, join(pluginDir, f));
}
console.log(`Installed to ${pluginDir}`);
console.log("Enable it in Obsidian → Settings → Community plugins.");
