import { existsSync, realpathSync } from "node:fs";
import { delimiter, dirname, join } from "node:path";
import { execFileSync } from "node:child_process";

/**
 * LibreOffice engine locator (办公文档引擎定位).
 *
 * Resolution order:
 *   1. `INNO_LIBREOFFICE_PATH` — explicit path to the soffice binary.
 *   2. Bundled copy shipped inside the desktop installer
 *      (`INNO_BUNDLED_TOOLS_DIR` is set by electron/main.js to
 *      `process.resourcesPath`; the engine lives under `libreoffice/` there).
 *   3. Well-known system install locations per platform.
 *   4. `soffice` on PATH.
 *
 * The backend runs both inside Electron (as a child process) and standalone
 * (server/CLI/Docker), so detection must never assume an Electron runtime.
 * Docker images install LibreOffice system-wide and land on branch 4.
 */

export interface LibreOfficeStatus {
	available: boolean;
	/** Absolute path to the soffice binary when available. */
	sofficePath: string | null;
	/** How the binary was found; useful for settings UI and diagnostics. */
	source: "env" | "bundled" | "system" | "path" | null;
	/** `soffice --version` first line, e.g. "LibreOffice 25.8.2.1 ...". */
	version: string | null;
}

/** Path-only lookup result, without the (potentially slow) version probe. */
export interface LibreOfficeLocation {
	sofficePath: string;
	source: NonNullable<LibreOfficeStatus["source"]>;
}

const BUNDLED_RELATIVE: Record<string, string> = {
	darwin: join("libreoffice", "LibreOffice.app", "Contents", "MacOS", "soffice"),
	win32: join("libreoffice", "LibreOffice", "program", "soffice.exe"),
	linux: join("libreoffice", "program", "soffice"),
};

// Development staging layout produced by scripts/prepare-libreoffice.*, so
// INNO_BUNDLED_TOOLS_DIR can point straight at vendor/libreoffice/<platform>/
// without repackaging. Same tree minus the packaged `libreoffice/` prefix.
const STAGED_RELATIVE: Record<string, string> = {
	darwin: join("LibreOffice.app", "Contents", "MacOS", "soffice"),
	win32: join("LibreOffice", "program", "soffice.exe"),
	linux: join("program", "soffice"),
};

const SYSTEM_CANDIDATES: Record<string, string[]> = {
	darwin: [
		"/Applications/LibreOffice.app/Contents/MacOS/soffice",
		"/opt/homebrew/bin/soffice",
		"/usr/local/bin/soffice",
	],
	win32: [
		"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
		"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
	],
	linux: ["/usr/bin/soffice", "/usr/local/bin/soffice", "/snap/bin/libreoffice"],
};

function probeVersion(sofficePath: string): string | null {
	try {
		const out = execFileSync(sofficePath, ["--version"], {
			encoding: "utf8",
			timeout: 15_000,
			windowsHide: true,
		});
		return out.trim().split("\n")[0] || null;
	} catch {
		return null;
	}
}

function status(sofficePath: string | null, source: LibreOfficeStatus["source"]): LibreOfficeStatus {
	if (!sofficePath) return { available: false, sofficePath: null, source: null, version: null };
	return { available: true, sofficePath, source, version: probeVersion(sofficePath) };
}

function fromPath(binary: string): string | null {
	const pathEnv = process.env.PATH ?? "";
	const exts = process.platform === "win32" ? [".exe", ".bat", ""] : [""];
	for (const dir of pathEnv.split(delimiter)) {
		if (!dir) continue;
		for (const ext of exts) {
			const candidate = join(dir, `${binary}${ext}`);
			if (existsSync(candidate)) return candidate;
		}
	}
	return null;
}

export function resolveLibreOffice(): LibreOfficeStatus {
	const found = findSoffice();
	if (!found) return status(null, null);
	return { available: true, ...found, version: probeVersion(found.sofficePath) };
}

/**
 * Locate the soffice binary without probing its version. Safe to call at
 * process startup: it only does filesystem checks and a PATH scan.
 */
export function findSoffice(): LibreOfficeLocation | null {
	const explicit = process.env.INNO_LIBREOFFICE_PATH;
	if (explicit && existsSync(explicit)) return { sofficePath: explicit, source: "env" };

	const toolsDir = process.env.INNO_BUNDLED_TOOLS_DIR;
	if (toolsDir) {
		for (const rel of [BUNDLED_RELATIVE[process.platform], STAGED_RELATIVE[process.platform]]) {
			if (!rel) continue;
			const bundled = join(toolsDir, rel);
			if (existsSync(bundled)) return { sofficePath: bundled, source: "bundled" };
		}
	}

	for (const candidate of SYSTEM_CANDIDATES[process.platform] ?? []) {
		if (existsSync(candidate)) return { sofficePath: candidate, source: "system" };
	}

	const onPath = fromPath("soffice");
	if (onPath) return { sofficePath: onPath, source: "path" };

	return null;
}

/**
 * Prepend the resolved engine's directory to PATH so downstream consumers
 * that do their own lookup — notably @llamaindex/liteparse, which shells out
 * to `which soffice` — find the bundled/explicit copy too. Returns the
 * location that was applied, or null when nothing was found. Idempotent.
 */
export function prependLibreOfficeToPath(): LibreOfficeLocation | null {
	const found = findSoffice();
	if (!found) return null;
	const dir = dirname(found.sofficePath);
	const entries = (process.env.PATH ?? "").split(delimiter).filter(Boolean);
	if (!entries.includes(dir)) {
		process.env.PATH = [dir, ...entries].join(delimiter);
	}
	return found;
}

/**
 * Locate the python interpreter that can drive UNO for a given soffice
 * binary. The UNO bridge must run under LibreOffice's bundled python (it is
 * the only interpreter with a matching pyuno):
 *   - macOS:  <LibreOffice.app>/Contents/Resources/python (wrapper script
 *     that sets PYTHONHOME/PYTHONPATH and execs the framework interpreter)
 *   - win32:  <program>/python.exe next to soffice.exe
 *   - linux:  <program>/python when present (distro packages usually ship
 *     python3-uno for the system python instead — returns null there and the
 *     caller may fall back to system python3 with PYTHONPATH set)
 *
 * macOS caveat: the bundled interpreter binary is signed with a launch
 * constraint and can only be exec'd from LibreOffice's own processes until
 * re-signed — scripts/prepare-libreoffice.sh re-signs it ad hoc at staging
 * time. A user's own /Applications install is NOT re-signed, so the bridge
 * may fail to spawn against `source: "system"` engines on macOS.
 */
export function findOfficePython(sofficePath: string): string | null {
	const realSoffice = (() => {
		try {
			return realpathSync(sofficePath);
		} catch {
			return sofficePath;
		}
	})();
	const dir = dirname(realSoffice);
	const candidates: string[] =
		process.platform === "darwin"
			// MacOS/soffice -> Contents/Resources/python
			? [join(dir, "..", "Resources", "python")]
			: process.platform === "win32"
				? [join(dir, "python.exe")]
				: [join(dir, "python"), join(dir, "python3")];
	for (const candidate of candidates) {
		if (existsSync(candidate)) return candidate;
	}
	return null;
}
