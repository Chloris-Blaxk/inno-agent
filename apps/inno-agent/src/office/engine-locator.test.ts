import { chmodSync, mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const ORIGINAL_ENV = { ...process.env };

let workDir: string;

async function freshLocator() {
	vi.resetModules();
	return import("./engine-locator.js");
}

beforeEach(() => {
	workDir = mkdtempSync(join(tmpdir(), "engine-locator-"));
	delete process.env.INNO_LIBREOFFICE_PATH;
	delete process.env.INNO_BUNDLED_TOOLS_DIR;
});

afterEach(() => {
	process.env = { ...ORIGINAL_ENV };
	rmSync(workDir, { recursive: true, force: true });
});

describe("resolveLibreOffice", () => {
	it("prefers INNO_LIBREOFFICE_PATH when the binary exists", async () => {
		const fake = join(workDir, "soffice");
		writeFileSync(fake, "#!/bin/sh\necho 'LibreOffice 25.8.0.0 test'\n");
		chmodSync(fake, 0o755);
		process.env.INNO_LIBREOFFICE_PATH = fake;
		const { resolveLibreOffice } = await freshLocator();
		const result = resolveLibreOffice();
		expect(result.available).toBe(true);
		expect(result.sofficePath).toBe(fake);
		expect(result.source).toBe("env");
		expect(result.version).toContain("LibreOffice");
	});

	it("ignores an env override pointing at a missing binary", async () => {
		process.env.INNO_LIBREOFFICE_PATH = join(workDir, "does-not-exist");
		const { resolveLibreOffice } = await freshLocator();
		const result = resolveLibreOffice();
		expect(result.sofficePath).not.toBe(join(workDir, "does-not-exist"));
		expect(result.source).not.toBe("env");
	});

	it("resolves the bundled copy under INNO_BUNDLED_TOOLS_DIR", async () => {
		const bundled = join(workDir, "libreoffice", "LibreOffice.app", "Contents", "MacOS");
		mkdirSync(bundled, { recursive: true });
		writeFileSync(join(bundled, "soffice"), "#!/bin/sh\necho 'LibreOffice 25.8.0.0 bundled'\n");
		chmodSync(join(bundled, "soffice"), 0o755);
		process.env.INNO_BUNDLED_TOOLS_DIR = workDir;
		const { resolveLibreOffice } = await freshLocator();
		const result = resolveLibreOffice();
		if (process.platform === "darwin") {
			expect(result.available).toBe(true);
			expect(result.source).toBe("bundled");
			expect(result.sofficePath).toBe(join(bundled, "soffice"));
		} else {
			// The darwin layout must not match on other platforms.
			expect(result.source).not.toBe("bundled");
		}
	});

	it("prepends the bundled engine directory to PATH exactly once", async () => {
		const bundled = join(workDir, "libreoffice", "LibreOffice.app", "Contents", "MacOS");
		mkdirSync(bundled, { recursive: true });
		const soffice = join(bundled, "soffice");
		writeFileSync(soffice, "#!/bin/sh\n");
		chmodSync(soffice, 0o755);
		process.env.INNO_BUNDLED_TOOLS_DIR = workDir;
		const { prependLibreOfficeToPath } = await freshLocator();
		if (process.platform !== "darwin") return; // darwin layout only matches on macOS

		const first = prependLibreOfficeToPath();
		expect(first?.sofficePath).toBe(soffice);
		expect(first?.source).toBe("bundled");
		const pathAfterFirst = process.env.PATH ?? "";
		expect(pathAfterFirst.split(":")).toContain(bundled);

		const before = process.env.PATH;
		prependLibreOfficeToPath();
		expect(process.env.PATH).toBe(before);
	});

	it("returns null from prependLibreOfficeToPath when no engine exists", async () => {
		// Point every discovery channel at empty locations.
		process.env.INNO_BUNDLED_TOOLS_DIR = join(workDir, "empty");
		process.env.PATH = join(workDir, "no-bin");
		const { prependLibreOfficeToPath } = await freshLocator();
		if (process.platform === "darwin") {
			// The system candidate /Applications/LibreOffice.app may genuinely
			// exist on a developer machine; skip rather than fail spuriously.
			const { existsSync } = await import("node:fs");
			if (existsSync("/Applications/LibreOffice.app/Contents/MacOS/soffice")) return;
		}
		expect(prependLibreOfficeToPath()).toBeNull();
	});
});
