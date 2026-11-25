// @vitest-environment node
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { afterEach, beforeEach, expect, it } from "vitest";
import { copyDocxodusWasm } from "../../../scripts/copy-docxodus-wasm.mjs";

let root: string;
let packageDir: string;
let targetDir: string;

function installPackage(version: string, loader: string) {
    mkdirSync(path.join(packageDir, "dist", "wasm", "_framework"), { recursive: true });
    writeFileSync(path.join(packageDir, "package.json"), JSON.stringify({ name: "docxodus", version }));
    writeFileSync(path.join(packageDir, "dist", "wasm", "_framework", "dotnet.js"), loader);
}
const servedLoader = () => readFileSync(path.join(targetDir, "wasm", "_framework", "dotnet.js"), "utf8");

beforeEach(() => {
    root = mkdtempSync(path.join(tmpdir(), "mike-docxodus-"));
    packageDir = path.join(root, "node_modules", "docxodus");
    targetDir = path.join(root, "public", "docxodus");
});
afterEach(() => rmSync(root, { recursive: true, force: true }));

it("copies the runtime to the path the renderer loads it from", () => {
    installPackage("17.0.0", "loader 17");
    expect(copyDocxodusWasm({ packageDir, targetDir })).toEqual({ copied: true, version: "17.0.0" });
    expect(servedLoader()).toBe("loader 17");
});

it("skips the copy when the installed version is already served", () => {
    installPackage("17.0.0", "loader 17");
    copyDocxodusWasm({ packageDir, targetDir });
    writeFileSync(path.join(targetDir, "wasm", "_framework", "dotnet.js"), "untouched");
    expect(copyDocxodusWasm({ packageDir, targetDir })).toMatchObject({ copied: false, reason: "current" });
    expect(servedLoader()).toBe("untouched");
});

it("replaces a stale runtime, including files the new version no longer ships", () => {
    installPackage("16.0.0", "loader 16");
    copyDocxodusWasm({ packageDir, targetDir });
    writeFileSync(path.join(targetDir, "wasm", "_framework", "removed.wasm"), "old");
    installPackage("17.0.0", "loader 17");
    expect(copyDocxodusWasm({ packageDir, targetDir })).toEqual({ copied: true, version: "17.0.0" });
    expect(servedLoader()).toBe("loader 17");
    expect(existsSync(path.join(targetDir, "wasm", "_framework", "removed.wasm"))).toBe(false);
});

it("leaves the build alone when the optional engine is not installed", () => {
    expect(copyDocxodusWasm({ packageDir, targetDir })).toEqual({ copied: false, reason: "missing" });
    expect(existsSync(targetDir)).toBe(false);
});
