// Copies the Docxodus WebAssembly runtime out of node_modules into public/,
// where Next serves it as static files. The runtime is loaded with a dynamic
// import of `<base>/_framework/dotnet.js` at view time, so it cannot be part
// of the JavaScript bundle. Runs before `dev` and `build`; the copy is
// git-ignored and skipped when it already matches the installed version.
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const frontend = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

export function copyDocxodusWasm({
    packageDir = path.join(frontend, "node_modules", "docxodus"),
    targetDir = path.join(frontend, "public", "docxodus"),
} = {}) {
    const source = path.join(packageDir, "dist", "wasm");
    if (!existsSync(source)) {
        // The Docxodus engine is opt-in; a checkout without the package must
        // still build with the default engine.
        return { copied: false, reason: "missing" };
    }
    const { version } = JSON.parse(readFileSync(path.join(packageDir, "package.json"), "utf8"));
    const stamp = path.join(targetDir, ".version");
    if (existsSync(stamp) && readFileSync(stamp, "utf8") === version) {
        return { copied: false, reason: "current", version };
    }
    rmSync(targetDir, { recursive: true, force: true });
    mkdirSync(targetDir, { recursive: true });
    cpSync(source, path.join(targetDir, "wasm"), { recursive: true });
    writeFileSync(stamp, version);
    return { copied: true, version };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
    const result = copyDocxodusWasm();
    if (result.copied) console.log(`Copied Docxodus ${result.version} WebAssembly runtime to public/docxodus/wasm`);
    else if (result.reason === "missing") console.warn("docxodus is not installed; the Docxodus DOCX engine will be unavailable");
}
