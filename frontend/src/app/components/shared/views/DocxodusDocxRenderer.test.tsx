import { StrictMode } from "react";
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, expect, it, vi } from "vitest";
import DocxodusDocxRenderer, { docxodusWasmBasePath } from "./DocxodusDocxRenderer";
import type { DocxSurface } from "./DocxRenderer.types";

const createEditor = vi.hoisted(() => vi.fn());
vi.mock("docxodus/embed", () => ({
    createEditor,
    TrackedChangeMode: { Accept: 0, RenderInline: 1, StripDeletions: 2 },
}));

function fakeEditor(mount: HTMLElement) {
    mount.innerHTML = '<p data-anchor="a">Fees are <ins class="rev-ins">net 30</ins>.</p>';
    return {
        close: vi.fn(),
        listRevisions: () => [{ id: "r1", type: "insert", constituentIds: ["11"] }],
        revisionElements: () => Array.from(mount.querySelectorAll<HTMLElement>("ins")),
        revisionAt: () => ({ id: "r1" }),
    };
}

const bytes = () => new Uint8Array([1, 2, 3]).buffer;

beforeEach(() => {
    HTMLElement.prototype.scrollIntoView = vi.fn();
    createEditor.mockReset().mockImplementation(async (mount: HTMLElement) => fakeEditor(mount));
});
afterEach(() => {
    cleanup();
    vi.unstubAllEnvs();
});

it("opens the document read-only with tracked changes shown and hands the host a working surface", async () => {
    const onReady = vi.fn<(surface: DocxSurface) => void>();
    const source = bytes();
    render(<DocxodusDocxRenderer bytes={source} mode="edit" onReady={onReady} onError={vi.fn()} />);
    await waitFor(() => expect(onReady).toHaveBeenCalledTimes(1));

    const [mount, opened, options] = createEditor.mock.calls[0];
    expect(Array.from(opened as Uint8Array)).toEqual([1, 2, 3]);
    // The engine may transfer its buffer to a worker; the host's copy must survive.
    expect((opened as Uint8Array).buffer).not.toBe(source);
    expect(options).toMatchObject({ editable: false, trackedChanges: 1, showRevisions: true, wasmBasePath: "/docxodus/wasm/" });

    const surface = onReady.mock.calls[0][0];
    expect(surface.content.contains(mount as HTMLElement)).toBe(true);
    expect(surface.scroll.contains(surface.content)).toBe(true);
    expect(surface.activateRevision?.({ ins: "11" })).toBe(true);
    expect(surface.selectText?.("Fees are net 30")).toBe(true);
    expect(screen.getByLabelText("Document").closest("[data-docx-renderer]")).toHaveAttribute("data-docx-renderer", "docxodus");
});

it("reports a document the engine cannot open instead of leaving a blank panel", async () => {
    createEditor.mockRejectedValue(new Error("not a docx"));
    const onError = vi.fn();
    const onReady = vi.fn();
    render(<DocxodusDocxRenderer bytes={bytes()} mode="view" onReady={onReady} onError={onError} />);
    await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
    expect(onReady).not.toHaveBeenCalled();
});

it("releases the WASM session on unmount and keeps one session under StrictMode", async () => {
    const editors: ReturnType<typeof fakeEditor>[] = [];
    createEditor.mockImplementation(async (mount: HTMLElement) => {
        const editor = fakeEditor(mount);
        editors.push(editor);
        return editor;
    });
    const onReady = vi.fn();
    const view = render(<StrictMode><DocxodusDocxRenderer bytes={bytes()} mode="view" onReady={onReady} onError={vi.fn()} /></StrictMode>);
    await waitFor(() => expect(onReady).toHaveBeenCalledTimes(1));
    // StrictMode mounts twice: the abandoned session is closed, the live one stays open.
    expect(editors).toHaveLength(2);
    expect(editors.filter((editor) => editor.close.mock.calls.length === 0)).toHaveLength(1);
    view.unmount();
    expect(editors.every((editor) => editor.close.mock.calls.length === 1)).toBe(true);
});

it("does not reopen the document when the host only passes new callbacks", async () => {
    const source = bytes();
    const first = vi.fn();
    const second = vi.fn();
    const view = render(<DocxodusDocxRenderer bytes={source} mode="view" onReady={first} onError={vi.fn()} />);
    await waitFor(() => expect(first).toHaveBeenCalledTimes(1));
    view.rerender(<DocxodusDocxRenderer bytes={source} mode="view" onReady={second} onError={vi.fn()} />);
    expect(createEditor).toHaveBeenCalledTimes(1);
    view.rerender(<DocxodusDocxRenderer bytes={bytes()} mode="view" onReady={second} onError={vi.fn()} />);
    await waitFor(() => expect(second).toHaveBeenCalledTimes(1));
    expect(createEditor).toHaveBeenCalledTimes(2);
});

it("routes Ctrl/Cmd+S to the host instead of the browser's save dialog", async () => {
    const onSave = vi.fn();
    const onReady = vi.fn();
    render(<DocxodusDocxRenderer bytes={bytes()} mode="view" onSave={onSave} onReady={onReady} onError={vi.fn()} />);
    await waitFor(() => expect(onReady).toHaveBeenCalled());
    const event = new KeyboardEvent("keydown", { key: "s", ctrlKey: true, bubbles: true, cancelable: true });
    fireEvent(screen.getByLabelText("Document"), event);
    expect(onSave).toHaveBeenCalledTimes(1);
    expect(event.defaultPrevented).toBe(true);
});

it("serves the runtime from public/ by default and from a configured origin otherwise", () => {
    expect(docxodusWasmBasePath()).toBe("/docxodus/wasm/");
    vi.stubEnv("NEXT_PUBLIC_DOCXODUS_WASM_PATH", "https://cdn.example.test/docxodus/17/wasm");
    expect(docxodusWasmBasePath()).toBe("https://cdn.example.test/docxodus/17/wasm/");
});
