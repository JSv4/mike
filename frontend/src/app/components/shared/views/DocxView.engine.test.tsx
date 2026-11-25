import { useEffect, useRef } from "react";
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, expect, it, vi } from "vitest";
import { authenticatedFetch } from "@/app/lib/authEvents";
import type { DocxRendererProps } from "./DocxRenderer.types";
import { DocxView } from "./DocxView";

const replaceVersion = vi.hoisted(() => vi.fn());
const exportDocx = vi.hoisted(() => vi.fn());

vi.mock("@/app/lib/authEvents", () => ({ authenticatedFetch: vi.fn() }));
vi.mock("@/app/lib/mikeApi", () => ({
    getDocumentFileUrl: (id: string) => "/api/document/" + id,
    replaceDocumentVersionFile: replaceVersion,
    listDocumentVersions: vi.fn().mockResolvedValue({ current_version_id: "v1", versions: [] }),
    MikeApiError: class extends Error {},
}));

/** Stands in for either engine: reports the props the host handed it. */
function engineStub(name: string) {
    return function Stub({ bytes, mode, toolbarVisible, onChange, onReady }: DocxRendererProps) {
        const scroll = useRef<HTMLDivElement>(null);
        const content = useRef<HTMLDivElement>(null);
        useEffect(() => {
            if (scroll.current && content.current) onReady({ scroll: scroll.current, content: content.current, exportDocx });
        }, [bytes, onReady]);
        return (
            <div ref={scroll} data-testid="engine" data-engine={name} data-mode={mode} data-toolbar={String(toolbarVisible)}>
                <button onClick={onChange}>Engine change event</button>
                <div ref={content}>{name} preview</div>
            </div>
        );
    };
}
vi.mock("./EigenpalDocxRenderer", () => ({ default: engineStub("eigenpal") }));
vi.mock("./DocxodusDocxRenderer", () => ({ default: engineStub("docxodus") }));

beforeEach(() => {
    replaceVersion.mockReset().mockResolvedValue({ id: "v1" });
    exportDocx.mockReset().mockResolvedValue(new Uint8Array([42]).buffer);
    vi.mocked(authenticatedFetch).mockReset().mockImplementation(async () =>
        new Response(new Uint8Array([1, 2, 3]), { status: 200 }));
});
afterEach(() => {
    cleanup();
    vi.unstubAllEnvs();
    vi.useRealTimers();
});

it("mounts the existing engine with editing when nothing is configured", async () => {
    render(<DocxView documentId="doc-1" versionId="v1" defaultMode="edit" canEdit toolbarVisible cacheBytes={false} />);
    const engine = await screen.findByTestId("engine");
    expect(engine).toHaveAttribute("data-engine", "eigenpal");
    expect(engine).toHaveAttribute("data-mode", "edit");
    expect(engine).toHaveAttribute("data-toolbar", "true");
});

it("mounts the build's configured engine", async () => {
    vi.stubEnv("NEXT_PUBLIC_DOCX_ENGINE", "docxodus");
    render(<DocxView documentId="doc-1" versionId="v1" cacheBytes={false} />);
    expect(await screen.findByTestId("engine")).toHaveAttribute("data-engine", "docxodus");
    expect(screen.getByText("docxodus preview")).toBeInTheDocument();
});

it("keeps a view-only engine read-only even for a user who may edit the document", async () => {
    const onSaveStateChange = vi.fn();
    render(<DocxView engine="docxodus" documentId="doc-1" versionId="v1" defaultMode="edit" canEdit toolbarVisible
        cacheBytes={false} onSaveStateChange={onSaveStateChange} />);
    const engine = await screen.findByTestId("engine");
    expect(engine).toHaveAttribute("data-mode", "view");
    expect(engine).toHaveAttribute("data-toolbar", "false");

    // A stray change event from a view-only engine must never reach autosave.
    vi.useFakeTimers();
    fireEvent.click(screen.getByText("Engine change event"));
    await vi.advanceTimersByTimeAsync(5000);
    vi.useRealTimers();
    expect(replaceVersion).not.toHaveBeenCalled();
    expect(exportDocx).not.toHaveBeenCalled();
    await waitFor(() => expect(onSaveStateChange).toHaveBeenCalled());
    expect(onSaveStateChange.mock.calls.every(([, state]) => !state || state.dirty === false)).toBe(true);
});

it("still offers Download from a view-only engine", async () => {
    let download: (() => Promise<void>) | null = null;
    const createObjectURL = vi.fn(() => "blob:mike");
    vi.stubGlobal("URL", Object.assign(URL, { createObjectURL, revokeObjectURL: vi.fn() }));
    render(<DocxView engine="docxodus" documentId="doc-1" versionId="v1" filename="NDA" cacheBytes={false}
        onDownloadReady={(next) => { download = next; }} />);
    await screen.findByTestId("engine");
    await waitFor(() => expect(download).not.toBeNull());
    await download!();
    expect(exportDocx).toHaveBeenCalledTimes(1);
    expect(createObjectURL).toHaveBeenCalledTimes(1);
});
