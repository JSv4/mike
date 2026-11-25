"use client";

import { useEffect, useRef } from "react";
import { createDocxodusSurface } from "./docxodusSurface";
import styles from "./DocxodusDocxRenderer.module.css";
import type { DocxRendererProps } from "./DocxRenderer.types";

/**
 * Where the Docxodus WebAssembly runtime is served from. The default matches
 * `scripts/copy-docxodus-wasm.mjs`, which copies it into `public/` before
 * `dev` and `build`; point this at a CDN or another origin to host it there.
 */
export function docxodusWasmBasePath(): string {
    const configured = process.env.NEXT_PUBLIC_DOCXODUS_WASM_PATH?.trim();
    const base = configured || "/docxodus/wasm/";
    return base.endsWith("/") ? base : `${base}/`;
}

/**
 * View-only DOCX renderer backed by Docxodus (MIT), a .NET OOXML engine
 * compiled to WebAssembly. The document is parsed and laid out in the browser;
 * nothing is sent to a conversion service. `mode`, `onChange` and the toolbar
 * props are ignored: the engine registry reports this engine as not editable,
 * so the host never offers Edit.
 */
export default function DocxodusDocxRenderer({ bytes, onSave, onReady, onError }: DocxRendererProps) {
    const scroll = useRef<HTMLDivElement>(null);
    const content = useRef<HTMLDivElement>(null);
    // The host passes fresh callbacks on every render; reopening the WASM
    // session for that would discard scroll position and highlights.
    const callbacks = useRef({ onReady, onError });
    useEffect(() => { callbacks.current = { onReady, onError }; });

    useEffect(() => {
        const scrollElement = scroll.current;
        const contentElement = content.current;
        if (!scrollElement || !contentElement) return;
        let cancelled = false;
        let close: (() => void) | undefined;
        // The engine owns everything under its own mount, so a remount in
        // StrictMode or for new bytes never shares nodes with the last one.
        const mount = document.createElement("div");
        contentElement.replaceChildren(mount);
        void (async () => {
            const { createEditor, TrackedChangeMode } = await import("docxodus/embed");
            const editor = await createEditor(mount, new Uint8Array(bytes.slice(0)), {
                wasmBasePath: docxodusWasmBasePath(),
                editable: false,
                // Existing tracked changes render inline as ins/del; nothing
                // is recorded because nothing is editable.
                trackedChanges: TrackedChangeMode.RenderInline,
                showRevisions: true,
                comments: true,
                fitToWidth: true,
            });
            if (cancelled) {
                editor.close();
                return;
            }
            close = () => editor.close();
            callbacks.current.onReady(createDocxodusSurface({
                content: contentElement, scroll: scrollElement, editor, bytes,
            }));
        })().catch(() => {
            if (!cancelled) callbacks.current.onError();
        });
        return () => {
            cancelled = true;
            close?.();
            mount.remove();
        };
    }, [bytes]);

    return (
        <div className={`${styles.host} relative flex h-full min-h-0 flex-col`} data-docx-renderer="docxodus"
            onKeyDownCapture={(event) => {
                if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "s") {
                    event.preventDefault();
                    event.stopPropagation();
                    void onSave?.();
                }
            }}>
            <div ref={scroll} className={`${styles.scroll} min-h-0 flex-1 overflow-auto`} tabIndex={0} aria-label="Document">
                <div ref={content} />
            </div>
        </div>
    );
}
