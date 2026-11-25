import { lazy, type ComponentType, type LazyExoticComponent } from "react";
import type { DocxRendererProps } from "./DocxRenderer.types";

/**
 * The DOCX rendering engines `DocxView` can mount. Every engine implements
 * the same `DocxRendererProps` in / `DocxSurface` out contract, so the host
 * (autosave, citations, edit cards, close guards) never names an engine.
 */
export type DocxEngineId = "eigenpal" | "docxodus";

export interface DocxEngineCapabilities {
    /** Direct editing and the formatting toolbar. */
    edit: boolean;
    /** Recording edits as tracked changes (Suggesting mode). */
    suggest: boolean;
    /** Creating, replying to and deleting comments. */
    comment: boolean;
}

export interface DocxEngine {
    id: DocxEngineId;
    label: string;
    capabilities: DocxEngineCapabilities;
    Renderer: LazyExoticComponent<ComponentType<DocxRendererProps>>;
}

export const DEFAULT_DOCX_ENGINE: DocxEngineId = "eigenpal";

// Each import stays a static literal so the bundler emits one lazy chunk per
// engine; an engine that is never selected is never downloaded.
const ENGINES: Record<DocxEngineId, DocxEngine> = {
    eigenpal: {
        id: "eigenpal",
        label: "EigenPal",
        capabilities: { edit: true, suggest: true, comment: true },
        Renderer: lazy(() => import("./EigenpalDocxRenderer")),
    },
    docxodus: {
        id: "docxodus",
        label: "Docxodus",
        // View-only for now: tracked changes and citations are shown, and the
        // host hides Edit instead of offering controls that would do nothing.
        capabilities: { edit: false, suggest: false, comment: false },
        Renderer: lazy(() => import("./DocxodusDocxRenderer")),
    },
};

export function isDocxEngineId(value: unknown): value is DocxEngineId {
    return typeof value === "string" && Object.hasOwn(ENGINES, value);
}

/** An unknown or missing name selects the default instead of breaking previews. */
export function resolveDocxEngineId(configured: string | null | undefined): DocxEngineId {
    const normalized = configured?.trim().toLowerCase();
    return isDocxEngineId(normalized) ? normalized : DEFAULT_DOCX_ENGINE;
}

export function getDocxEngine(id: DocxEngineId): DocxEngine {
    return ENGINES[id];
}

/**
 * The engine this build renders DOCX with. `NEXT_PUBLIC_DOCX_ENGINE` is read
 * as a literal member expression so Next inlines it into the client bundle.
 */
export function activeDocxEngine(): DocxEngine {
    return getDocxEngine(resolveDocxEngineId(process.env.NEXT_PUBLIC_DOCX_ENGINE));
}
