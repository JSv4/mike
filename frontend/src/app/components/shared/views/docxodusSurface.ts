import { clearDocxQuoteHighlights, highlightDocxQuote } from "./highlightDocxQuote";
import type { DocxRevisionIds, DocxSurface } from "./DocxRenderer.types";

export const DOCXODUS_ACTIVE_REVISION_CLASS = "docxodus-revision-active";

/** The slice of Docxodus's `RevisionListEntry` this adapter reads. */
export interface DocxodusRevision {
    id: string;
    type: string;
    /** The OOXML `w:id` values of the markup that makes up this revision. */
    constituentIds: string[];
}

/** The slice of Docxodus's `DocxEditor` this adapter reads. */
export interface DocxodusRevisionSource {
    listRevisions(): DocxodusRevision[];
    revisionElements(): HTMLElement[];
    revisionAt(element: HTMLElement): { id: string } | null;
}

function reveal(element: HTMLElement) {
    // jsdom and very old browsers have no scrollIntoView.
    element.scrollIntoView?.({ block: "center" });
}

/**
 * Mike's host-facing surface over a mounted, read-only Docxodus document.
 *
 * Tracked changes resolve through the engine's revision registry by OOXML
 * `w:id`, so identical wording elsewhere cannot capture the highlight. The
 * read-only editor exposes no searchable blocks (`find()` only walks editable
 * ones), so citations use Mike's tolerant DOM quote matcher instead.
 */
export function createDocxodusSurface(params: {
    content: HTMLElement;
    scroll: HTMLElement;
    editor: DocxodusRevisionSource;
    /** View-only: Download returns exactly the bytes that were opened. */
    bytes: ArrayBuffer;
}): DocxSurface {
    const { content, scroll, editor, bytes } = params;
    const clearRevisionHighlight = () => {
        content.querySelectorAll(`.${DOCXODUS_ACTIVE_REVISION_CLASS}`)
            .forEach((element) => element.classList.remove(DOCXODUS_ACTIVE_REVISION_CLASS));
    };
    const revisionIdsFor = ({ ins, del }: DocxRevisionIds): Set<string> => {
        const matched = new Set<string>();
        for (const revision of editor.listRevisions()) {
            // An insertion and a deletion may legally share a numeric w:id, so
            // each half only matches a revision of its own kind.
            const isDeletion = /delet|move_?from/i.test(revision.type);
            const wanted = isDeletion ? del : ins;
            if (wanted != null && revision.constituentIds.includes(wanted)) matched.add(revision.id);
        }
        return matched;
    };
    return {
        content,
        scroll,
        exportDocx: async () => bytes.slice(0),
        activateRevision: (ids) => {
            clearRevisionHighlight();
            const revisionIds = revisionIdsFor(ids);
            if (!revisionIds.size) return false;
            const elements = editor.revisionElements().filter((element) => {
                const revision = editor.revisionAt(element);
                return !!revision && revisionIds.has(revision.id);
            });
            if (!elements.length) return false;
            clearDocxQuoteHighlights(content);
            elements.forEach((element) => element.classList.add(DOCXODUS_ACTIVE_REVISION_CLASS));
            reveal(elements[0]);
            return true;
        },
        clearRevisionHighlight,
        selectText: (text) => {
            const first = highlightDocxQuote(content, text);
            if (!first) return false;
            reveal(first);
            return true;
        },
        clearTextSelection: () => clearDocxQuoteHighlights(content),
    };
}
