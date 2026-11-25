import { beforeEach, describe, expect, it, vi } from "vitest";
import { createDocxodusSurface, DOCXODUS_ACTIVE_REVISION_CLASS, type DocxodusRevision } from "./docxodusSurface";
import { QUOTE_HIGHLIGHT_CLASS } from "./highlightDocxQuote";

/** The markup Docxodus renders for tracked changes, and its revision registry. */
function mountDocument(revisions: (DocxodusRevision & { selector: string })[], html: string) {
    const scroll = document.createElement("div");
    const content = document.createElement("div");
    content.innerHTML = html;
    scroll.append(content);
    document.body.replaceChildren(scroll);
    const editor = {
        listRevisions: () => revisions,
        revisionElements: () => Array.from(content.querySelectorAll<HTMLElement>("ins, del")),
        revisionAt: (element: HTMLElement) => revisions.find((revision) => element.matches(revision.selector)) ?? null,
    };
    const bytes = new Uint8Array([80, 75, 3, 4]).buffer;
    return { content, scroll, bytes, surface: createDocxodusSurface({ content, scroll, editor, bytes }) };
}

const NOTICE = `
    <p data-anchor="a">Payment is due within thirty days.</p>
    <p data-anchor="b">The notice period is <del id="d7" class="rev-del">thirty</del><ins id="i8" class="rev-ins">sixty</ins> days.</p>
    <p data-anchor="c">The later notice period is <del id="d27" class="rev-del">thirty</del><ins id="i28" class="rev-ins">sixty</ins> days.</p>`;
const REVISIONS = [
    { id: "rev-d7", type: "delete", constituentIds: ["7"], selector: "#d7" },
    { id: "rev-i8", type: "insert", constituentIds: ["8"], selector: "#i8" },
    { id: "rev-d27", type: "delete", constituentIds: ["27"], selector: "#d27" },
    { id: "rev-i28", type: "insert", constituentIds: ["28"], selector: "#i28" },
];
const active = (root: HTMLElement) =>
    Array.from(root.querySelectorAll(`.${DOCXODUS_ACTIVE_REVISION_CLASS}`)).map((element) => element.id);

beforeEach(() => {
    HTMLElement.prototype.scrollIntoView = vi.fn();
});

describe("docxodus surface", () => {
    it("activates a tracked change by w:id, not by its repeated wording", () => {
        const { content, surface } = mountDocument(REVISIONS, NOTICE);
        expect(surface.activateRevision?.({ ins: "28", del: "27" })).toBe(true);
        expect(active(content)).toEqual(["d27", "i28"]);
        expect(HTMLElement.prototype.scrollIntoView).toHaveBeenCalledTimes(1);
        expect(vi.mocked(HTMLElement.prototype.scrollIntoView).mock.contexts[0]).toBe(content.querySelector("#d27"));
    });

    it("moves the highlight when another edit is activated and clears it on dismiss", () => {
        const { content, surface } = mountDocument(REVISIONS, NOTICE);
        surface.activateRevision?.({ ins: "28", del: "27" });
        surface.activateRevision?.({ ins: "8", del: null });
        expect(active(content)).toEqual(["i8"]);
        surface.clearRevisionHighlight?.();
        expect(active(content)).toEqual([]);
    });

    it("keeps an insertion and a deletion that share a numeric w:id apart", () => {
        const { content, surface } = mountDocument([
            { id: "rev-del", type: "delete", constituentIds: ["5"], selector: "#d7" },
            { id: "rev-ins", type: "insert", constituentIds: ["5"], selector: "#i8" },
        ], NOTICE);
        expect(surface.activateRevision?.({ ins: "5" })).toBe(true);
        expect(active(content)).toEqual(["i8"]);
    });

    it("reports an unknown or unrendered revision so the host can fall back to text", () => {
        const { content, surface } = mountDocument(REVISIONS, NOTICE);
        expect(surface.activateRevision?.({ ins: "404" })).toBe(false);
        expect(surface.activateRevision?.({})).toBe(false);
        const hidden = mountDocument([{ id: "rev-x", type: "insert", constituentIds: ["9"], selector: "#missing" }], NOTICE);
        expect(hidden.surface.activateRevision?.({ ins: "9" })).toBe(false);
        expect(active(content)).toEqual([]);
    });

    it("paints a citation across runs without touching the document text", () => {
        const { content, surface } = mountDocument(REVISIONS, NOTICE);
        const before = content.textContent;
        expect(surface.selectText?.("The later notice period is thirtysixty days")).toBe(true);
        const highlighted = Array.from(content.querySelectorAll(`.${QUOTE_HIGHLIGHT_CLASS}`));
        expect(highlighted.length).toBeGreaterThan(1);
        expect(highlighted.every((span) => span.closest('[data-anchor="c"]'))).toBe(true);
        expect(content.textContent).toBe(before);
        surface.clearTextSelection?.();
        expect(content.querySelector(`.${QUOTE_HIGHLIGHT_CLASS}`)).toBeNull();
        expect(content.textContent).toBe(before);
    });

    it("returns false for a quote the document does not contain", () => {
        const { surface } = mountDocument(REVISIONS, NOTICE);
        expect(surface.selectText?.("Governing law is Delaware")).toBe(false);
    });

    it("replaces a citation highlight when an edit is activated", () => {
        const { content, surface } = mountDocument(REVISIONS, NOTICE);
        surface.selectText?.("Payment is due within thirty days.");
        surface.activateRevision?.({ ins: "8", del: "7" });
        expect(content.querySelector(`.${QUOTE_HIGHLIGHT_CLASS}`)).toBeNull();
        expect(active(content)).toEqual(["d7", "i8"]);
    });

    it("downloads a copy of exactly the bytes that were opened", async () => {
        const { surface, bytes } = mountDocument(REVISIONS, NOTICE);
        const exported = await surface.exportDocx!();
        expect(exported).not.toBe(bytes);
        expect(Array.from(new Uint8Array(exported))).toEqual([80, 75, 3, 4]);
    });
});
