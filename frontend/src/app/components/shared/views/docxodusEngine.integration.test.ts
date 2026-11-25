// @vitest-environment node
//
// Runs the real Docxodus WebAssembly engine (no mocks) against a Word document
// built the way Mike's tracked-edit pipeline writes one. It pins the two engine
// behaviours the view-only adapter depends on, so a Docxodus upgrade that
// changes either fails here instead of in a browser:
//   1. tracked changes render as <ins>/<del> elements, and
//   2. the revision registry exposes each change's OOXML w:id, which is how an
//      edit card (ins_w_id / del_w_id) finds its change in the document.
import { beforeAll, describe, expect, it } from "vitest";
import { DeletedTextRun, Document, HeadingLevel, InsertedTextRun, Packer, Paragraph, TextRun } from "docx";
import { JSDOM } from "jsdom";
import { convertDocxToHtml, getRevisions, initialize } from "docxodus/core";

const DATE = "2026-09-23T00:00:00Z";
// The engine runs under Node here, where there is no DOM to parse its HTML.
const bodyOf = (html: string) => new JSDOM(html).window.document.body;

async function trackedAgreement(): Promise<Uint8Array> {
    const notice = (label: string, del: number, ins: number) => new Paragraph({ children: [
        new TextRun(`${label} notice period is `),
        new DeletedTextRun({ text: "thirty", id: del, author: "Reviewer", date: DATE }),
        new InsertedTextRun({ text: "sixty", id: ins, author: "Reviewer", date: DATE }),
        new TextRun(" days."),
    ] });
    const document = new Document({ sections: [{ children: [
        new Paragraph({ text: "Sample services agreement", heading: HeadingLevel.HEADING_1 }),
        new Paragraph("Payment is due within thirty days."),
        notice("The", 7, 8),
        // Same wording, different ids: only the id can tell the two edits apart.
        notice("The later", 27, 28),
    ] }] });
    return new Uint8Array(await Packer.toBuffer(document));
}

describe("Docxodus engine (real WebAssembly runtime)", () => {
    let docx: Uint8Array;
    beforeAll(async () => {
        await initialize();
        docx = await trackedAgreement();
    }, 60_000);

    it("renders tracked changes inline with their author", async () => {
        const html = await convertDocxToHtml(docx, { renderTrackedChanges: true });
        const body = bodyOf(html);
        const texts = (selector: string) => Array.from(body.querySelectorAll(selector)).map((element) => element.textContent);
        expect(texts("ins")).toEqual(["sixty", "sixty"]);
        expect(texts("del")).toEqual(["thirty", "thirty"]);
        expect(body.querySelector("ins")?.getAttribute("data-author")).toBe("Reviewer");
        expect(body.textContent).toContain("Payment is due within thirty days.");
    });

    it("renders the accepted text when tracked changes are not shown", async () => {
        const html = await convertDocxToHtml(docx, { renderTrackedChanges: false });
        const body = bodyOf(html);
        expect(body.querySelector("ins, del")).toBeNull();
        expect(body.textContent).toContain("The notice period is sixty days.");
        expect(body.textContent).not.toContain("thirtysixty");
    });

    it("lists every tracked change with the w:id Mike stores on its edit cards", async () => {
        const revisions = await getRevisions(docx);
        const byId = new Map(revisions.map((revision) => [revision.constituentIds.join(","), revision]));
        expect([...byId.keys()].sort()).toEqual(["27", "28", "7", "8"]);
        expect(byId.get("28")).toMatchObject({ type: "insert", text: "sixty", author: "Reviewer" });
        expect(byId.get("27")).toMatchObject({ type: "delete", text: "thirty", author: "Reviewer" });
    });

    it("rejects bytes that are not a Word document", async () => {
        await expect(convertDocxToHtml(new TextEncoder().encode("invalid docx"))).rejects.toThrow();
    });
});
