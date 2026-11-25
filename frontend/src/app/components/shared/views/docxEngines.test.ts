import { afterEach, describe, expect, it, vi } from "vitest";
import { activeDocxEngine, DEFAULT_DOCX_ENGINE, getDocxEngine, isDocxEngineId, resolveDocxEngineId } from "./docxEngines";

afterEach(() => vi.unstubAllEnvs());

describe("docx engine registry", () => {
    it("keeps the existing engine as the default", () => {
        expect(DEFAULT_DOCX_ENGINE).toBe("eigenpal");
        expect(activeDocxEngine().id).toBe("eigenpal");
    });

    it("selects the configured engine, ignoring case and whitespace", () => {
        expect(resolveDocxEngineId(" Docxodus ")).toBe("docxodus");
        vi.stubEnv("NEXT_PUBLIC_DOCX_ENGINE", "docxodus");
        expect(activeDocxEngine().id).toBe("docxodus");
    });

    it.each([undefined, null, "", "superdoc", "constructor", "__proto__"])(
        "falls back to the default for %j instead of breaking previews", (configured) => {
            expect(resolveDocxEngineId(configured)).toBe(DEFAULT_DOCX_ENGINE);
        });

    it("only recognises registered engine names", () => {
        expect(isDocxEngineId("eigenpal")).toBe(true);
        expect(isDocxEngineId("toString")).toBe(false);
        expect(isDocxEngineId(7)).toBe(false);
    });

    it("reports Docxodus as view-only so hosts do not offer editing", () => {
        expect(getDocxEngine("docxodus").capabilities).toEqual({ edit: false, suggest: false, comment: false });
        expect(getDocxEngine("eigenpal").capabilities.edit).toBe(true);
    });
});
