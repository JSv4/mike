// The standalone DOCX catalog targets the same modern browsers as the app.
const config = {
    // Supply an empty browser environment for Next's Sentry adapter.
    define: { "process.env": "{}" },
    build: { target: "esnext" },
    optimizeDeps: {
        // Preserve EigenPal's import.meta.url font and WASM assets, and
        // Docxodus's runtime import of its WebAssembly loader.
        exclude: ["@docx-editor.dev/fonts", "@docx-editor.dev/core", "@docx-editor.dev/react", "docxodus"],
        esbuildOptions: { target: "esnext" },
    },
};

export default config;
