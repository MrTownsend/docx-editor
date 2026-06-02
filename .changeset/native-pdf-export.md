---
'@eigenpal/docx-editor-core': patch
'@eigenpal/docx-editor-react': patch
'@eigenpal/docx-editor-vue': patch
---

Add native vector PDF export. A new core `exportToPdf` (dynamically imported `@eigenpal/docx-editor-core/pdf`) walks the editor's computed layout to produce a real vector PDF — selectable text, embedded subset fonts, vector tables/borders, and embedded images. Both adapters gain an `exportPdf()` ref method and a File ▸ Export ▸ (.docx / .pdf) menu, and printing now routes through the generated PDF (no DOM-clone style loss). In-browser only.
