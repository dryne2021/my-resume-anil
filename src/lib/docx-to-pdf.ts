/**
 * Converts a generated .docx (bytes) into a PDF that visually matches the Word
 * document. Rendering happens inside an isolated iframe so the app's Tailwind
 * v4 stylesheet (which uses modern `lab()` colors html2canvas cannot parse)
 * never leaks into the captured DOM.
 */

const FONT_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Tinos:ital,wght@0,400;0,700;1,400;1,700&display=swap');
html, body { margin: 0; padding: 0; background: #ffffff; color: #000000; }
.docx-wrapper { background: #ffffff !important; padding: 0 !important; }
.docx-wrapper > section.docx { box-shadow: none !important; margin: 0 !important; background: #ffffff !important; }
`;

export async function docxToPdfBlob(bytes: Uint8Array): Promise<Blob> {
  const [{ renderAsync }, html2canvasMod, jsPdfMod] = await Promise.all([
    import("docx-preview"),
    import("html2canvas"),
    import("jspdf"),
  ]);
  const html2canvas = html2canvasMod.default;
  const { jsPDF } = jsPdfMod;

  const iframe = document.createElement("iframe");
  iframe.style.position = "fixed";
  iframe.style.left = "-10000px";
  iframe.style.top = "0";
  iframe.style.width = "900px";
  iframe.style.height = "1200px";
  iframe.style.border = "0";
  document.body.appendChild(iframe);

  try {
    const doc = iframe.contentDocument!;
    doc.open();
    doc.write(`<!doctype html><html><head><meta charset="utf-8"><style>${FONT_CSS}</style></head><body><div id="root"></div></body></html>`);
    doc.close();

    const container = doc.getElementById("root")!;
    const buf = bytes.slice().buffer as ArrayBuffer;
    await renderAsync(buf, container, undefined, {
      inWrapper: true,
      breakPages: true,
      ignoreWidth: false,
      ignoreHeight: false,
      experimental: true,
      useBase64URL: true,
    });

    // Let webfonts settle before rasterizing.
    try {
      await (doc as Document & { fonts?: FontFaceSet }).fonts?.ready;
    } catch {
      /* ignore */
    }
    await new Promise((r) => setTimeout(r, 250));

    const pages = Array.from(
      container.querySelectorAll<HTMLElement>(".docx-wrapper > section.docx"),
    );
    const targets = pages.length > 0 ? pages : [container];

    const pdf = new jsPDF({ unit: "pt", format: "letter", orientation: "portrait" });
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();

    for (let i = 0; i < targets.length; i++) {
      const el = targets[i];
      const canvas = await html2canvas(el, {
        scale: 2,
        backgroundColor: "#ffffff",
        useCORS: true,
        logging: false,
        windowWidth: el.scrollWidth,
        windowHeight: el.scrollHeight,
      });
      const img = canvas.toDataURL("image/jpeg", 0.95);

      // Fit the rendered page onto a Letter sheet, top-aligned.
      const ratio = Math.min(pageW / canvas.width, pageH / canvas.height);
      const w = canvas.width * ratio;
      const h = canvas.height * ratio;
      const x = (pageW - w) / 2;

      if (i > 0) pdf.addPage("letter", "portrait");
      pdf.addImage(img, "JPEG", x, 0, w, h);
    }

    return pdf.output("blob");
  } finally {
    iframe.remove();
  }
}
