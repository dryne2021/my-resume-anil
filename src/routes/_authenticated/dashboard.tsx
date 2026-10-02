import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useEffect, useRef, useState } from "react";
import { useServerFn } from "@tanstack/react-start";
import { toast } from "sonner";
import { Upload, FileText, Sparkles, Download, LogOut, Check, Trash2 } from "lucide-react";

import { logout } from "@/lib/auth.functions";
import { generateResume } from "@/lib/generate-resume.functions";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";


export const Route = createFileRoute("/_authenticated/dashboard")({
  head: () => ({
    meta: [{ title: "Dashboard — Dryne Agency" }, { name: "robots", content: "noindex,nofollow" }],
  }),
  component: Dashboard,
});

interface StoredTemplate {
  file_name: string;
  updated_at: string;
  base64: string;
}

// The template is remembered only in this browser (localStorage) so you don't
// have to re-upload it. Nothing is saved on the server.
const TEMPLATE_KEY = "dryne_template_v1";
function loadTemplate(): StoredTemplate | null {
  try {
    const raw = localStorage.getItem(TEMPLATE_KEY);
    return raw ? (JSON.parse(raw) as StoredTemplate) : null;
  } catch {
    return null;
  }
}

function saveTemplate(t: StoredTemplate | null) {
  try {
    if (t) localStorage.setItem(TEMPLATE_KEY, JSON.stringify(t));
    else localStorage.removeItem(TEMPLATE_KEY);
  } catch {
    /* storage unavailable — template stays in memory for this visit */
  }
}

function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(",")[1] ?? "");
    reader.onerror = () => reject(reader.error ?? new Error("Could not read file"));
    reader.readAsDataURL(file);
  });
}




function Dashboard() {
  const signOut = useServerFn(logout);
  const navigate = useNavigate();
  const generate = useServerFn(generateResume);

  const [template, setTemplate] = useState<StoredTemplate | null>(null);
  const [loadingTpl, setLoadingTpl] = useState(true);
  const [uploading, setUploading] = useState(false);
  const [jd, setJd] = useState("");
  const [busy, setBusy] = useState(false);
  const [elapsed, setElapsed] = useState(0);

  useEffect(() => {
    if (!busy) return;
    setElapsed(0);
    const started = Date.now();
    const id = setInterval(() => setElapsed(Math.floor((Date.now() - started) / 1000)), 1000);
    return () => clearInterval(id);
  }, [busy]);
  const [pdfBusy, setPdfBusy] = useState(false);
  const [result, setResult] = useState<{ fileName: string; base64: string } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setTemplate(loadTemplate());
    setLoadingTpl(false);
  }, []);

  const handleUpload = async (file: File) => {
    if (!file.name.toLowerCase().endsWith(".docx")) {
      toast.error("Please upload a .docx file");
      return;
    }
    if (file.size > 3 * 1024 * 1024) {
      toast.error("File too large (max 3MB)");
      return;
    }
    setUploading(true);
    try {
      const t: StoredTemplate = {
        file_name: file.name,
        updated_at: new Date().toISOString(),
        base64: await fileToBase64(file),
      };
      setTemplate(t);
      saveTemplate(t);
      toast.success("Master template ready");
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "Upload failed");
    } finally {
      setUploading(false);
    }
  };

  const handleRemoveTemplate = () => {
    setTemplate(null);
    saveTemplate(null);
    setResult(null);
  };

  const handleGenerate = async () => {
    if (!template) {
      toast.error("Upload your master template first");
      return;
    }
    if (jd.trim().length < 20) {
      toast.error("Paste a job description (min ~20 characters)");
      return;
    }
    setBusy(true);
    setResult(null);
    try {
      const res = await generate({
        data: { jobDescription: jd, templateBase64: template.base64 },
      });
      const base = template.file_name.replace(/\.docx$/i, "");
      setResult({ ...res, fileName: `${base}_tailored.docx` });
      toast.success("Resume generated");
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "Generation failed");
    } finally {
      setBusy(false);
    }
  };






  const resultBytes = () => {
    const bin = atob(result!.base64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  };

  const saveBlob = (blob: Blob, fileName: string) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleDownloadDocx = () => {
    if (!result) return;
    try {
      const blob = new Blob([resultBytes()], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      saveBlob(blob, result.fileName.endsWith(".docx") ? result.fileName : `${result.fileName}.docx`);
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "Word export failed");
    }
  };

  const handleDownloadPdf = async () => {
    if (!result) return;
    setPdfBusy(true);
    try {
      const { docxToPdfBlob } = await import("@/lib/docx-to-pdf");
      const blob = await docxToPdfBlob(resultBytes());
      saveBlob(blob, result.fileName.replace(/\.docx$/i, "") + ".pdf");
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "PDF export failed");
    } finally {
      setPdfBusy(false);
    }
  };


  return (
    <div className="min-h-screen grain">
      <header className="border-b border-border/60">
        <div className="max-w-5xl mx-auto px-6 py-5 flex items-center justify-between">
          <div>
            <p className="text-[10px] uppercase tracking-[0.3em] text-muted-foreground">
              Private Portal
            </p>
            <h1 className="text-2xl font-display text-gold leading-none mt-1">Dryne Agency</h1>
          </div>
          <Button
            variant="ghost"
            size="sm"
            onClick={async () => {
              await signOut({});
              navigate({ to: "/login" });
            }}
            className="text-muted-foreground hover:text-foreground"
          >
            <LogOut className="w-4 h-4 mr-2" /> Sign out
          </Button>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-12 space-y-10">
        <section>
          <h2 className="text-4xl font-display">Resume Atelier</h2>
          <p className="text-muted-foreground mt-2 max-w-xl">
            Upload your master Word template once. Paste any job description. Receive a tailored{" "}
            <em className="text-gold">Word document</em> that keeps your exact layout, fonts,
            headings, and structure.
          </p>

        </section>

        {/* Step 1: Template */}
        <section className="rounded-xl border border-border bg-card/60 p-6">
          <div className="flex items-start justify-between gap-6">
            <div>
              <p className="text-[10px] uppercase tracking-[0.25em] text-gold">Step 01</p>
              <h3 className="text-2xl font-display mt-1">Master Template</h3>
              <p className="text-sm text-muted-foreground mt-1">
                Used for layout and format only; the content comes from the job description. Remembered in this browser only.
              </p>
            </div>
            <input
              ref={fileInputRef}
              type="file"
              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              className="hidden"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) handleUpload(f);
                e.target.value = "";
              }}
            />
            <Button
              onClick={() => fileInputRef.current?.click()}
              disabled={uploading}
              variant="outline"
              className="border-gold/40 text-gold hover:bg-gold/10 hover:text-gold"
            >
              <Upload className="w-4 h-4 mr-2" />
              {uploading ? "Uploading…" : template ? "Replace" : "Upload .docx"}
            </Button>
          </div>

          <div className="mt-5 rounded-lg border border-dashed border-border bg-background/40 p-4 flex items-center gap-3">
            {loadingTpl ? (
              <span className="text-sm text-muted-foreground">Loading…</span>
            ) : template ? (
              <>
                <Check className="w-5 h-5 text-gold" />
                <FileText className="w-5 h-5 text-muted-foreground" />
                <div className="flex-1 min-w-0">
                  <p className="text-sm truncate">{template.file_name}</p>
                  <p className="text-xs text-muted-foreground">
                    Added · {new Date(template.updated_at).toLocaleString()}
                  </p>
                </div>
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={handleRemoveTemplate}
                  className="text-muted-foreground hover:text-foreground"
                  aria-label="Remove template"
                >
                  <Trash2 className="w-4 h-4" />
                </Button>
              </>
            ) : (
              <>
                <FileText className="w-5 h-5 text-muted-foreground" />
                <p className="text-sm text-muted-foreground">
                  No template uploaded yet. Upload your sample resume to begin.
                </p>
              </>
            )}
          </div>
        </section>

        {/* Step 2: JD + Generate */}
        <section className="rounded-xl border border-border bg-card/60 p-6">
          <p className="text-[10px] uppercase tracking-[0.25em] text-gold">Step 02</p>
          <h3 className="text-2xl font-display mt-1">Job Description</h3>
          <p className="text-sm text-muted-foreground mt-1">
            Paste the full posting. Personal details, dates, schools, and companies stay intact.
          </p>

          <Textarea
            value={jd}
            onChange={(e) => setJd(e.target.value)}
            placeholder="Paste the job description here…"
            rows={12}
            className="mt-5 bg-background/40 resize-y min-h-[260px] font-sans text-sm"
          />

          <div className="mt-5 flex flex-wrap items-center gap-3">
            <Button
              onClick={handleGenerate}
              disabled={busy || !template}
              className="bg-gold-gradient text-primary-foreground hover:opacity-90"
            >
              <Sparkles className="w-4 h-4 mr-2" />
              {busy
                ? `Generating… ${Math.floor(elapsed / 60)}:${String(elapsed % 60).padStart(2, "0")}`
                : "Generate Resume"}
            </Button>
            {!template && (
              <span className="text-xs text-muted-foreground">Upload a template first.</span>
            )}
            {busy && (
              <span className="text-xs text-muted-foreground">
                Usually takes 1–2 minutes. Keep this tab open.
              </span>
            )}
          </div>
        </section>

        {/* Result */}
        {result && (
          <section className="rounded-xl border border-gold/40 bg-card/80 p-6">
            <p className="text-[10px] uppercase tracking-[0.25em] text-gold">Ready</p>
            <h3 className="text-2xl font-display mt-1">{result.fileName}</h3>
            <p className="text-sm text-muted-foreground mt-1">
              Your tailored resume is ready — same layout, fonts, and structure as your uploaded
              template.
            </p>
            <div className="mt-5 flex flex-wrap gap-3">
              <Button
                onClick={handleDownloadDocx}
                className="bg-gold-gradient text-primary-foreground hover:opacity-90"
              >
                <Download className="w-4 h-4 mr-2" />
                Download Word
              </Button>
              <Button
                onClick={handleDownloadPdf}
                disabled={pdfBusy}
                variant="outline"
                className="border-gold/40 text-gold hover:bg-gold/10 hover:text-gold"
              >
                <Download className="w-4 h-4 mr-2" />
                {pdfBusy ? "Preparing PDF…" : "Download PDF"}
              </Button>
            </div>
          </section>
        )}


        <footer className="pt-6 text-center text-xs text-muted-foreground">
          Dryne Agency · Private Resume Portal
        </footer>
      </main>
    </div>
  );
}
