import { useState, useEffect, useRef, useCallback } from "react";
import { useParams, useLocation } from "wouter";
import { ArrowLeft, FileText, Loader2, Printer, Download, Image as ImageIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { generateSummaryStream, fetchSummary } from "@/lib/api";
import { useToast } from "@/hooks/use-toast";

export default function ExecutiveSummaryPage() {
  const params = useParams<{ id: string }>();
  const [, navigate] = useLocation();
  const { toast } = useToast();
  const reportId = Number(params.id);

  const [content, setContent] = useState("");
  const [isGenerating, setIsGenerating] = useState(false);
  const [isDone, setIsDone] = useState(false);
  const [wordCount, setWordCount] = useState(0);
  const [elapsedTime, setElapsedTime] = useState(0);
  const [isExporting, setIsExporting] = useState<"pdf" | "png" | null>(null);
  const contentRef = useRef<HTMLDivElement>(null);
  const exportRef = useRef<HTMLDivElement>(null);
  const timerRef = useRef<ReturnType<typeof setInterval> | null>(null);

  useEffect(() => {
    fetchSummary(reportId).then((existing) => {
      if (existing) {
        setContent(existing.content);
        setIsDone(true);
        setWordCount(existing.content.split(/\s+/).filter(Boolean).length);
      } else {
        const stored = sessionStorage.getItem(`summary_payload_${reportId}`);
        if (stored) {
          const payload = JSON.parse(stored);
          startGeneration(payload);
        }
      }
    });
    return () => {
      if (timerRef.current) clearInterval(timerRef.current);
    };
  }, [reportId]);

  const startGeneration = (payload: any) => {
    setIsGenerating(true);
    setContent("");
    setElapsedTime(0);
    const startTime = Date.now();
    timerRef.current = setInterval(() => {
      setElapsedTime(Math.floor((Date.now() - startTime) / 1000));
    }, 1000);

    generateSummaryStream(
      reportId,
      payload,
      (chunk) => {
        setContent((prev) => {
          const next = prev + chunk;
          setWordCount(next.split(/\s+/).filter(Boolean).length);
          return next;
        });
      },
      () => {
        setIsGenerating(false);
        setIsDone(true);
        if (timerRef.current) clearInterval(timerRef.current);
        sessionStorage.removeItem(`summary_payload_${reportId}`);
      },
      (err) => {
        setIsGenerating(false);
        if (timerRef.current) clearInterval(timerRef.current);
        setContent((prev) => prev + `\n\nError: ${err}`);
      }
    );
  };

  const handleRegenerate = () => {
    const stored = sessionStorage.getItem(`summary_payload_${reportId}`);
    if (stored) {
      startGeneration(JSON.parse(stored));
    }
  };

  useEffect(() => {
    if (contentRef.current && isGenerating) {
      contentRef.current.scrollTop = contentRef.current.scrollHeight;
    }
  }, [content, isGenerating]);

  const renderMarkdown = (md: string) => {
    const lines = md.split("\n");
    const blocks: string[] = [];
    let i = 0;

    const inline = (text: string) =>
      text
        .replace(/\*\*(.+?)\*\*/g, '<strong class="font-semibold text-foreground">$1</strong>')
        .replace(/\*(.+?)\*/g, "<em>$1</em>")
        .replace(/`(.+?)`/g, '<code class="px-1.5 py-0.5 bg-muted/60 rounded text-xs font-mono">$1</code>');

    const cellClass = (raw: string) => {
      const hasEmoji = /🔴|🟡|🟢/.test(raw);
      const hasDollar = /\$[\d,]+/.test(raw);
      const hasPlus = /^\+/.test(raw.trim());
      const hasMinus = /^[-−]/.test(raw.trim()) && hasDollar;
      let cls = "text-sm";
      if (hasDollar && hasPlus) cls += " text-red-600 dark:text-red-400 font-medium";
      else if (hasDollar && hasMinus) cls += " text-emerald-600 dark:text-emerald-400 font-medium";
      else if (hasDollar) cls += " font-medium tabular-nums";
      if (hasEmoji) cls += " text-center";
      return cls;
    };

    while (i < lines.length) {
      const line = lines[i];

      if (/^#{1,4} /.test(line)) {
        const match = line.match(/^(#{1,4}) (.+)$/);
        if (match) {
          const level = match[1].length;
          const text = inline(match[2]);
          if (level === 1)
            blocks.push(`<h1 class="text-2xl font-bold mt-2 mb-6 font-display text-foreground text-center tracking-tight">${text}</h1>`);
          else if (level === 2)
            blocks.push(`<div class="mt-10 mb-4"><h2 class="text-xl font-bold font-display text-primary pb-2 border-b-2 border-primary/20">${text}</h2></div>`);
          else if (level === 3)
            blocks.push(`<h3 class="text-lg font-bold mt-7 mb-3 font-display text-foreground flex items-center gap-2"><span class="inline-block w-1 h-5 bg-primary/70 rounded-full"></span>${text}</h3>`);
          else
            blocks.push(`<h4 class="text-base font-semibold mt-5 mb-2 font-display text-foreground/90">${text}</h4>`);
        }
        i++;
        continue;
      }

      if (/^---+$|^\*\*\*+$|^___+$/.test(line.trim())) {
        blocks.push('<hr class="my-6 border-border/40" />');
        i++;
        continue;
      }

      if (/^> /.test(line)) {
        const bqLines: string[] = [];
        while (i < lines.length && /^> (.*)$/.test(lines[i])) {
          bqLines.push(lines[i].replace(/^> /, ""));
          i++;
        }
        blocks.push(`<blockquote class="border-l-4 border-primary/60 bg-primary/5 pl-4 pr-3 py-3 my-4 text-sm italic text-foreground/90 rounded-r-md">${bqLines.map(inline).join("<br/>")}</blockquote>`);
        continue;
      }

      if (/^\|/.test(line)) {
        const tableRows: string[][] = [];
        let headerIdx = -1;
        let rowIdx = 0;
        while (i < lines.length && /^\|/.test(lines[i])) {
          const cells = lines[i].split("|").filter(Boolean).map((c) => c.trim());
          if (cells.every((c) => /^[-:]+$/.test(c))) {
            headerIdx = rowIdx;
          } else {
            tableRows.push(cells);
            rowIdx++;
          }
          i++;
        }

        let tableHtml = "";
        tableRows.forEach((cells, idx) => {
          const isHeader = headerIdx >= 0 && idx === 0;
          const tag = isHeader ? "th" : "td";
          const trClass = isHeader
            ? "bg-muted/50 font-medium text-xs uppercase tracking-wider"
            : "hover:bg-muted/20 transition-colors";
          tableHtml += `<tr class="${trClass}">${cells
            .map((c) => {
              const cls = isHeader ? "text-sm font-medium" : cellClass(c);
              return `<${tag} class="border border-border/40 px-3 py-2.5 ${cls}">${inline(c)}</${tag}>`;
            })
            .join("")}</tr>`;
        });

        blocks.push(`<div class="overflow-x-auto my-5 rounded-lg border border-border/50 shadow-sm"><table class="w-full border-collapse text-sm">${tableHtml}</table></div>`);
        continue;
      }

      if (/^(\d+)\. /.test(line)) {
        const items: string[] = [];
        while (i < lines.length && /^(\d+)\. (.+)$/.test(lines[i])) {
          const m = lines[i].match(/^(\d+)\. (.+)$/);
          if (m) items.push(`<li class="ml-5 mb-2 text-sm leading-relaxed list-decimal"><span class="font-semibold text-primary">${m[1]}.</span> ${inline(m[2])}</li>`);
          i++;
        }
        blocks.push(`<ol class="my-3 space-y-1">${items.join("")}</ol>`);
        continue;
      }

      if (/^[-*] /.test(line)) {
        const items: string[] = [];
        while (i < lines.length && /^[-*] (.+)$/.test(lines[i])) {
          const m = lines[i].match(/^[-*] (.+)$/);
          if (m) items.push(`<li class="ml-5 mb-1.5 text-sm leading-relaxed list-disc">${inline(m[1])}</li>`);
          i++;
        }
        blocks.push(`<ul class="my-3 space-y-0.5">${items.join("")}</ul>`);
        continue;
      }

      if (line.trim() === "") {
        i++;
        continue;
      }

      const paraLines: string[] = [];
      while (i < lines.length && lines[i].trim() !== "" && !/^#{1,4} |^\||^[-*] |^\d+\. |^> |^---+$|^\*\*\*+$|^___+$/.test(lines[i])) {
        paraLines.push(lines[i]);
        i++;
      }
      blocks.push(`<p class="mb-3 text-sm leading-relaxed text-foreground/85">${inline(paraLines.join(" "))}</p>`);
    }

    return `<div class="text-sm leading-relaxed text-foreground/85">${blocks.join("\n")}</div>`;
  };

  const resolveColor = (val: string): string => {
    const c = document.createElement("canvas");
    c.width = 1;
    c.height = 1;
    const ctx = c.getContext("2d");
    if (!ctx) return val;
    ctx.fillStyle = val;
    ctx.fillRect(0, 0, 1, 1);
    const [r, g, b, a] = ctx.getImageData(0, 0, 1, 1).data;
    return a < 255 ? `rgba(${r},${g},${b},${(a / 255).toFixed(3)})` : `rgb(${r},${g},${b})`;
  };

  const inlineAllColors = (node: HTMLElement) => {
    const computed = window.getComputedStyle(node);
    const colorProps = [
      "color", "background-color", "border-color",
      "border-top-color", "border-bottom-color",
      "border-left-color", "border-right-color",
      "outline-color", "text-decoration-color",
      "box-shadow",
    ];
    for (const prop of colorProps) {
      const val = computed.getPropertyValue(prop);
      if (val && val !== "none" && val !== "initial" && val !== "inherit") {
        if (val.includes("oklab") || val.includes("oklch") || val.includes("color-mix") || val.includes("lab(") || val.includes("lch(")) {
          if (prop === "box-shadow") {
            node.style.setProperty(prop, val.replace(/(?:oklab|oklch|color-mix|lab|lch)\([^)]*(?:\([^)]*\))*[^)]*\)/g, (match) => resolveColor(match)));
          } else {
            node.style.setProperty(prop, resolveColor(val));
          }
        } else {
          node.style.setProperty(prop, val);
        }
      }
    }

    const bg = computed.getPropertyValue("background");
    if (bg && (bg.includes("oklab") || bg.includes("oklch") || bg.includes("color-mix"))) {
      node.style.setProperty("background", "none");
      const bgColor = computed.getPropertyValue("background-color");
      if (bgColor) node.style.setProperty("background-color", resolveColor(bgColor));
    }

    for (const child of Array.from(node.children)) {
      if (child instanceof HTMLElement) inlineAllColors(child);
    }
  };

  const captureContent = useCallback(async () => {
    const el = exportRef.current;
    if (!el) throw new Error("Content not available for export");

    const iframe = document.createElement("iframe");
    iframe.style.position = "fixed";
    iframe.style.left = "-10000px";
    iframe.style.top = "0";
    iframe.style.width = el.scrollWidth + "px";
    iframe.style.height = el.scrollHeight + "px";
    iframe.style.border = "none";
    document.body.appendChild(iframe);

    try {
      const iframeDoc = iframe.contentDocument!;
      iframeDoc.open();
      iframeDoc.write("<!DOCTYPE html><html><head></head><body></body></html>");
      iframeDoc.close();

      const flattenStyles = (source: HTMLElement, target: HTMLElement) => {
        const computed = window.getComputedStyle(source);
        const props = [
          "color", "background-color", "border-color",
          "border-top-color", "border-bottom-color",
          "border-left-color", "border-right-color",
          "font-family", "font-size", "font-weight", "font-style",
          "line-height", "letter-spacing", "text-align", "text-decoration",
          "text-transform", "white-space", "word-spacing",
          "padding-top", "padding-right", "padding-bottom", "padding-left",
          "margin-top", "margin-right", "margin-bottom", "margin-left",
          "display", "flex-direction", "align-items", "justify-content", "gap",
          "width", "min-width", "max-width",
          "position", "top", "right", "bottom", "left",
          "border-width", "border-style", "border-radius",
          "border-top-width", "border-right-width", "border-bottom-width", "border-left-width",
          "border-top-style", "border-right-style", "border-bottom-style", "border-left-style",
          "opacity", "visibility", "vertical-align", "list-style-type",
          "table-layout", "border-collapse", "border-spacing",
          "flex-grow", "flex-shrink", "flex-basis", "flex-wrap", "order",
        ];
        for (const prop of props) {
          let val = computed.getPropertyValue(prop);
          if (val && val !== "initial" && val !== "inherit") {
            if (val.includes("oklab") || val.includes("oklch") || val.includes("color-mix") || val.includes("lab(") || val.includes("lch(")) {
              val = resolveColor(val);
            }
            target.style.setProperty(prop, val);
          }
        }
        target.style.boxShadow = "none";
        target.style.overflow = "visible";
        target.style.overflowX = "visible";
        target.style.overflowY = "visible";
        target.style.height = "auto";
        target.style.minHeight = "0";
        target.style.maxHeight = "none";

        const sourceChildren = source.children;
        const targetChildren = target.children;
        for (let i = 0; i < sourceChildren.length; i++) {
          if (sourceChildren[i] instanceof HTMLElement && targetChildren[i] instanceof HTMLElement) {
            flattenStyles(sourceChildren[i] as HTMLElement, targetChildren[i] as HTMLElement);
          }
        }
      };

      const clone = el.cloneNode(true) as HTMLElement;
      iframeDoc.body.appendChild(clone);
      iframeDoc.body.style.margin = "0";
      iframeDoc.body.style.padding = "0";
      clone.style.margin = "0";
      clone.style.position = "static";
      clone.style.overflow = "visible";

      flattenStyles(el, clone);

      clone.style.backgroundColor = "#ffffff";

      const cloneHeight = clone.scrollHeight;
      const cloneWidth = clone.scrollWidth;
      iframe.style.height = cloneHeight + "px";
      iframe.style.width = cloneWidth + "px";

      const html2canvas = (await import("html2canvas")).default;
      const canvas = await html2canvas(clone, {
        scale: 2,
        useCORS: true,
        backgroundColor: "#ffffff",
        logging: false,
        width: cloneWidth,
        height: cloneHeight,
        windowWidth: cloneWidth,
        windowHeight: cloneHeight,
      });
      return canvas;
    } finally {
      document.body.removeChild(iframe);
    }
  }, []);

  const handleExportPDF = useCallback(async () => {
    if (!exportRef.current) return;
    setIsExporting("pdf");
    try {
      const fullCanvas = await captureContent();
      const { jsPDF } = await import("jspdf");

      const pdfPageWidth = 210;
      const pdfPageHeight = 297;
      const margin = 10;
      const contentWidth = pdfPageWidth - margin * 2;
      const contentHeight = pdfPageHeight - margin * 2;

      const scaleFactor = contentWidth / fullCanvas.width;
      const pageCanvasHeight = contentHeight / scaleFactor;
      const totalPages = Math.ceil(fullCanvas.height / pageCanvasHeight);

      const pdf = new jsPDF("p", "mm", "a4");

      for (let page = 0; page < totalPages; page++) {
        if (page > 0) pdf.addPage();

        const sourceY = page * pageCanvasHeight;
        const sourceH = Math.min(pageCanvasHeight, fullCanvas.height - sourceY);

        const pageCanvas = document.createElement("canvas");
        pageCanvas.width = fullCanvas.width;
        pageCanvas.height = sourceH;
        const ctx = pageCanvas.getContext("2d");
        if (!ctx) continue;

        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
        ctx.drawImage(
          fullCanvas,
          0, sourceY, fullCanvas.width, sourceH,
          0, 0, fullCanvas.width, sourceH
        );

        const pageImgData = pageCanvas.toDataURL("image/jpeg", 0.92);
        const imgH = sourceH * scaleFactor;
        pdf.addImage(pageImgData, "JPEG", margin, margin, contentWidth, imgH);
      }

      const dateStr = new Date().toISOString().split("T")[0];
      pdf.save(`M365-Executive-Briefing-${dateStr}.pdf`);
      toast({ title: "PDF exported", description: `Saved as ${totalPages}-page PDF.` });
    } catch (err: any) {
      console.error("PDF export error:", err);
      toast({ title: "Export failed", description: err.message, variant: "destructive" });
    } finally {
      setIsExporting(null);
    }
  }, [captureContent, toast]);

  const handleExportPNG = useCallback(async () => {
    if (!exportRef.current) return;
    setIsExporting("png");
    try {
      const canvas = await captureContent();
      const link = document.createElement("a");
      const dateStr = new Date().toISOString().split("T")[0];
      link.download = `M365-Executive-Briefing-${dateStr}.png`;
      link.href = canvas.toDataURL("image/png");
      link.click();
      toast({ title: "Image exported", description: "Your executive briefing has been saved as a PNG image." });
    } catch (err: any) {
      console.error("PNG export error:", err);
      toast({ title: "Export failed", description: err.message, variant: "destructive" });
    } finally {
      setIsExporting(null);
    }
  }, [captureContent, toast]);

  const formatTime = (seconds: number) => {
    const m = Math.floor(seconds / 60);
    const s = seconds % 60;
    return m > 0 ? `${m}m ${s}s` : `${s}s`;
  };

  return (
    <div className="min-h-screen bg-gradient-to-b from-background to-muted/20 flex flex-col font-sans text-foreground">
      <header className="sticky top-0 z-10 bg-card/90 backdrop-blur-lg border-b border-border/50 px-6 py-3 flex items-center justify-between shadow-sm">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 cursor-pointer" onClick={() => navigate("/")} data-testid="link-home">
            <div className="h-7 w-7 rounded-md bg-primary flex items-center justify-center text-primary-foreground font-bold text-sm">
              A
            </div>
          </div>
          <nav className="hidden sm:flex items-center gap-1">
            <Button variant="ghost" size="sm" className="text-muted-foreground hover:text-foreground" onClick={() => navigate("/")} data-testid="nav-dashboard">Dashboard</Button>
            <Button variant="ghost" size="sm" className="text-muted-foreground hover:text-foreground" onClick={() => navigate("/licenses")} data-testid="nav-licenses">License Guide</Button>
          </nav>
          <span className="text-border hidden sm:inline">|</span>
          <Button variant="ghost" size="sm" onClick={() => navigate("/")} data-testid="button-back" className="gap-2 sm:hidden">
            <ArrowLeft className="h-4 w-4" />
            Back
          </Button>
          {isGenerating && (
            <div className="flex items-center gap-3 ml-4 text-xs text-muted-foreground">
              <div className="flex items-center gap-1.5">
                <div className="h-2 w-2 rounded-full bg-primary animate-pulse" />
                <span>Generating</span>
              </div>
              <span className="tabular-nums">{formatTime(elapsedTime)}</span>
              <span className="text-border">|</span>
              <span className="tabular-nums">{wordCount.toLocaleString()} words</span>
            </div>
          )}
        </div>
        <div className="flex items-center gap-2">
          {isDone && (
            <>
              <span className="text-xs text-muted-foreground mr-2 tabular-nums" data-testid="text-word-count">
                {wordCount.toLocaleString()} words
              </span>
              <Button variant="outline" size="sm" onClick={handleExportPDF} disabled={isExporting !== null} data-testid="button-export-pdf" className="gap-2">
                {isExporting === "pdf" ? <Loader2 className="h-4 w-4 animate-spin" /> : <Download className="h-4 w-4" />}
                PDF
              </Button>
              <Button variant="outline" size="sm" onClick={handleExportPNG} disabled={isExporting !== null} data-testid="button-export-png" className="gap-2">
                {isExporting === "png" ? <Loader2 className="h-4 w-4 animate-spin" /> : <ImageIcon className="h-4 w-4" />}
                PNG
              </Button>
              <Button variant="outline" size="sm" onClick={() => window.print()} data-testid="button-print" className="gap-2">
                <Printer className="h-4 w-4" />
                Print
              </Button>
            </>
          )}
        </div>
      </header>

      <main className="flex-1 p-4 md:p-8 max-w-5xl mx-auto w-full">
        <Card className="shadow-xl border-border/30 bg-card/95 backdrop-blur-sm" ref={exportRef}>
          <CardHeader className="border-b border-border/30 bg-gradient-to-r from-primary/5 to-transparent px-8 py-6">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-4">
                <div className="p-2.5 rounded-xl bg-primary/10 ring-1 ring-primary/20">
                  <FileText className="h-6 w-6 text-primary" />
                </div>
                <div>
                  <CardTitle className="font-display text-xl tracking-tight" data-testid="text-summary-title">Executive Briefing</CardTitle>
                  <p className="text-sm text-muted-foreground mt-1">
                    AI-powered vCIO analysis — Microsoft 365 licensing optimization
                  </p>
                </div>
              </div>
              {isDone && (
                <div className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 rounded-full text-xs font-medium ring-1 ring-emerald-500/20">
                  <div className="h-1.5 w-1.5 rounded-full bg-emerald-500" />
                  Complete
                </div>
              )}
            </div>
          </CardHeader>
          <CardContent className="p-6 md:p-10">
            {!content && isGenerating && (
              <div className="flex flex-col items-center justify-center py-20 gap-5">
                <div className="relative">
                  <div className="absolute inset-0 rounded-full bg-primary/20 animate-ping" />
                  <Loader2 className="h-10 w-10 animate-spin text-primary relative" />
                </div>
                <div className="text-center">
                  <p className="font-medium text-foreground">Analyzing your Microsoft 365 environment...</p>
                  <p className="text-sm text-muted-foreground mt-1">Building comprehensive executive briefing</p>
                </div>
              </div>
            )}

            {!content && !isGenerating && !isDone && (
              <div className="flex flex-col items-center justify-center py-20 gap-5">
                <div className="p-4 rounded-full bg-muted/50">
                  <FileText className="h-8 w-8 text-muted-foreground" />
                </div>
                <div className="text-center">
                  <p className="font-medium text-foreground">No summary found for this report</p>
                  <p className="text-sm text-muted-foreground mt-1">Go back to the dashboard to generate one</p>
                </div>
                <Button variant="outline" onClick={() => navigate("/")} data-testid="button-go-back" className="gap-2 mt-2">
                  <ArrowLeft className="h-4 w-4" />
                  Back to Dashboard
                </Button>
              </div>
            )}

            {content && (
              <div
                ref={contentRef}
                className="max-w-none executive-report print:text-black"
                dangerouslySetInnerHTML={{ __html: renderMarkdown(content) }}
                data-testid="text-summary-content"
              />
            )}

            {isGenerating && content && (
              <div className="flex items-center gap-3 mt-8 pt-4 border-t border-border/30">
                <Loader2 className="h-4 w-4 animate-spin text-primary" />
                <span className="text-sm text-muted-foreground">
                  Generating analysis... {wordCount.toLocaleString()} words written ({formatTime(elapsedTime)})
                </span>
                <div className="flex-1" />
                <div className="flex gap-1">
                  {[0, 1, 2].map((i) => (
                    <div
                      key={i}
                      className="h-1.5 w-1.5 rounded-full bg-primary/60 animate-bounce"
                      style={{ animationDelay: `${i * 0.15}s` }}
                    />
                  ))}
                </div>
              </div>
            )}
          </CardContent>
        </Card>

        {isDone && (
          <div className="mt-6 text-center text-xs text-muted-foreground pb-2">
            Generated by Astra vCIO Advisory Engine — {new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" })}
          </div>
        )}
      </main>
      <footer className="py-4 text-center text-xs text-muted-foreground border-t border-border/30">
        &copy; 2026 Cavaridge, LLC. All rights reserved.
      </footer>
    </div>
  );
}
