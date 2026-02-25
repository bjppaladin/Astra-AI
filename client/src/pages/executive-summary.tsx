import { useState, useEffect, useRef } from "react";
import { useParams, useLocation } from "wouter";
import { ArrowLeft, FileText, Loader2, Printer } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { generateSummaryStream, fetchSummary } from "@/lib/api";

export default function ExecutiveSummaryPage() {
  const params = useParams<{ id: string }>();
  const [, navigate] = useLocation();
  const reportId = Number(params.id);

  const [content, setContent] = useState("");
  const [isGenerating, setIsGenerating] = useState(false);
  const [isDone, setIsDone] = useState(false);
  const contentRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    fetchSummary(reportId).then((existing) => {
      if (existing) {
        setContent(existing.content);
        setIsDone(true);
      } else {
        const stored = sessionStorage.getItem(`summary_payload_${reportId}`);
        if (stored) {
          const payload = JSON.parse(stored);
          setIsGenerating(true);
          generateSummaryStream(
            reportId,
            payload,
            (chunk) => setContent((prev) => prev + chunk),
            () => {
              setIsGenerating(false);
              setIsDone(true);
              sessionStorage.removeItem(`summary_payload_${reportId}`);
            },
            (err) => {
              setIsGenerating(false);
              setContent((prev) => prev + `\n\nError: ${err}`);
            }
          );
        }
      }
    });
  }, [reportId]);

  useEffect(() => {
    if (contentRef.current && isGenerating) {
      contentRef.current.scrollTop = contentRef.current.scrollHeight;
    }
  }, [content, isGenerating]);

  const renderMarkdown = (md: string) => {
    let html = md
      .replace(/^### (.+)$/gm, '<h3 class="text-lg font-semibold mt-6 mb-2 font-display">$1</h3>')
      .replace(/^## (.+)$/gm, '<h2 class="text-xl font-bold mt-8 mb-3 font-display text-primary">$1</h2>')
      .replace(/^# (.+)$/gm, '<h1 class="text-2xl font-bold mt-6 mb-4 font-display">$1</h1>')
      .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
      .replace(/\*(.+?)\*/g, '<em>$1</em>')
      .replace(/^\| (.+)$/gm, (match) => {
        const cells = match.split("|").filter(Boolean).map((c) => c.trim());
        const isHeader = cells.every((c) => /^[-:]+$/.test(c));
        if (isHeader) return '';
        return `<tr>${cells.map((c) => `<td class="border border-border/50 px-3 py-2 text-sm">${c}</td>`).join("")}</tr>`;
      })
      .replace(/^- (.+)$/gm, '<li class="ml-4 mb-1">$1</li>')
      .replace(/^(\d+)\. (.+)$/gm, '<li class="ml-4 mb-1"><span class="font-medium">$1.</span> $2</li>')
      .replace(/\n\n/g, '<br/><br/>');

    html = html.replace(/((?:<tr>.*?<\/tr>\s*)+)/gs, '<table class="w-full border-collapse my-4 rounded-lg overflow-hidden">$1</table>');
    html = html.replace(/((?:<li.*?<\/li>\s*)+)/gs, '<ul class="my-2">$1</ul>');

    return html;
  };

  return (
    <div className="min-h-screen bg-background flex flex-col font-sans text-foreground">
      <header className="sticky top-0 z-10 bg-card/80 backdrop-blur-md border-b border-border px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <Button variant="ghost" size="sm" onClick={() => navigate("/")} data-testid="button-back">
            <ArrowLeft className="h-4 w-4 mr-2" />
            Back to Dashboard
          </Button>
        </div>
        <div className="flex items-center gap-2">
          {isDone && (
            <Button variant="outline" size="sm" onClick={() => window.print()} data-testid="button-print">
              <Printer className="h-4 w-4 mr-2" />
              Print / PDF
            </Button>
          )}
        </div>
      </header>

      <main className="flex-1 p-8 max-w-4xl mx-auto w-full">
        <Card className="shadow-lg border-border/50">
          <CardHeader className="border-b border-border/50 bg-muted/10">
            <div className="flex items-center gap-3">
              <div className="p-2 rounded-lg bg-primary/10">
                <FileText className="h-5 w-5 text-primary" />
              </div>
              <div>
                <CardTitle className="font-display text-xl">Executive Summary</CardTitle>
                <p className="text-sm text-muted-foreground mt-1">
                  AI-generated vCIO analysis for C-Suite review
                </p>
              </div>
            </div>
          </CardHeader>
          <CardContent className="p-8">
            {!content && isGenerating && (
              <div className="flex flex-col items-center justify-center py-16 gap-4">
                <Loader2 className="h-8 w-8 animate-spin text-primary" />
                <p className="text-muted-foreground">Generating executive summary...</p>
              </div>
            )}

            {!content && !isGenerating && !isDone && (
              <div className="flex flex-col items-center justify-center py-16 gap-4">
                <p className="text-muted-foreground">No summary found for this report.</p>
                <Button variant="outline" onClick={() => navigate("/")} data-testid="button-go-back">
                  Go back to generate one
                </Button>
              </div>
            )}

            {content && (
              <div
                ref={contentRef}
                className="prose prose-sm max-w-none dark:prose-invert leading-relaxed"
                dangerouslySetInnerHTML={{ __html: renderMarkdown(content) }}
              />
            )}

            {isGenerating && content && (
              <div className="flex items-center gap-2 mt-6 text-sm text-muted-foreground">
                <Loader2 className="h-4 w-4 animate-spin" />
                Generating...
              </div>
            )}
          </CardContent>
        </Card>
      </main>
    </div>
  );
}
