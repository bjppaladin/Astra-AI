import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import { useLocation } from "wouter";
import {
  RefreshCcw,
  Download,
  Users,
  Database,
  CreditCard,
  Search,
  Filter,
  CheckCircle2,
  AlertCircle,
  Shield,
  TrendingDown,
  Scale,
  ArrowRight,
  FileText,
  Loader2,
  Upload,
  X,
  Info,
  LogIn,
  LogOut,
  Cloud,
  ChevronDown,
  ChevronUp,
  Package,
  Image,
  FileDown,
  HelpCircle,
  Newspaper,
  ExternalLink,
  Sparkles,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import * as XLSX from "xlsx";
import {
  saveReport,
  uploadUsersFile,
  uploadMailboxFile,
  getMicrosoftAuthStatus,
  getMicrosoftLoginUrl,
  disconnectMicrosoft,
  syncMicrosoftData,
  fetchSubscriptions,
  fetchGreeting,
  fetchNews,
} from "@/lib/api";
import { useToast } from "@/hooks/use-toast";
import { useAuth } from "@/hooks/use-auth";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { LICENSES } from "@/lib/license-data";

type UserRow = {
  id: string;
  displayName: string;
  upn: string;
  department: string;
  licenses: string[];
  usageGB: number;
  maxGB: number;
  cost: number;
  status: string;
};

const KEY_FEATURES = [
  "desktopApps", "mailboxSize", "oneDriveStorage", "teams",
  "conditionalAccess", "intune", "defenderOffice365", "dlp",
  "copilotEligible", "powerBI",
];
const KEY_FEATURE_LABELS: Record<string, string> = {
  desktopApps: "Desktop Office apps",
  mailboxSize: "Mailbox size",
  oneDriveStorage: "OneDrive storage",
  teams: "Microsoft Teams",
  conditionalAccess: "Conditional Access",
  intune: "Microsoft Intune",
  defenderOffice365: "Defender for Office 365",
  dlp: "Data Loss Prevention",
  copilotEligible: "Copilot eligible",
  powerBI: "Power BI",
};

function LicensePopoverContent({ licenseName }: { licenseName: string }) {
  const info = LICENSES.find(l => l.displayName === licenseName);
  if (!info) {
    return (
      <div className="space-y-2">
        <div className="font-semibold text-sm">{licenseName}</div>
        <p className="text-xs text-muted-foreground">No detailed feature data available for this license. Check the License Guide for more info.</p>
      </div>
    );
  }
  return (
    <div className="space-y-3">
      <div>
        <div className="font-semibold text-sm">{info.displayName}</div>
        <div className="flex items-center gap-2 mt-1">
          <Badge variant="secondary" className="text-[10px]">{info.category}</Badge>
          <span className="text-xs font-medium text-primary">${info.costPerMonth}/user/mo</span>
        </div>
        <p className="text-xs text-muted-foreground mt-1.5 leading-relaxed">{info.description}</p>
      </div>
      <div className="border-t border-border/50 pt-2 space-y-1.5">
        <div className="text-[10px] font-medium text-muted-foreground uppercase tracking-wider">Key Capabilities</div>
        {KEY_FEATURES.map(key => {
          const val = info.features[key];
          const label = KEY_FEATURE_LABELS[key];
          if (val === undefined) return null;
          const included = val === true || (typeof val === "string" && val !== "false");
          return (
            <div key={key} className="flex items-center justify-between text-xs">
              <span className={included ? "text-foreground" : "text-muted-foreground"}>{label}</span>
              {val === false ? (
                <X className="h-3.5 w-3.5 text-muted-foreground/50" />
              ) : val === true ? (
                <CheckCircle2 className="h-3.5 w-3.5 text-green-500" />
              ) : (
                <span className="text-[11px] font-medium text-primary">{val}</span>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

type Strategy = "current" | "security" | "cost" | "balanced" | "custom";

function useAnimatedNumber(value: number, duration = 600): string {
  const [display, setDisplay] = useState(value);
  const prevRef = useRef(value);
  useEffect(() => {
    const from = prevRef.current;
    const to = value;
    if (from === to) return;
    prevRef.current = to;
    const start = performance.now();
    const tick = (now: number) => {
      const t = Math.min((now - start) / duration, 1);
      const eased = 1 - Math.pow(1 - t, 3);
      setDisplay(from + (to - from) * eased);
      if (t < 1) requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  }, [value, duration]);
  return display === 0 && value === 0 ? "0" : display.toFixed(value % 1 === 0 ? 0 : 2);
}

const TUTORIAL_STEPS = [
  { target: "button-import", title: "Connect your data", body: "Start by signing in with Microsoft 365 for automatic data sync, or upload CSV/XLSX exports from the M365 Admin Center." },
  { target: "text-total-users", title: "Explore your licenses", body: "The dashboard shows all your licensed users, their assigned licenses with feature popovers, and mailbox usage. Click any license badge for details." },
  { target: "strategy-security", title: "Choose a strategy", body: "Pick from Security, Cost, Balanced, or Custom optimization. Each shows projected user impact and cost changes before you apply." },
  { target: "strategy-custom", title: "Customize rules", body: "In Custom mode, configure each rule's department scope, override thresholds, and see live impact counts. The Recommendations panel suggests data-driven actions." },
  { target: "button-generate-summary", title: "Generate insights", body: "Save your analysis and generate an AI-powered Executive Briefing — a board-ready report with per-user recommendations, cost projections, and strategic guidance." },
  { target: "nav-licenses", title: "Compare licenses", body: "Use the License Guide to compare up to 3 M365 licenses side-by-side across 8 feature categories, from core apps to security and AI." },
];

export default function Dashboard() {
  const [, navigate] = useLocation();
  const { toast } = useToast();
  const { user, logout } = useAuth();
  const [isSyncing, setIsSyncing] = useState(false);
  const [isGeneratingSummary, setIsGeneratingSummary] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [data, setData] = useState<UserRow[]>([]);
  const [dataSource, setDataSource] = useState<"none" | "uploaded" | "microsoft">("none");
  const [showUploadPanel, setShowUploadPanel] = useState(false);
  const [uploadedUserFile, setUploadedUserFile] = useState<string | null>(null);
  const [uploadedMailboxFile, setUploadedMailboxFile] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState("");
  const [showFilters, setShowFilters] = useState(false);
  const [filterDepartment, setFilterDepartment] = useState<string>("all");
  const [filterStatus, setFilterStatus] = useState<string>("all");
  const [filterModified, setFilterModified] = useState<string>("all");
  const [strategy, setStrategy] = useState<Strategy>("current");
  const [commitment, setCommitment] = useState<"monthly" | "annual">("annual");
  const userFileRef = useRef<HTMLInputElement>(null);
  const mailboxFileRef = useRef<HTMLInputElement>(null);

  const [msAuth, setMsAuth] = useState<{
    configured: boolean;
    connected: boolean;
    user?: { displayName: string; email: string };
    tenantId?: string;
  }>({ configured: false, connected: false });
  const [msLoading, setMsLoading] = useState(false);
  const [subscriptions, setSubscriptions] = useState<{
    skuId: string; skuPartNumber: string; displayName: string;
    costPerUser: number; enabled: number; consumed: number;
    available: number; capabilityStatus: string; appliesTo: string;
  }[]>([]);
  const [showSubscriptions, setShowSubscriptions] = useState(false);
  const [subsLoading, setSubsLoading] = useState(false);

  const [isExporting, setIsExporting] = useState<"pdf" | "png" | null>(null);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const dashboardRef = useRef<HTMLDivElement>(null);

  const [greeting, setGreeting] = useState<{ message: string; subtitle: string; loginCount: number; firstName: string } | null>(null);

  const [tutorialStep, setTutorialStep] = useState(-1);
  const [tutorialTooltipPos, setTutorialTooltipPos] = useState<{ top: number; left: number } | null>(null);

  const [newsItems, setNewsItems] = useState<{ title: string; link: string; date: string; summary: string }[]>([]);
  const [showNews, setShowNews] = useState(false);
  const [newsLoading, setNewsLoading] = useState(false);
  const [newsCachedAt, setNewsCachedAt] = useState<string | null>(null);

  const [strategyKey, setStrategyKey] = useState(0);

  type RuleScope = "all" | "security" | "custom";
  type ScopedRule = { enabled: boolean; scope: RuleScope; departments: string[]; threshold?: number };
  type CustomRulesState = {
    upgradeUnderprovisioned: ScopedRule;
    upgradeToE5: ScopedRule;
    upgradeBasicToStandard: ScopedRule;
    upgradeToBizPremium: ScopedRule;
    downgradeUnderutilizedE5: ScopedRule;
    downgradeOverprovisionedE3: ScopedRule;
    downgradeUnderutilizedBizPremium: ScopedRule;
    downgradeBizStandardToBasic: ScopedRule;
    removeUnusedAddons: boolean;
    consolidateOverlap: boolean;
    removeRedundantAddons: boolean;
    addCopilotPowerUsers: boolean;
    usageThreshold: number;
  };

  const [customRules, setCustomRules] = useState<CustomRulesState>({
    upgradeUnderprovisioned: { enabled: true, scope: "all", departments: [] },
    upgradeToE5: { enabled: false, scope: "security", departments: [] },
    upgradeBasicToStandard: { enabled: true, scope: "all", departments: [] },
    upgradeToBizPremium: { enabled: false, scope: "security", departments: [] },
    downgradeUnderutilizedE5: { enabled: true, scope: "all", departments: [], threshold: undefined },
    downgradeOverprovisionedE3: { enabled: false, scope: "all", departments: [], threshold: undefined },
    downgradeUnderutilizedBizPremium: { enabled: true, scope: "all", departments: [], threshold: undefined },
    downgradeBizStandardToBasic: { enabled: false, scope: "all", departments: [], threshold: undefined },
    removeUnusedAddons: true,
    consolidateOverlap: true,
    removeRedundantAddons: true,
    addCopilotPowerUsers: false,
    usageThreshold: 20,
  });

  const checkAuthStatus = useCallback(async () => {
    try {
      const status = await getMicrosoftAuthStatus();
      setMsAuth(status);
    } catch {}
  }, []);

  useEffect(() => {
    checkAuthStatus();

    const params = new URLSearchParams(window.location.search);
    if (params.get("auth_success")) {
      toast({ title: "Connected to Microsoft 365", description: "You can now sync your tenant data." });
      checkAuthStatus();
      window.history.replaceState({}, "", "/");
    }
    if (params.get("auth_error")) {
      toast({ title: "Authentication failed", description: decodeURIComponent(params.get("auth_error")!), variant: "destructive" });
      window.history.replaceState({}, "", "/");
    }
  }, []);

  useEffect(() => {
    if (msAuth.connected) {
      fetchGreeting().then(r => { if (r.greeting) setGreeting(r.greeting); }).catch(() => {});
    } else {
      setGreeting(null);
    }
  }, [msAuth.connected]);

  useEffect(() => {
    const completed = localStorage.getItem("astra_tutorial_completed");
    if (!completed && data.length === 0) {
      setTutorialStep(0);
    }
  }, []);

  useEffect(() => {
    if (tutorialStep < 0 || tutorialStep >= TUTORIAL_STEPS.length) {
      setTutorialTooltipPos(null);
      return;
    }
    const step = TUTORIAL_STEPS[tutorialStep];
    const el = document.querySelector(`[data-testid="${step.target}"]`);
    if (el) {
      el.scrollIntoView({ behavior: "smooth", block: "center" });
      setTimeout(() => {
        const rect = el.getBoundingClientRect();
        el.classList.add("tutorial-highlight");
        const top = rect.bottom + 12;
        const left = Math.max(16, Math.min(rect.left, window.innerWidth - 360));
        setTutorialTooltipPos({ top, left });
      }, 300);
    }
    return () => {
      const prevEl = document.querySelector(`[data-testid="${step.target}"]`);
      if (prevEl) prevEl.classList.remove("tutorial-highlight");
    };
  }, [tutorialStep]);

  const endTutorial = () => {
    const step = TUTORIAL_STEPS[tutorialStep];
    if (step) {
      const el = document.querySelector(`[data-testid="${step.target}"]`);
      if (el) el.classList.remove("tutorial-highlight");
    }
    setTutorialStep(-1);
    setTutorialTooltipPos(null);
    localStorage.setItem("astra_tutorial_completed", "true");
  };

  const nextTutorialStep = () => {
    const step = TUTORIAL_STEPS[tutorialStep];
    if (step) {
      const el = document.querySelector(`[data-testid="${step.target}"]`);
      if (el) el.classList.remove("tutorial-highlight");
    }
    if (tutorialStep >= TUTORIAL_STEPS.length - 1) {
      endTutorial();
    } else {
      setTutorialStep(tutorialStep + 1);
    }
  };

  useEffect(() => {
    if (!showExportMenu) return;
    const close = () => setShowExportMenu(false);
    const timer = setTimeout(() => document.addEventListener("click", close), 0);
    return () => { clearTimeout(timer); document.removeEventListener("click", close); };
  }, [showExportMenu]);

  const loadNews = useCallback(async () => {
    setNewsLoading(true);
    try {
      const result = await fetchNews();
      setNewsItems(result.items);
      if (result.cachedAt) setNewsCachedAt(result.cachedAt);
    } catch {
    } finally {
      setNewsLoading(false);
    }
  }, []);

  const resolveColor = (val: string): string => {
    const c = document.createElement("canvas");
    c.width = 1; c.height = 1;
    const ctx = c.getContext("2d");
    if (!ctx) return val;
    ctx.fillStyle = val;
    ctx.fillRect(0, 0, 1, 1);
    const [r, g, b, a] = ctx.getImageData(0, 0, 1, 1).data;
    return a < 255 ? `rgba(${r},${g},${b},${(a / 255).toFixed(3)})` : `rgb(${r},${g},${b})`;
  };

  const captureContent = useCallback(async () => {
    const el = dashboardRef.current;
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
        const props = ["color","background-color","border-color","border-top-color","border-bottom-color","border-left-color","border-right-color","font-family","font-size","font-weight","font-style","line-height","letter-spacing","text-align","text-decoration","text-transform","white-space","word-spacing","padding-top","padding-right","padding-bottom","padding-left","margin-top","margin-right","margin-bottom","margin-left","display","flex-direction","align-items","justify-content","gap","width","min-width","max-width","position","top","right","bottom","left","border-width","border-style","border-radius","border-top-width","border-right-width","border-bottom-width","border-left-width","border-top-style","border-right-style","border-bottom-style","border-left-style","opacity","visibility","vertical-align","list-style-type","table-layout","border-collapse","border-spacing","flex-grow","flex-shrink","flex-basis","flex-wrap","order"];
        for (const prop of props) {
          let val = computed.getPropertyValue(prop);
          if (val && val !== "initial" && val !== "inherit") {
            if (val.includes("oklab") || val.includes("oklch") || val.includes("color-mix") || val.includes("lab(") || val.includes("lch(")) val = resolveColor(val);
            target.style.setProperty(prop, val);
          }
        }
        target.style.boxShadow = "none";
        target.style.overflow = "visible";
        target.style.height = "auto";
        target.style.minHeight = "0";
        target.style.maxHeight = "none";
        for (let i = 0; i < source.children.length; i++) {
          if (source.children[i] instanceof HTMLElement && target.children[i] instanceof HTMLElement) flattenStyles(source.children[i] as HTMLElement, target.children[i] as HTMLElement);
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
      return await html2canvas(clone, { scale: 2, useCORS: true, backgroundColor: "#ffffff", logging: false, width: cloneWidth, height: cloneHeight, windowWidth: cloneWidth, windowHeight: cloneHeight });
    } finally {
      document.body.removeChild(iframe);
    }
  }, []);

  const handleExportPDF = useCallback(async () => {
    if (!dashboardRef.current) return;
    setIsExporting("pdf");
    setShowExportMenu(false);
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
        ctx.drawImage(fullCanvas, 0, sourceY, fullCanvas.width, sourceH, 0, 0, fullCanvas.width, sourceH);
        const pageImgData = pageCanvas.toDataURL("image/jpeg", 0.92);
        const imgH = sourceH * scaleFactor;
        pdf.addImage(pageImgData, "JPEG", margin, margin, contentWidth, imgH);
      }
      const dateStr = new Date().toISOString().split("T")[0];
      pdf.save(`Astra-M365-Insights-${strategy}-${dateStr}.pdf`);
      toast({ title: "PDF exported", description: `Saved as ${totalPages}-page PDF.` });
    } catch (err: any) {
      toast({ title: "Export failed", description: err.message, variant: "destructive" });
    } finally {
      setIsExporting(null);
    }
  }, [captureContent, toast, strategy]);

  const handleExportPNG = useCallback(async () => {
    if (!dashboardRef.current) return;
    setIsExporting("png");
    setShowExportMenu(false);
    try {
      const canvas = await captureContent();
      const link = document.createElement("a");
      const dateStr = new Date().toISOString().split("T")[0];
      link.download = `Astra-M365-Insights-${strategy}-${dateStr}.png`;
      link.href = canvas.toDataURL("image/png");
      link.click();
      toast({ title: "Image exported", description: "Dashboard exported as PNG." });
    } catch (err: any) {
      toast({ title: "Export failed", description: err.message, variant: "destructive" });
    } finally {
      setIsExporting(null);
    }
  }, [captureContent, toast, strategy]);

  const handleUserFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsUploading(true);
    try {
      const result = await uploadUsersFile(file);
      setData(result.users);
      setDataSource("uploaded");
      setUploadedUserFile(result.fileName);
      toast({
        title: "Users imported",
        description: `Found ${result.licensedUsers} licensed users out of ${result.totalParsed} rows.`,
      });
    } catch (err: any) {
      toast({ title: "Upload failed", description: err.message, variant: "destructive" });
    } finally {
      setIsUploading(false);
      if (userFileRef.current) userFileRef.current.value = "";
    }
  };

  const handleMailboxFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsUploading(true);
    try {
      const result = await uploadMailboxFile(file);
      setUploadedMailboxFile(result.fileName);

      let matchedCount = 0;
      setData((prev) =>
        prev.map((user) => {
          const mailbox = result.mailboxData[user.upn.toLowerCase()];
          if (mailbox) {
            matchedCount++;
            const ratio = mailbox.usageGB / mailbox.maxGB;
            return {
              ...user,
              usageGB: mailbox.usageGB,
              maxGB: mailbox.maxGB,
              status: ratio > 0.9 ? "Critical" : ratio > 0.7 ? "Warning" : "Active",
            };
          }
          return user;
        })
      );
      toast({
        title: "Mailbox data merged",
        description: `Matched ${matchedCount} of ${result.totalMailboxes} mailboxes to user records.`,
      });
    } catch (err: any) {
      toast({ title: "Upload failed", description: err.message, variant: "destructive" });
    } finally {
      setIsUploading(false);
      if (mailboxFileRef.current) mailboxFileRef.current.value = "";
    }
  };

  const handleClearUploads = () => {
    setData([]);
    setDataSource("none");
    setUploadedUserFile(null);
    setUploadedMailboxFile(null);
    setShowUploadPanel(false);
    setStrategy("current");
    toast({ title: "Data cleared", description: "Connect Microsoft 365 or upload reports to get started." });
  };

  const handleMicrosoftLogin = async () => {
    setMsLoading(true);
    try {
      const authUrl = await getMicrosoftLoginUrl();
      window.location.href = authUrl;
    } catch (err: any) {
      toast({ title: "Login failed", description: err.message, variant: "destructive" });
      setMsLoading(false);
    }
  };

  const handleMicrosoftDisconnect = async () => {
    try {
      await disconnectMicrosoft();
      setMsAuth({ configured: true, connected: false });
      if (dataSource === "microsoft") {
        setData([]);
        setDataSource("none");
      }
      toast({ title: "Disconnected", description: "Microsoft 365 account disconnected." });
    } catch (err: any) {
      toast({ title: "Error", description: err.message, variant: "destructive" });
    }
  };

  const handleMicrosoftSync = async () => {
    setIsSyncing(true);
    try {
      const result = await syncMicrosoftData();
      setData(result.users);
      setDataSource("microsoft");
      setShowUploadPanel(false);
      toast({
        title: "Data synced from Microsoft 365",
        description: `Loaded ${result.users.length} licensed users with mailbox data.`,
      });
    } catch (err: any) {
      toast({ title: "Sync failed", description: err.message, variant: "destructive" });
    } finally {
      setIsSyncing(false);
    }
  };

  const loadSubscriptions = useCallback(async () => {
    if (!msAuth.connected) return;
    setSubsLoading(true);
    try {
      const result = await fetchSubscriptions();
      setSubscriptions(result.subscriptions);
    } catch (err: any) {
      console.error("Failed to load subscriptions:", err.message);
    } finally {
      setSubsLoading(false);
    }
  }, [msAuth.connected]);

  useEffect(() => {
    if (msAuth.connected && subscriptions.length === 0) {
      loadSubscriptions();
    }
  }, [msAuth.connected]);

  const handleSync = () => {
    if (msAuth.connected) {
      handleMicrosoftSync();
    } else {
      setIsSyncing(true);
      setTimeout(() => {
        setData([...data]);
        setIsSyncing(false);
      }, 1000);
    }
  };

  const handleExportXlsx = () => {
    setShowExportMenu(false);
    const rows = optimizedData.map((u) => ({
      "Display Name": u.displayName,
      UPN: u.upn,
      Department: u.department,
      Licenses: u.licenses.join("; "),
      "Mailbox Usage (GB)": Number(u.usageGB.toFixed(1)),
      "Mailbox Max (GB)": u.maxGB,
      "Est. Monthly License Cost": Number(u.cost.toFixed(2)),
      "Commitment Type": commitment,
      Strategy: strategy,
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Combined Report");

    const fileName = `M365_Insights_${strategy}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  const LICENSE_COSTS: Record<string, number> = {
    "Microsoft 365 E5": 57, "Microsoft 365 E3": 36, "Office 365 E5": 38,
    "Office 365 E3": 23, "Office 365 E1": 10, "Microsoft 365 F1": 2.25,
    "Microsoft 365 F3": 8, "Office 365 F3": 4, "Microsoft 365 F5 Security": 12,
    "Microsoft 365 F5 Compliance": 12,
    "Microsoft 365 Business Premium": 22, "Microsoft 365 Business Standard": 12.50,
    "Microsoft 365 Business Basic": 6, "Microsoft 365 Apps for business": 12.50,
    "Microsoft 365 Apps for enterprise": 12, "Microsoft 365 Copilot": 30,
    "GitHub Copilot": 20, "Visio Plan 2": 15, "Visio Plan 1": 5,
    "Project Plan 5": 55, "Project Plan 3": 30, "Project Plan 1": 10,
    "Power BI Pro": 10, "Power BI Premium Per User": 20, "Power BI Free": 0,
    "Exchange Online Plan 1": 4, "Exchange Online Plan 2": 8,
    "Exchange Online Kiosk": 2, "Exchange Online Essentials": 2,
    "Exchange Online Protection": 0,
    "Defender for Office 365 P1": 2, "Defender for Office 365 P2": 5,
    "Defender for Endpoint P1": 3, "Defender for Endpoint P2": 5.20,
    "Defender for Business": 3, "Defender for Identity": 5.50,
    "Defender for Cloud Apps": 3.50,
    "Enterprise Mobility + Security E3": 11.60, "Enterprise Mobility + Security E5": 16.40,
    "Entra ID P1": 6, "Entra ID P2": 9,
    "Microsoft Intune Plan 1": 8,
    "Azure Information Protection P1": 2, "Azure Information Protection P2": 5,
    "Rights Management Adhoc": 0,
    "Teams Phone System": 8, "Teams Phone System Virtual User": 0,
    "Domestic Calling Plan": 12, "International Calling Plan": 24,
    "Domestic Calling Plan (120 min)": 0, "Audio Conferencing": 4,
    "Teams Rooms Standard": 15, "Teams Rooms Pro": 40,
    "OneDrive for Business P1": 5, "OneDrive for Business P2": 0,
    "SharePoint Online Plan 1": 5, "SharePoint Online Plan 2": 10,
    "Microsoft 365 E5 Security": 12, "Microsoft 365 E5 Compliance": 12,
    "Microsoft 365 F5 Security + Compliance": 12,
    "Windows 10/11 Enterprise E3": 7, "Windows 10/11 Enterprise E5": 11,
    "Windows Store for Business": 0,
    "Power Apps per user": 20, "Power Apps per app": 5,
    "Power Automate per user": 15,
    "Dynamics 365 Customer Voice": 0, "Dynamics 365 Sales Professional": 65,
    "Dynamics 365 Sales Enterprise": 95, "Dynamics 365 Plan": 115,
    "Dynamics 365 Team Members": 8,
    "Teams Exploratory": 0, "Microsoft Teams (Free)": 0, "Microsoft Teams Trial": 0,
    "Power Automate Free": 0, "Power Apps Trial": 0, "Microsoft Stream": 0,
    "Power Virtual Agents Trial": 0, "Communication Compliance": 0,
    "Power Automate RPA Attended": 0, "Microsoft Clipchamp": 0, "Windows Autopatch": 0,
  };

  const SUITE_LICENSES = new Set([
    "Microsoft 365 E5", "Microsoft 365 E3", "Office 365 E5", "Office 365 E3",
    "Office 365 E1", "Microsoft 365 F1", "Microsoft 365 F3", "Office 365 F3",
    "Microsoft 365 Business Premium", "Microsoft 365 Business Standard",
    "Microsoft 365 Business Basic", "Microsoft 365 Apps for enterprise",
    "Microsoft 365 Apps for business",
  ]);

  const SECURITY_DEPTS = new Set(["IT", "Engineering", "Compliance", "Security", "InfoSec"]);

  const sortLicenses = (licenses: string[]) => {
    const suites = licenses.filter(l => SUITE_LICENSES.has(l)).sort();
    const addons = licenses.filter(l => !SUITE_LICENSES.has(l)).sort();
    return [...suites, ...addons];
  };

  type UserRec = { licenses: string[]; cost: number; reasons: string[] };

  const computeCost = (licenses: string[]) => {
    return licenses.reduce((sum, l) => sum + (LICENSE_COSTS[l] ?? 0), 0);
  };

  const scopeMatchesDept = (rule: ScopedRule, dept: string): boolean => {
    if (rule.scope === "all") return true;
    if (rule.scope === "security") return SECURITY_DEPTS.has(dept);
    return rule.departments.includes(dept);
  };

  const analyzeUser = useCallback((user: UserRow, strat: Strategy, rules: CustomRulesState): UserRec => {
    let newLicenses = [...user.licenses];
    const reasons: string[] = [];
    const hasMailboxData = user.maxGB > 0;
    const usageRatio = hasMailboxData ? (user.usageGB / user.maxGB) * 100 : -1;
    const isSecurityDept = SECURITY_DEPTS.has(user.department);
    const globalThreshold = rules.usageThreshold;

    const getRulesForStrategy = (): CustomRulesState => {
      if (strat === "custom") return rules;
      const s = (enabled: boolean, scope: RuleScope = "all"): ScopedRule => ({ enabled, scope, departments: [] });
      if (strat === "security") return {
        upgradeUnderprovisioned: s(true), upgradeToE5: s(true, "security"),
        upgradeBasicToStandard: s(false), upgradeToBizPremium: s(true, "security"),
        downgradeUnderutilizedE5: s(false), downgradeOverprovisionedE3: s(false),
        downgradeUnderutilizedBizPremium: s(false), downgradeBizStandardToBasic: s(false),
        removeUnusedAddons: false, consolidateOverlap: true,
        removeRedundantAddons: true, addCopilotPowerUsers: true, usageThreshold: 10,
      };
      if (strat === "cost") return {
        upgradeUnderprovisioned: s(false), upgradeToE5: s(false),
        upgradeBasicToStandard: s(false), upgradeToBizPremium: s(false),
        downgradeUnderutilizedE5: s(true), downgradeOverprovisionedE3: s(true),
        downgradeUnderutilizedBizPremium: s(true), downgradeBizStandardToBasic: s(true),
        removeUnusedAddons: true, consolidateOverlap: true,
        removeRedundantAddons: true, addCopilotPowerUsers: false, usageThreshold: 30,
      };
      return {
        upgradeUnderprovisioned: s(true), upgradeToE5: s(false),
        upgradeBasicToStandard: s(true), upgradeToBizPremium: s(false),
        downgradeUnderutilizedE5: s(true), downgradeOverprovisionedE3: s(false),
        downgradeUnderutilizedBizPremium: s(true), downgradeBizStandardToBasic: s(false),
        removeUnusedAddons: true, consolidateOverlap: true,
        removeRedundantAddons: true, addCopilotPowerUsers: false, usageThreshold: 20,
      };
    };

    const r = getRulesForStrategy();
    const effThreshold = strat === "custom" ? globalThreshold : r.usageThreshold;
    const ruleThreshold = (rule: ScopedRule) => rule.threshold ?? effThreshold;

    if (r.upgradeUnderprovisioned.enabled && scopeMatchesDept(r.upgradeUnderprovisioned, user.department)) {
      if (newLicenses.includes("Office 365 E1") && hasMailboxData && usageRatio > 50) {
        newLicenses = newLicenses.filter(l => l !== "Office 365 E1");
        newLicenses.push("Microsoft 365 E3");
        reasons.push(`E1 at ${usageRatio.toFixed(0)}% capacity — missing MFA enforcement, DLP policies, and data retention controls that E3 provides. Reduces breach risk and supports compliance requirements.`);
      } else if (newLicenses.includes("Office 365 E1")) {
        if (strat === "security") {
          newLicenses = newLicenses.filter(l => l !== "Office 365 E1");
          newLicenses.push("Microsoft 365 E3");
          reasons.push(`E1 provides no endpoint management, conditional access, or data loss prevention — key gaps for regulatory compliance. E3 closes these gaps at $26/user/mo additional.`);
        }
      }
      if (newLicenses.includes("Microsoft 365 F1") && hasMailboxData && usageRatio > 30) {
        newLicenses = newLicenses.filter(l => l !== "Microsoft 365 F1");
        newLicenses.push("Microsoft 365 Business Basic");
        reasons.push(`F1 limits mailbox to 2GB, but this user is at ${usageRatio.toFixed(0)}% of quota — Business Basic provides a full 50GB mailbox and web Office apps for $3.75/mo more.`);
      }
    }

    if (r.upgradeBasicToStandard.enabled && scopeMatchesDept(r.upgradeBasicToStandard, user.department)) {
      if (newLicenses.includes("Microsoft 365 Business Basic") && hasMailboxData && usageRatio > 50) {
        newLicenses = newLicenses.filter(l => l !== "Microsoft 365 Business Basic");
        newLicenses.push("Microsoft 365 Business Standard");
        reasons.push(`High engagement user (${usageRatio.toFixed(0)}% mailbox) on Basic — Standard adds installable desktop Office apps and Bookings, enabling offline productivity and reducing shadow IT risk. $6.50/user/mo uplift.`);
      }
    }

    if (r.upgradeToE5.enabled && scopeMatchesDept(r.upgradeToE5, user.department)) {
      if (newLicenses.includes("Microsoft 365 E3")) {
        newLicenses = newLicenses.filter(l => l !== "Microsoft 365 E3");
        newLicenses.push("Microsoft 365 E5");
        reasons.push(`${user.department} ${isSecurityDept ? "handles sensitive data/systems" : "selected for E5 upgrade"} — E5 adds Defender for Office 365 P2, Cloud App Security, auto-investigation & response, and eDiscovery Premium. $21/user/mo uplift.`);
      }
    }

    if (r.upgradeToBizPremium.enabled && scopeMatchesDept(r.upgradeToBizPremium, user.department)) {
      if (newLicenses.includes("Microsoft 365 Business Standard")) {
        newLicenses = newLicenses.filter(l => l !== "Microsoft 365 Business Standard");
        newLicenses.push("Microsoft 365 Business Premium");
        reasons.push(`${user.department} ${isSecurityDept ? "requires" : "selected for"} device management and threat protection — Business Premium adds Intune MDM, Defender for Business, and Conditional Access policies. $9.50/user/mo uplift.`);
      }
      if (newLicenses.includes("Microsoft 365 Business Basic")) {
        newLicenses = newLicenses.filter(l => l !== "Microsoft 365 Business Basic");
        newLicenses.push("Microsoft 365 Business Premium");
        reasons.push(`${user.department} ${isSecurityDept ? "needs" : "selected for"} endpoint protection and identity governance — Business Premium adds Intune, Defender for Business, and Conditional Access. $16/user/mo uplift.`);
      }
    }

    if (r.downgradeUnderutilizedE5.enabled && hasMailboxData) {
      const th = ruleThreshold(r.downgradeUnderutilizedE5);
      const inScope = scopeMatchesDept(r.downgradeUnderutilizedE5, user.department);
      if (inScope && newLicenses.includes("Microsoft 365 E5") && usageRatio < th) {
        newLicenses = newLicenses.filter(l => l !== "Microsoft 365 E5");
        newLicenses.push("Microsoft 365 E3");
        reasons.push(`E5 in ${user.department} at only ${usageRatio.toFixed(0)}% utilization — advanced threat protection, Phone System, and audio conferencing are unused capabilities. E3 retains core security and compliance. Saves $21/user/mo.`);
      } else if (inScope && newLicenses.includes("Office 365 E5") && usageRatio < th) {
        newLicenses = newLicenses.filter(l => l !== "Office 365 E5");
        newLicenses.push("Office 365 E3");
        reasons.push(`Office 365 E5 in ${user.department} at ${usageRatio.toFixed(0)}% — premium analytics and advanced voice features unused. E3 retains full productivity suite. Saves $15/user/mo.`);
      }
    }

    if (r.downgradeOverprovisionedE3.enabled && hasMailboxData) {
      const th = ruleThreshold(r.downgradeOverprovisionedE3);
      if (scopeMatchesDept(r.downgradeOverprovisionedE3, user.department) && newLicenses.includes("Microsoft 365 E3") && usageRatio < th) {
        newLicenses = newLicenses.filter(l => l !== "Microsoft 365 E3");
        newLicenses.push("Microsoft 365 Business Standard");
        reasons.push(`E3 user at ${usageRatio.toFixed(0)}% in ${user.department} — E3's enterprise compliance and Windows Enterprise licensing are underutilized. Business Standard covers desktop apps, email, and Teams. Saves $23.50/user/mo.`);
      }
    }

    if (r.downgradeUnderutilizedBizPremium.enabled && hasMailboxData) {
      const th = ruleThreshold(r.downgradeUnderutilizedBizPremium);
      if (scopeMatchesDept(r.downgradeUnderutilizedBizPremium, user.department) && newLicenses.includes("Microsoft 365 Business Premium") && usageRatio < th) {
        newLicenses = newLicenses.filter(l => l !== "Microsoft 365 Business Premium");
        newLicenses.push("Microsoft 365 Business Standard");
        reasons.push(`Business Premium in ${user.department} at ${usageRatio.toFixed(0)}% — Intune and Defender capabilities are underutilized. Standard retains desktop apps and collaboration tools. Saves $9.50/user/mo.`);
      }
    }

    if (r.downgradeBizStandardToBasic.enabled && hasMailboxData) {
      const th = ruleThreshold(r.downgradeBizStandardToBasic);
      if (scopeMatchesDept(r.downgradeBizStandardToBasic, user.department) && newLicenses.includes("Microsoft 365 Business Standard") && usageRatio < th) {
        newLicenses = newLicenses.filter(l => l !== "Microsoft 365 Business Standard");
        newLicenses.push("Microsoft 365 Business Basic");
        reasons.push(`Standard at ${usageRatio.toFixed(0)}% in ${user.department} — desktop Office apps are likely unused at this activity level. Basic provides full web/mobile access to Outlook, Teams, and OneDrive. Saves $6.50/user/mo.`);
      }
    }

    if (r.removeUnusedAddons) {
      if (newLicenses.includes("Visio Plan 2") && !["Engineering", "Architecture", "Design", "PMO"].includes(user.department)) {
        newLicenses = newLicenses.filter(l => l !== "Visio Plan 2");
        reasons.push(`Visio Plan 2 ($15/mo) assigned to ${user.department} — diagramming tools are typically used by Engineering, Architecture, and PMO. Consider reassigning or removing to avoid license waste.`);
      }
      if (newLicenses.includes("Project Plan 3") && hasMailboxData && usageRatio < effThreshold && !["PMO", "IT", "Engineering"].includes(user.department)) {
        newLicenses = newLicenses.filter(l => l !== "Project Plan 3");
        reasons.push(`Project Plan 3 ($30/mo) with low activity in ${user.department} — no evidence of active project management usage. Planner (included in suite) covers basic task management needs.`);
      }
      if (newLicenses.includes("Project Plan 5") && !["PMO", "IT"].includes(user.department)) {
        newLicenses = newLicenses.filter(l => l !== "Project Plan 5");
        newLicenses.push("Project Plan 3");
        reasons.push(`Project Plan 5 ($55/mo) in ${user.department} — Plan 5's portfolio management, resource optimization, and demand management are enterprise PMO features. Plan 3 covers standard project scheduling. Saves $25/user/mo.`);
      }
      if (newLicenses.includes("Power BI Premium Per User") && !["Finance", "Analytics", "IT", "Engineering"].includes(user.department)) {
        newLicenses = newLicenses.filter(l => l !== "Power BI Premium Per User");
        newLicenses.push("Power BI Pro");
        reasons.push(`Power BI Premium Per User ($20/mo) in ${user.department} — Premium features like paginated reports, AI insights, and deployment pipelines are typically data-team needs. Pro covers standard reporting. Saves $10/user/mo.`);
      }
    }

    if (r.removeRedundantAddons) {
      const hasBizPremium = newLicenses.includes("Microsoft 365 Business Premium");
      const hasE5 = newLicenses.includes("Microsoft 365 E5");
      if ((hasBizPremium || hasE5) && newLicenses.includes("Defender for Office 365 P1")) {
        newLicenses = newLicenses.filter(l => l !== "Defender for Office 365 P1");
        reasons.push(`Defender for Office 365 P1 ($2/mo) is redundant — ${hasBizPremium ? "Business Premium includes Defender for Business with equivalent anti-phishing and safe attachments" : "E5 includes Defender P2 which supersedes P1"}. Direct cost savings with no capability loss.`);
      }
      if (hasE5 && newLicenses.includes("Defender for Office 365 P2")) {
        newLicenses = newLicenses.filter(l => l !== "Defender for Office 365 P2");
        reasons.push(`Defender for Office 365 P2 ($5/mo) is redundant — E5 already bundles full Defender P2 capabilities including auto-investigation and attack simulation training.`);
      }
      const hasAnySuite = newLicenses.some(l => SUITE_LICENSES.has(l));
      if (hasAnySuite && newLicenses.includes("OneDrive for Business P2")) {
        newLicenses = newLicenses.filter(l => l !== "OneDrive for Business P2");
        reasons.push(`OneDrive standalone license is redundant — suite license already includes OneDrive for Business with 1TB+ storage. Duplicate assignment creates unnecessary cost.`);
      }
      if (hasAnySuite && newLicenses.includes("OneDrive for Business P1")) {
        newLicenses = newLicenses.filter(l => l !== "OneDrive for Business P1");
        reasons.push(`OneDrive standalone ($5/mo) is redundant — suite license already includes OneDrive for Business. Removing eliminates duplicate spend.`);
      }
      const hasEMSE3orHigher = newLicenses.some(l => ["Enterprise Mobility + Security E3", "Enterprise Mobility + Security E5", "Microsoft 365 E3", "Microsoft 365 E5"].includes(l));
      if (hasEMSE3orHigher && newLicenses.includes("Entra ID P1")) {
        newLicenses = newLicenses.filter(l => l !== "Entra ID P1");
        reasons.push(`Entra ID P1 ($6/mo) is redundant — already bundled in the EMS or M365 E3+ suite. Conditional access and MFA are available without this standalone license.`);
      }
      if (hasEMSE3orHigher && newLicenses.includes("Microsoft Intune Plan 1")) {
        newLicenses = newLicenses.filter(l => l !== "Microsoft Intune Plan 1");
        reasons.push(`Intune Plan 1 ($8/mo) is redundant — device management is already included in the EMS or M365 E3+ suite. Removing eliminates double billing.`);
      }
    }

    if (r.consolidateOverlap) {
      const hasSuiteWithExchange = newLicenses.some(l => [
        "Microsoft 365 E3", "Microsoft 365 E5", "Office 365 E3", "Office 365 E5",
        "Microsoft 365 Business Basic", "Microsoft 365 Business Standard", "Microsoft 365 Business Premium",
      ].includes(l));
      if (hasSuiteWithExchange && newLicenses.includes("Exchange Online Plan 1")) {
        newLicenses = newLicenses.filter(l => l !== "Exchange Online Plan 1");
        reasons.push(`Exchange Online Plan 1 ($4/mo) is redundant — the suite license already provides Exchange Online with equal or greater mailbox capacity. Likely a leftover from migration.`);
      }
      if (hasSuiteWithExchange && newLicenses.includes("Exchange Online Plan 2")) {
        newLicenses = newLicenses.filter(l => l !== "Exchange Online Plan 2");
        reasons.push(`Exchange Online Plan 2 ($8/mo) is redundant — suite license includes full Exchange functionality. If 100GB mailbox is specifically needed, confirm before removing.`);
      }
      if (hasSuiteWithExchange && newLicenses.includes("Exchange Online Kiosk")) {
        newLicenses = newLicenses.filter(l => l !== "Exchange Online Kiosk");
        reasons.push(`Exchange Online Kiosk ($2/mo) is redundant — suite license already provides a full mailbox. Kiosk is likely a provisioning artifact.`);
      }
      const freeTrials = newLicenses.filter(l => [
        "Teams Exploratory", "Power Automate Free", "Power Apps Trial", "Microsoft Stream",
        "Microsoft Teams (Free)", "Microsoft Teams Trial", "Power Virtual Agents Trial",
        "Microsoft Clipchamp", "Windows Autopatch", "Power Automate RPA Attended",
        "Communication Compliance", "Rights Management Adhoc",
      ].includes(l));
      if (hasSuiteWithExchange && freeTrials.length > 0) {
        newLicenses = newLicenses.filter(l => !freeTrials.includes(l));
        reasons.push(`Removed ${freeTrials.length} trial/free license(s) — these are self-service signups that create license clutter. All functionality is covered by the assigned suite.`);
      }
    }

    if (r.addCopilotPowerUsers) {
      const powerUserDepts = new Set(["Engineering", "IT", "Design", "Analytics"]);
      if (powerUserDepts.has(user.department) && hasMailboxData && usageRatio > 50 && !newLicenses.includes("GitHub Copilot") && !newLicenses.includes("Microsoft 365 Copilot")) {
        if (user.department === "Engineering") {
          newLicenses.push("GitHub Copilot");
          reasons.push(`High-engagement Engineering user — GitHub Copilot delivers 30-55% developer productivity gains per industry benchmarks. ROI typically exceeds cost within the first month of adoption.`);
        } else {
          newLicenses.push("Microsoft 365 Copilot");
          reasons.push(`Power user in ${user.department} with ${usageRatio.toFixed(0)}% engagement — M365 Copilot accelerates document drafting, email triage, and meeting summaries. Best ROI for high-activity knowledge workers.`);
        }
      }
    }

    const finalLicenses = sortLicenses(newLicenses);
    const newCost = computeCost(finalLicenses);
    return { licenses: finalLicenses, cost: newCost, reasons };
  }, []);

  const analyzeAllUsers = useCallback((strat: Strategy, rules: typeof customRules) => {
    return data.map(user => {
      if (strat === "current") return { ...user, licenses: sortLicenses(user.licenses), reasons: [] as string[] };
      const rec = analyzeUser(user, strat, rules);
      return { ...user, ...rec };
    });
  }, [data, analyzeUser]);

  const optimizedData = useMemo(() => {
    return analyzeAllUsers(strategy, customRules);
  }, [strategy, customRules, analyzeAllUsers]);

  const departments = useMemo(() => {
    const depts = new Set(data.map(u => u.department));
    return Array.from(depts).sort();
  }, [data]);

  const filteredData = useMemo(() => {
    return optimizedData.filter(item => {
      const matchesSearch = searchTerm === "" ||
        item.displayName.toLowerCase().includes(searchTerm.toLowerCase()) ||
        item.upn.toLowerCase().includes(searchTerm.toLowerCase()) ||
        item.department.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesDept = filterDepartment === "all" || item.department === filterDepartment;
      const matchesStatus = filterStatus === "all" || item.status === filterStatus;
      if (filterModified !== "all" && strategy !== "current") {
        const orig = data.find(u => u.id === item.id);
        const isChanged = orig && JSON.stringify(sortLicenses(orig.licenses)) !== JSON.stringify(item.licenses);
        if (filterModified === "changed" && !isChanged) return false;
        if (filterModified === "unchanged" && isChanged) return false;
      }
      return matchesSearch && matchesDept && matchesStatus;
    });
  }, [optimizedData, searchTerm, filterDepartment, filterStatus, filterModified, strategy, data]);

  const activeFilterCount = [filterDepartment, filterStatus, filterModified].filter(f => f !== "all").length;

  const getStrategyStats = useCallback((strat: Strategy) => {
    const result = analyzeAllUsers(strat, customRules);
    const baseCost = data.reduce((a, c) => a + c.cost, 0);
    const newCost = result.reduce((a, c) => a + c.cost, 0);
    const affected = result.filter((u, i) => JSON.stringify(sortLicenses(data[i]?.licenses || [])) !== JSON.stringify(u.licenses)).length;
    const upgrades = result.reduce((a, u) => a + u.reasons.filter(r => r.toLowerCase().includes("upgrade")).length, 0);
    const downgrades = result.reduce((a, u) => a + u.reasons.filter(r => r.toLowerCase().includes("downgrade") || r.toLowerCase().includes("remove") || r.toLowerCase().includes("redundant")).length, 0);
    return { baseCost, newCost, delta: newCost - baseCost, affected, upgrades, downgrades };
  }, [analyzeAllUsers, customRules, data]);

  const costForStrategy = useCallback((strat: Strategy) => {
    return getStrategyStats(strat).newCost;
  }, [getStrategyStats]);

  const baseTotalCost = data.reduce((acc, curr) => acc + curr.cost, 0);
  const projectedTotalCost = optimizedData.reduce((acc, curr) => acc + curr.cost, 0);
  const costDiff = projectedTotalCost - baseTotalCost;

  const commitmentMultiplier = commitment === "annual" ? 0.85 : 1;
  const totalCost = projectedTotalCost * commitmentMultiplier;

  const baseTotalCostCommitted = baseTotalCost * commitmentMultiplier;
  const costDiffCommitted = totalCost - baseTotalCostCommitted;

  const totalUsers = optimizedData.length;
  const totalStorage = optimizedData.reduce((acc, curr) => acc + curr.usageGB, 0);

  const animatedUsers = useAnimatedNumber(totalUsers);
  const animatedStorage = useAnimatedNumber(totalStorage);
  const animatedCost = useAnimatedNumber(totalCost);

  const handleStrategyChange = (key: Strategy) => {
    setStrategy(key);
    setStrategyKey(prev => prev + 1);
  };

  const handleGenerateSummary = async () => {
    if (data.length === 0) return;
    setIsGeneratingSummary(true);
    try {
      const mul = commitmentMultiplier;
      const costCurrent = costForStrategy("current") * mul;
      const costSecurity = costForStrategy("security") * mul;
      const costSaving = costForStrategy("cost") * mul;
      const costBalanced = costForStrategy("balanced") * mul;
      const costCustom = costForStrategy("custom") * mul;

      const report = await saveReport({
        name: `M365 Report - ${new Date().toLocaleDateString()}`,
        strategy,
        commitment,
        userData: data,
        customRules: strategy === "custom" ? customRules : undefined,
      });

      const payload = {
        costCurrent,
        costSecurity,
        costSaving,
        costBalanced,
        costCustom,
        commitment,
        userData: data,
      };

      sessionStorage.setItem(`summary_payload_${report.id}`, JSON.stringify(payload));
      navigate(`/report/${report.id}/summary`);
    } catch (err) {
      console.error("Failed to generate summary:", err);
    } finally {
      setIsGeneratingSummary(false);
    }
  };

  return (
    <div className="min-h-screen bg-background flex flex-col font-sans text-foreground">
      {/* Top Navigation */}
      <header className="sticky top-0 z-10 bg-card/80 backdrop-blur-md border-b border-border px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 cursor-pointer" onClick={() => navigate("/")} data-testid="link-home">
            <div className="h-8 w-8 rounded-md bg-primary flex items-center justify-center text-primary-foreground font-bold">
              A
            </div>
            <h1 className="text-xl font-semibold tracking-tight">Astra</h1>
          </div>
          <nav className="hidden sm:flex items-center gap-1 ml-2">
            <Button variant="ghost" size="sm" className="text-foreground font-medium bg-muted/50" data-testid="nav-dashboard">Dashboard</Button>
            <Button variant="ghost" size="sm" className="text-muted-foreground hover:text-foreground" onClick={() => navigate("/licenses")} data-testid="nav-licenses">License Guide</Button>
          </nav>
        </div>
        <div className="flex items-center gap-3">
          {dataSource === "uploaded" && (
            <Badge variant="outline" className="text-xs font-normal text-green-600 border-green-300">
              Imported Data
              <button onClick={handleClearUploads} className="ml-1.5 hover:text-green-800" data-testid="button-clear-data">
                <X className="h-3 w-3 inline" />
              </button>
            </Badge>
          )}
          {dataSource === "microsoft" && (
            <Badge variant="outline" className="text-xs font-normal text-blue-600 border-blue-300">
              <Cloud className="h-3 w-3 mr-1 inline" />
              Live M365 Data
              <button onClick={handleClearUploads} className="ml-1.5 hover:text-blue-800" data-testid="button-clear-microsoft">
                <X className="h-3 w-3 inline" />
              </button>
            </Badge>
          )}
          {msAuth.connected && (
            <Button
              variant="outline"
              size="sm"
              className="gap-2"
              onClick={handleMicrosoftSync}
              disabled={isSyncing}
              data-testid="button-sync-m365"
            >
              {isSyncing ? <Loader2 className="h-4 w-4 animate-spin" /> : <RefreshCcw className="h-4 w-4" />}
              Sync M365
            </Button>
          )}
          <Button 
            variant="outline" 
            size="sm" 
            onClick={() => setShowUploadPanel(!showUploadPanel)}
            className="gap-2"
            data-testid="button-import"
          >
            <Upload className="h-4 w-4" />
            Import Data
          </Button>
          <div className="relative">
            <Button size="sm" className="gap-2" onClick={() => setShowExportMenu(!showExportMenu)} data-testid="button-export">
              {isExporting ? <Loader2 className="h-4 w-4 animate-spin" /> : <Download className="h-4 w-4" />}
              Export
              <ChevronDown className="h-3 w-3" />
            </Button>
            {showExportMenu && (
              <div className="absolute right-0 top-full mt-1 w-48 bg-card border border-border rounded-lg shadow-lg py-1 z-50 animate-in fade-in slide-in-from-top-1 duration-150">
                <button className="w-full text-left px-3 py-2 text-sm hover:bg-muted/50 flex items-center gap-2" onClick={handleExportXlsx} data-testid="export-xlsx">
                  <Download className="h-4 w-4" />
                  Export as XLSX
                </button>
                <button className="w-full text-left px-3 py-2 text-sm hover:bg-muted/50 flex items-center gap-2" onClick={handleExportPDF} disabled={!!isExporting} data-testid="export-pdf">
                  <FileDown className="h-4 w-4" />
                  Export as PDF
                </button>
                <button className="w-full text-left px-3 py-2 text-sm hover:bg-muted/50 flex items-center gap-2" onClick={handleExportPNG} disabled={!!isExporting} data-testid="export-png">
                  <Image className="h-4 w-4" />
                  Export as PNG
                </button>
              </div>
            )}
          </div>
          <Button 
            size="sm" 
            className="gap-2 bg-primary"
            onClick={handleGenerateSummary}
            disabled={isGeneratingSummary || data.length === 0}
            data-testid="button-generate-summary"
          >
            {isGeneratingSummary ? <Loader2 className="h-4 w-4 animate-spin" /> : <FileText className="h-4 w-4" />}
            Executive Summary
          </Button>
          <Button
            variant="ghost"
            size="icon"
            className="h-9 w-9"
            onClick={() => { setTutorialStep(0); }}
            title="Start tutorial"
            data-testid="button-tutorial"
          >
            <HelpCircle className="h-4 w-4" />
          </Button>
          {user && (
            <div className="flex items-center gap-2 ml-1 pl-3 border-l border-border/50">
              {user.profileImageUrl && (
                <img src={user.profileImageUrl} alt="" className="h-7 w-7 rounded-full" />
              )}
              <span className="text-sm font-medium hidden md:inline" data-testid="text-user-name">
                {user.firstName || user.email || "User"}
              </span>
              <Button variant="ghost" size="sm" className="text-xs text-muted-foreground" onClick={() => logout()} data-testid="button-logout">
                <LogOut className="h-3.5 w-3.5 mr-1" />
                Sign Out
              </Button>
            </div>
          )}
        </div>
      </header>

      {/* Import Panel */}
      {showUploadPanel && (
        <div className="border-b border-border bg-muted/30 px-6 py-5 animate-in slide-in-from-top-2 duration-300">
          <div className="max-w-7xl mx-auto">
            <div className="flex items-start justify-between mb-4">
              <div>
                <h3 className="text-sm font-semibold flex items-center gap-2">
                  <Database className="h-4 w-4" />
                  Import Microsoft 365 Data
                </h3>
                <p className="text-xs text-muted-foreground mt-1">
                  Sign in with Microsoft to pull data automatically, or upload exported CSV/XLSX files.
                </p>
              </div>
              <Button variant="ghost" size="sm" onClick={() => setShowUploadPanel(false)} className="h-7 w-7 p-0" data-testid="button-close-upload">
                <X className="h-4 w-4" />
              </Button>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div className="border border-border rounded-lg p-4 bg-background/50 flex flex-col">
                <div className="flex items-center gap-2 mb-2">
                  <Cloud className="h-4 w-4 text-blue-600" />
                  <div className="text-sm font-medium">Microsoft 365 Sign-In</div>
                  <Badge variant="secondary" className="text-[10px] ml-auto">Recommended</Badge>
                </div>
                {msAuth.connected ? (
                  <div className="space-y-2">
                    <div className="text-xs bg-blue-50 dark:bg-blue-950 border border-blue-200 dark:border-blue-800 rounded-md p-2 flex items-center gap-2">
                      <CheckCircle2 className="h-3.5 w-3.5 text-blue-600 shrink-0" />
                      <div className="min-w-0">
                        <div className="font-medium text-blue-800 dark:text-blue-200">{msAuth.user?.displayName}</div>
                        <div className="text-blue-600 dark:text-blue-400 truncate">{msAuth.user?.email}</div>
                        {msAuth.tenantId && (
                          <div className="text-blue-500 dark:text-blue-500 truncate">Tenant: {msAuth.tenantId}</div>
                        )}
                      </div>
                    </div>
                    <div className="flex gap-2">
                      <Button
                        size="sm"
                        className="gap-2 flex-1"
                        onClick={handleMicrosoftSync}
                        disabled={isSyncing}
                        data-testid="button-sync-panel"
                      >
                        {isSyncing ? <Loader2 className="h-4 w-4 animate-spin" /> : <RefreshCcw className="h-4 w-4" />}
                        Sync Data
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        className="gap-2"
                        onClick={handleMicrosoftDisconnect}
                        data-testid="button-disconnect"
                      >
                        <LogOut className="h-4 w-4" />
                      </Button>
                    </div>
                  </div>
                ) : msAuth.configured ? (
                  <div className="space-y-2">
                    <p className="text-xs text-muted-foreground flex-1">
                      Click below to sign in with your Microsoft 365 account. You'll be taken to the Microsoft consent screen to authorize access.
                    </p>
                    <Button
                      size="sm"
                      className="gap-2 w-full bg-[#0078d4] hover:bg-[#106ebe] text-white"
                      onClick={handleMicrosoftLogin}
                      disabled={msLoading}
                      data-testid="button-microsoft-login"
                    >
                      {msLoading ? <Loader2 className="h-4 w-4 animate-spin" /> : <LogIn className="h-4 w-4" />}
                      Sign in with Microsoft
                    </Button>
                  </div>
                ) : (
                  <div className="text-xs text-muted-foreground bg-muted rounded-md p-3">
                    <Info className="h-3 w-3 inline mr-1" />
                    Microsoft sign-in is not yet available. Use CSV/XLSX upload instead.
                  </div>
                )}
              </div>
              <div className="border border-dashed border-border rounded-lg p-4 bg-background/50">
                <div className="flex items-center justify-between mb-2">
                  <div className="text-sm font-medium">Active Users Report</div>
                  {uploadedUserFile && (
                    <Badge variant="outline" className="text-xs text-green-600 border-green-300">{uploadedUserFile}</Badge>
                  )}
                </div>
                <p className="text-xs text-muted-foreground mb-3">
                  M365 Admin Center &rarr; Reports &rarr; Usage &rarr; Active Users &rarr; Export
                </p>
                <input
                  ref={userFileRef}
                  type="file"
                  accept=".csv,.xlsx,.xls"
                  onChange={handleUserFileUpload}
                  className="hidden"
                  data-testid="input-user-file"
                />
                <Button
                  variant="outline"
                  size="sm"
                  className="gap-2 w-full"
                  onClick={() => userFileRef.current?.click()}
                  disabled={isUploading}
                  data-testid="button-upload-users"
                >
                  {isUploading ? <Loader2 className="h-4 w-4 animate-spin" /> : <Upload className="h-4 w-4" />}
                  {uploadedUserFile ? "Replace file" : "Upload Active Users CSV/XLSX"}
                </Button>
              </div>
              <div className="border border-dashed border-border rounded-lg p-4 bg-background/50">
                <div className="flex items-center justify-between mb-2">
                  <div className="text-sm font-medium">Mailbox Usage Report</div>
                  {uploadedMailboxFile && (
                    <Badge variant="outline" className="text-xs text-green-600 border-green-300">{uploadedMailboxFile}</Badge>
                  )}
                </div>
                <p className="text-xs text-muted-foreground mb-3">
                  M365 Admin Center &rarr; Reports &rarr; Usage &rarr; Email activity &rarr; Export
                </p>
                <input
                  ref={mailboxFileRef}
                  type="file"
                  accept=".csv,.xlsx,.xls"
                  onChange={handleMailboxFileUpload}
                  className="hidden"
                  data-testid="input-mailbox-file"
                />
                <Button
                  variant="outline"
                  size="sm"
                  className="gap-2 w-full"
                  onClick={() => mailboxFileRef.current?.click()}
                  disabled={isUploading || dataSource === "none"}
                  data-testid="button-upload-mailbox"
                >
                  {isUploading ? <Loader2 className="h-4 w-4 animate-spin" /> : <Upload className="h-4 w-4" />}
                  {uploadedMailboxFile ? "Replace file" : "Upload Mailbox Usage CSV/XLSX"}
                </Button>
                {dataSource === "none" && (
                  <p className="text-xs text-muted-foreground mt-2 flex items-center gap-1">
                    <Info className="h-3 w-3" />
                    Upload Active Users first
                  </p>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Main Content */}
      <main className="flex-1 p-8 max-w-7xl mx-auto w-full space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500" ref={dashboardRef}>
        
        <div className="flex flex-col gap-4">
          <div className="flex flex-col gap-2">
            <div className="flex items-center justify-between">
              <div>
                <h2 className="text-3xl font-display font-semibold">M365 Insights</h2>
                <p className="text-muted-foreground">Automated merge of Active Users and Mailbox Usage reports with actionable insights.</p>
              </div>
              {greeting && (
                <div className="text-right animate-in fade-in slide-in-from-right-4 duration-500">
                  <div className="flex items-center gap-2 justify-end">
                    <Sparkles className="h-4 w-4 text-primary" />
                    <span className="text-lg font-semibold font-display">{greeting.message}</span>
                  </div>
                  <p className="text-sm text-muted-foreground">{greeting.subtitle}</p>
                </div>
              )}
            </div>
          </div>

          {data.length === 0 && !isSyncing ? (
            <Card className="border-dashed border-2 border-border/60 bg-muted/5 shadow-none">
              <CardContent className="py-12 flex flex-col items-center text-center">
                <div className="w-16 h-16 rounded-2xl bg-primary/10 flex items-center justify-center mb-6">
                  <Cloud className="h-8 w-8 text-primary" />
                </div>
                <h3 className="text-xl font-semibold mb-2">Connect your Microsoft 365 tenant</h3>
                <p className="text-muted-foreground max-w-md mb-8">
                  Astra analyzes your real licensing and mailbox data to find savings and security gaps. Get started in under a minute.
                </p>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6 w-full max-w-2xl">
                  <div className="flex flex-col items-center p-6 rounded-xl border border-border/60 bg-card">
                    <div className="w-10 h-10 rounded-lg bg-blue-500/10 flex items-center justify-center mb-3">
                      <Cloud className="h-5 w-5 text-blue-600" />
                    </div>
                    <div className="font-medium mb-1">Sign in with Microsoft</div>
                    <p className="text-xs text-muted-foreground mb-4 text-center">Recommended. One-click sign-in pulls users, licenses, and mailbox usage automatically via Microsoft Graph.</p>
                    {msAuth.connected ? (
                      <Button size="sm" className="gap-2 w-full" onClick={handleMicrosoftSync} disabled={isSyncing} data-testid="button-onboard-sync">
                        {isSyncing ? <Loader2 className="h-4 w-4 animate-spin" /> : <RefreshCcw className="h-4 w-4" />}
                        Sync Tenant Data
                      </Button>
                    ) : (
                      <Button size="sm" className="gap-2 w-full" onClick={handleMicrosoftLogin} disabled={msLoading} data-testid="button-onboard-login">
                        {msLoading ? <Loader2 className="h-4 w-4 animate-spin" /> : <LogIn className="h-4 w-4" />}
                        Sign in with Microsoft
                      </Button>
                    )}
                  </div>
                  <div className="flex flex-col items-center p-6 rounded-xl border border-border/60 bg-card">
                    <div className="w-10 h-10 rounded-lg bg-green-500/10 flex items-center justify-center mb-3">
                      <Upload className="h-5 w-5 text-green-600" />
                    </div>
                    <div className="font-medium mb-1">Upload CSV / XLSX</div>
                    <p className="text-xs text-muted-foreground mb-4 text-center">Export from M365 Admin Center → Reports → Usage → Active Users, then upload the file here.</p>
                    <Button size="sm" variant="outline" className="gap-2 w-full" onClick={() => { setShowUploadPanel(true); userFileRef.current?.click(); }} data-testid="button-onboard-upload">
                      <Upload className="h-4 w-4" />
                      Upload Active Users Report
                    </Button>
                  </div>
                </div>
                <div className="mt-6 flex items-center gap-4 text-xs text-muted-foreground">
                  <div className="flex items-center gap-1.5">
                    <Shield className="h-3.5 w-3.5" />
                    Data stays in your session
                  </div>
                  <div className="flex items-center gap-1.5">
                    <CheckCircle2 className="h-3.5 w-3.5" />
                    Read-only access
                  </div>
                  <div className="flex items-center gap-1.5">
                    <Users className="h-3.5 w-3.5" />
                    Works with any M365 tenant
                  </div>
                </div>
              </CardContent>
            </Card>
          ) : (
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="flex items-center gap-2">
              <Badge variant="secondary" className="font-normal">Billing basis</Badge>
              <div className="w-full sm:w-[240px]">
                <Select value={commitment} onValueChange={(v) => setCommitment(v as any)}>
                  <SelectTrigger className="bg-card" data-testid="select-commitment">
                    <SelectValue placeholder="Select commitment" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="monthly">Monthly commitment</SelectItem>
                    <SelectItem value="annual">Annual commitment (est. discount)</SelectItem>
                  </SelectContent>
                </Select>
              </div>
            </div>

            <div className="text-xs text-muted-foreground max-w-xl">
              Costs shown are a prototype estimate. Annual commitment applies an estimated per-month discount for comparison.
            </div>
          </div>
          )}
        </div>

        {/* KPIs */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <Card className="shadow-sm border-border/50 animate-stagger-1">
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium text-muted-foreground">Total Active Users</CardTitle>
              <Users className="h-4 w-4 text-muted-foreground" />
            </CardHeader>
            <CardContent>
              {isSyncing ? (
                <Skeleton className="h-8 w-20" />
              ) : (
                <div className="text-3xl font-bold font-display animate-count-up" data-testid="text-total-users">{animatedUsers}</div>
              )}
            </CardContent>
          </Card>
          
          <Card className="shadow-sm border-border/50 animate-stagger-2">
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium text-muted-foreground">Total Mailbox Usage</CardTitle>
              <Database className="h-4 w-4 text-muted-foreground" />
            </CardHeader>
            <CardContent>
               {isSyncing ? (
                <Skeleton className="h-8 w-24" />
              ) : (
                <div className="text-3xl font-bold font-display animate-count-up" data-testid="text-total-storage">{animatedStorage} GB</div>
              )}
            </CardContent>
          </Card>

          <Card className={`shadow-sm border-border/50 animate-stagger-3 transition-colors ${strategy !== 'current' ? 'bg-primary/5 border-primary/20' : ''}`}>
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium text-muted-foreground">Est. Monthly License Cost</CardTitle>
              <CreditCard className={`h-4 w-4 ${strategy !== 'current' ? 'text-primary' : 'text-muted-foreground'}`} />
            </CardHeader>
            <CardContent>
               {isSyncing ? (
                <Skeleton className="h-8 w-28" />
              ) : (
                <div className="flex items-end gap-3" data-testid="container-cost-metric">
                  <div className={`text-3xl font-bold font-display ${strategy !== 'current' ? 'text-primary' : ''}`}>
                    ${animatedCost}
                  </div>
                  <div className="flex flex-col items-end gap-0.5 mb-1">
                    <div className="text-xs text-muted-foreground">
                      {commitment === 'annual' ? 'Annual commitment (est.)' : 'Monthly commitment'}
                    </div>
                    {strategy !== 'current' && (
                      <div className={`text-sm font-medium ${costDiffCommitted > 0 ? 'text-amber-500' : 'text-green-500'}`}>
                        {costDiffCommitted > 0 ? '+' : ''}{costDiffCommitted.toFixed(2)} /mo
                      </div>
                    )}
                  </div>
                </div>
              )}
            </CardContent>
          </Card>
        </div>

        {/* Tenant Subscriptions */}
        {(subscriptions.length > 0 || msAuth.connected) && (
          <Card className="border-border/50 shadow-sm">
            <CardHeader
              className="cursor-pointer py-3 px-6"
              onClick={() => setShowSubscriptions(!showSubscriptions)}
            >
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="p-1.5 rounded-md bg-primary/10">
                    <Package className="h-4 w-4 text-primary" />
                  </div>
                  <div>
                    <CardTitle className="text-base">Tenant Subscriptions</CardTitle>
                    <CardDescription className="text-xs">
                      {subscriptions.length > 0
                        ? `${subscriptions.length} subscription${subscriptions.length !== 1 ? 's' : ''} · ${subscriptions.reduce((a, s) => a + s.consumed, 0)} total assigned licenses`
                        : "Connect to Microsoft 365 to view subscriptions"}
                    </CardDescription>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  {msAuth.connected && (
                    <Button
                      variant="ghost"
                      size="sm"
                      className="h-7 text-xs"
                      onClick={(e) => { e.stopPropagation(); loadSubscriptions(); }}
                      disabled={subsLoading}
                      data-testid="button-refresh-subs"
                    >
                      <RefreshCcw className={`h-3 w-3 mr-1 ${subsLoading ? 'animate-spin' : ''}`} />
                      Refresh
                    </Button>
                  )}
                  {showSubscriptions ? <ChevronUp className="h-4 w-4 text-muted-foreground" /> : <ChevronDown className="h-4 w-4 text-muted-foreground" />}
                </div>
              </div>
            </CardHeader>
            {showSubscriptions && (
              <CardContent className="pt-0 px-6 pb-4">
                {subsLoading ? (
                  <div className="space-y-2">
                    {Array.from({ length: 3 }).map((_, i) => (
                      <Skeleton key={i} className="h-10 w-full" />
                    ))}
                  </div>
                ) : subscriptions.length === 0 ? (
                  <div className="text-sm text-muted-foreground text-center py-6">
                    {msAuth.connected ? "No subscriptions found. Click Refresh to try again." : "Sign in with Microsoft 365 to load tenant subscriptions."}
                  </div>
                ) : (
                  <div className="overflow-x-auto">
                    <Table>
                      <TableHeader className="bg-muted/30">
                        <TableRow>
                          <TableHead>Subscription</TableHead>
                          <TableHead className="text-center">Status</TableHead>
                          <TableHead className="text-right">Purchased</TableHead>
                          <TableHead className="text-right">Assigned</TableHead>
                          <TableHead className="text-right">Available</TableHead>
                          <TableHead className="text-right">Utilization</TableHead>
                          <TableHead className="text-right">Cost/User</TableHead>
                          <TableHead className="text-right">Monthly Spend</TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {subscriptions
                          .sort((a, b) => (b.consumed * b.costPerUser) - (a.consumed * a.costPerUser))
                          .map((sub) => {
                            const utilization = sub.enabled > 0 ? (sub.consumed / sub.enabled) * 100 : 0;
                            const monthlySpend = sub.consumed * sub.costPerUser;
                            return (
                              <TableRow key={sub.skuId} data-testid={`row-sub-${sub.skuId}`}>
                                <TableCell>
                                  <div className="font-medium text-sm">{sub.displayName}</div>
                                  <div className="text-xs text-muted-foreground">{sub.skuPartNumber}</div>
                                </TableCell>
                                <TableCell className="text-center">
                                  <Badge
                                    variant="outline"
                                    className={`text-xs ${sub.capabilityStatus === 'Enabled' ? 'border-green-500/30 text-green-600 bg-green-500/5' : sub.capabilityStatus === 'Warning' ? 'border-amber-500/30 text-amber-600 bg-amber-500/5' : 'border-red-500/30 text-red-600 bg-red-500/5'}`}
                                  >
                                    {sub.capabilityStatus}
                                  </Badge>
                                </TableCell>
                                <TableCell className="text-right font-medium">{sub.enabled.toLocaleString()}</TableCell>
                                <TableCell className="text-right">{sub.consumed.toLocaleString()}</TableCell>
                                <TableCell className="text-right">
                                  <span className={sub.available <= 0 ? 'text-destructive font-medium' : sub.available <= 5 ? 'text-amber-600' : ''}>
                                    {sub.available.toLocaleString()}
                                  </span>
                                </TableCell>
                                <TableCell className="text-right">
                                  <div className="flex items-center justify-end gap-2">
                                    <div className="w-16 h-1.5 bg-secondary rounded-full overflow-hidden">
                                      <div
                                        className={`h-full rounded-full ${utilization > 90 ? 'bg-destructive' : utilization > 70 ? 'bg-amber-500' : 'bg-primary'}`}
                                        style={{ width: `${Math.min(utilization, 100)}%` }}
                                      />
                                    </div>
                                    <span className="text-xs w-10 text-right">{utilization.toFixed(0)}%</span>
                                  </div>
                                </TableCell>
                                <TableCell className="text-right text-xs text-muted-foreground">
                                  {sub.costPerUser > 0 ? `$${sub.costPerUser.toFixed(2)}` : 'Free'}
                                </TableCell>
                                <TableCell className="text-right font-medium">
                                  {monthlySpend > 0 ? `$${monthlySpend.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : '—'}
                                </TableCell>
                              </TableRow>
                            );
                          })}
                        <TableRow className="bg-muted/20 font-medium">
                          <TableCell colSpan={2}>Total</TableCell>
                          <TableCell className="text-right">{subscriptions.reduce((a, s) => a + s.enabled, 0).toLocaleString()}</TableCell>
                          <TableCell className="text-right">{subscriptions.reduce((a, s) => a + s.consumed, 0).toLocaleString()}</TableCell>
                          <TableCell className="text-right">{subscriptions.reduce((a, s) => a + s.available, 0).toLocaleString()}</TableCell>
                          <TableCell />
                          <TableCell />
                          <TableCell className="text-right">
                            ${subscriptions.reduce((a, s) => a + (s.consumed * s.costPerUser), 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                          </TableCell>
                        </TableRow>
                      </TableBody>
                    </Table>
                  </div>
                )}
              </CardContent>
            )}
          </Card>
        )}

        {/* Strategy Selector */}
        <div className="space-y-4">
          <div className="flex items-center gap-2">
            <h3 className="text-lg font-semibold font-display">Optimization Strategy</h3>
            <Badge variant="secondary" className="font-normal text-xs">Usage-Aware</Badge>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
            {([
              { key: "current" as Strategy, icon: CheckCircle2, title: "Current State", desc: "Keep existing license assignments unchanged." },
              { key: "security" as Strategy, icon: Shield, title: "Maximize Security", desc: "Upgrade tiers for compliance, add Copilot for power users." },
              { key: "cost" as Strategy, icon: TrendingDown, title: "Minimize Cost", desc: "Downgrade underutilized licenses, remove unused add-ons." },
              { key: "balanced" as Strategy, icon: Scale, title: "Balanced", desc: "Upgrade underprovisioned, downgrade overprovisioned." },
              { key: "custom" as Strategy, icon: Filter, title: "Custom", desc: "Configure your own analysis rules and thresholds." },
            ]).map(({ key, icon: Icon, title, desc }) => {
              const stats = data.length > 0 ? getStrategyStats(key) : null;
              return (
                <Card
                  key={key}
                  className={`cursor-pointer transition-all hover:border-primary/50 ${strategy === key ? 'ring-2 ring-primary border-primary' : 'border-border/50'}`}
                  onClick={() => handleStrategyChange(key)}
                  data-testid={`strategy-${key}`}
                >
                  <CardHeader className="p-4 pb-2">
                    <CardTitle className="text-base flex items-center gap-2">
                      <div className={`p-1.5 rounded-md ${strategy === key ? 'bg-primary/10 text-primary' : 'bg-muted text-muted-foreground'}`}>
                        <Icon className="h-4 w-4" />
                      </div>
                      {title}
                    </CardTitle>
                  </CardHeader>
                  <CardContent className="p-4 pt-1">
                    <CardDescription className="text-xs">{desc}</CardDescription>
                    {stats && key !== "current" && stats.affected > 0 && (
                      <div className="mt-2 pt-2 border-t border-border/30 flex flex-col gap-1">
                        <div className="flex justify-between text-xs">
                          <span className="text-muted-foreground">Users affected</span>
                          <span className="font-medium">{stats.affected}</span>
                        </div>
                        <div className="flex justify-between text-xs">
                          <span className="text-muted-foreground">Net cost</span>
                          <span className={`font-medium ${stats.delta < 0 ? 'text-green-600' : stats.delta > 0 ? 'text-amber-600' : ''}`}>
                            {stats.delta > 0 ? '+' : ''}{stats.delta < 0 ? '-' : ''}${Math.abs(stats.delta).toFixed(0)}/mo
                          </span>
                        </div>
                        <div className="flex gap-3 text-xs text-muted-foreground">
                          {stats.upgrades > 0 && <span className="text-amber-600">{stats.upgrades} upgrade{stats.upgrades !== 1 ? 's' : ''}</span>}
                          {stats.downgrades > 0 && <span className="text-green-600">{stats.downgrades} saving{stats.downgrades !== 1 ? 's' : ''}</span>}
                        </div>
                      </div>
                    )}
                  </CardContent>
                </Card>
              );
            })}
          </div>

          {strategy === 'custom' && (() => {
            const allLicenses = data.flatMap(u => u.licenses);
            const licSet = new Set(allLicenses);
            const hasE1 = licSet.has("Office 365 E1");
            const hasF1 = licSet.has("Microsoft 365 F1") || licSet.has("Microsoft 365 F3") || licSet.has("Office 365 F3");
            const hasE3 = licSet.has("Microsoft 365 E3") || licSet.has("Office 365 E3");
            const hasE5 = licSet.has("Microsoft 365 E5") || licSet.has("Office 365 E5");
            const hasBizBasic = licSet.has("Microsoft 365 Business Basic");
            const hasBizStd = licSet.has("Microsoft 365 Business Standard");
            const hasBizPrem = licSet.has("Microsoft 365 Business Premium");
            const hasVisio = licSet.has("Visio Plan 2") || licSet.has("Visio Plan 1");
            const hasProject = licSet.has("Project Plan 3") || licSet.has("Project Plan 5") || licSet.has("Project Plan 1");
            const hasPBIPremium = licSet.has("Power BI Premium Per User");
            const hasDefenderAddon = licSet.has("Defender for Office 365 P1") || licSet.has("Defender for Office 365 P2") || licSet.has("Defender for Business");
            const hasOneDriveStandalone = licSet.has("OneDrive for Business P1") || licSet.has("OneDrive for Business P2");
            const hasEMSAddon = licSet.has("Entra ID P1") || licSet.has("Entra ID P2") || licSet.has("Microsoft Intune Plan 1");
            const hasExchangeStandalone = licSet.has("Exchange Online Plan 1") || licSet.has("Exchange Online Plan 2") || licSet.has("Exchange Online Kiosk");
            const hasTrials = allLicenses.some(l => ["Teams Exploratory", "Power Automate Free", "Power Apps Trial", "Microsoft Stream", "Microsoft Teams (Free)", "Microsoft Teams Trial", "Power Virtual Agents Trial", "Microsoft Clipchamp", "Rights Management Adhoc"].includes(l));
            const hasRedundant = hasDefenderAddon || hasOneDriveStandalone || hasEMSAddon;
            const hasOverlap = hasExchangeStandalone || hasTrials;
            const hasAddons = hasVisio || hasProject || hasPBIPremium;

            const isScopedRule = (key: string): key is keyof CustomRulesState & ('upgradeUnderprovisioned' | 'upgradeToE5' | 'upgradeBasicToStandard' | 'upgradeToBizPremium' | 'downgradeUnderutilizedE5' | 'downgradeOverprovisionedE3' | 'downgradeUnderutilizedBizPremium' | 'downgradeBizStandardToBasic') => {
              return typeof (customRules as any)[key] === 'object' && (customRules as any)[key] !== null;
            };

            const isRuleEnabled = (key: string): boolean => {
              const val = (customRules as any)[key];
              if (typeof val === 'boolean') return val;
              if (typeof val === 'object' && val !== null) return val.enabled;
              return false;
            };

            const toggleRule = (key: string) => {
              setCustomRules(prev => {
                const val = (prev as any)[key];
                if (typeof val === 'boolean') return { ...prev, [key]: !val };
                if (typeof val === 'object' && val !== null) return { ...prev, [key]: { ...val, enabled: !val.enabled } };
                return prev;
              });
            };

            const setScopeForRule = (key: string, scope: RuleScope) => {
              setCustomRules(prev => {
                const val = (prev as any)[key];
                if (typeof val === 'object' && val !== null) return { ...prev, [key]: { ...val, scope } };
                return prev;
              });
            };

            const setDepartmentsForRule = (key: string, dept: string) => {
              setCustomRules(prev => {
                const val = (prev as any)[key];
                if (typeof val === 'object' && val !== null) {
                  const depts: string[] = val.departments || [];
                  const newDepts = depts.includes(dept) ? depts.filter((d: string) => d !== dept) : [...depts, dept];
                  return { ...prev, [key]: { ...val, departments: newDepts } };
                }
                return prev;
              });
            };

            const setRuleThreshold = (key: string, threshold: number | undefined) => {
              setCustomRules(prev => {
                const val = (prev as any)[key];
                if (typeof val === 'object' && val !== null) return { ...prev, [key]: { ...val, threshold } };
                return prev;
              });
            };

            const getRuleImpact = (key: string): { affected: number; delta: number } => {
              const baseState: CustomRulesState = {
                upgradeUnderprovisioned: { enabled: false, scope: "all", departments: [] },
                upgradeToE5: { enabled: false, scope: "all", departments: [] },
                upgradeBasicToStandard: { enabled: false, scope: "all", departments: [] },
                upgradeToBizPremium: { enabled: false, scope: "all", departments: [] },
                downgradeUnderutilizedE5: { enabled: false, scope: "all", departments: [] },
                downgradeOverprovisionedE3: { enabled: false, scope: "all", departments: [] },
                downgradeUnderutilizedBizPremium: { enabled: false, scope: "all", departments: [] },
                downgradeBizStandardToBasic: { enabled: false, scope: "all", departments: [] },
                removeUnusedAddons: false, consolidateOverlap: false, removeRedundantAddons: false,
                addCopilotPowerUsers: false, usageThreshold: customRules.usageThreshold,
              };
              const testRules = { ...baseState };
              const currentVal = (customRules as any)[key];
              if (typeof currentVal === 'boolean') {
                (testRules as any)[key] = true;
              } else if (typeof currentVal === 'object') {
                (testRules as any)[key] = { ...currentVal, enabled: true };
              }
              const result = analyzeAllUsers("custom", testRules);
              const baseCost = data.reduce((a, c) => a + c.cost, 0);
              const newCost = result.reduce((a, c) => a + c.cost, 0);
              const affected = result.filter((u, i) => JSON.stringify(sortLicenses(data[i]?.licenses || [])) !== JSON.stringify(u.licenses)).length;
              return { affected, delta: newCost - baseCost };
            };

            type RuleDef = { key: string; label: string; hint: string; scoped: boolean; hasThreshold: boolean };

            const upgradeRules: RuleDef[] = [
              (hasE1 || hasF1) ? { key: 'upgradeUnderprovisioned', label: 'Upgrade underprovisioned E1/F1', hint: `${hasE1 ? 'E1→E3 for high-usage users.' : ''} ${hasF1 ? 'F1→Basic for users over 2GB.' : ''}`.trim(), scoped: true, hasThreshold: false } : null,
              hasBizBasic ? { key: 'upgradeBasicToStandard', label: 'Upgrade Business Basic → Standard', hint: `${allLicenses.filter(l => l === "Microsoft 365 Business Basic").length} Basic users. Adds desktop apps. +$6.50/user/mo.`, scoped: true, hasThreshold: false } : null,
              hasE3 ? { key: 'upgradeToE5', label: 'Upgrade E3 → E5', hint: 'Adds Defender P2, Cloud App Security, eDiscovery Premium. +$21/user/mo.', scoped: true, hasThreshold: false } : null,
              (hasBizStd || hasBizBasic) ? { key: 'upgradeToBizPremium', label: 'Upgrade Basic/Standard → Premium', hint: 'Adds Intune MDM, Defender for Business, Conditional Access. +$9.50–16/user/mo.', scoped: true, hasThreshold: false } : null,
              { key: 'addCopilotPowerUsers', label: 'Add Copilot for power users', hint: 'GitHub Copilot for Engineering, M365 Copilot for IT/Design/Analytics.', scoped: false, hasThreshold: false },
            ].filter(Boolean) as RuleDef[];

            const downgradeRules: RuleDef[] = [
              hasE5 ? { key: 'downgradeUnderutilizedE5', label: 'Downgrade E5 → E3', hint: `${allLicenses.filter(l => l === "Microsoft 365 E5" || l === "Office 365 E5").length} E5 users. Saves $21/user/mo.`, scoped: true, hasThreshold: true } : null,
              hasE3 ? { key: 'downgradeOverprovisionedE3', label: 'Downgrade E3 → Business Standard', hint: 'Saves $23.50/user/mo for users not needing enterprise compliance.', scoped: true, hasThreshold: true } : null,
              hasBizPrem ? { key: 'downgradeUnderutilizedBizPremium', label: 'Downgrade Premium → Standard', hint: `${allLicenses.filter(l => l === "Microsoft 365 Business Premium").length} Premium users. Saves $9.50/user/mo.`, scoped: true, hasThreshold: true } : null,
              hasBizStd ? { key: 'downgradeBizStandardToBasic', label: 'Downgrade Standard → Basic', hint: `${allLicenses.filter(l => l === "Microsoft 365 Business Standard").length} Standard users. Saves $6.50/user/mo.`, scoped: true, hasThreshold: true } : null,
            ].filter(Boolean) as RuleDef[];

            const cleanupRules: RuleDef[] = [
              hasAddons ? { key: 'removeUnusedAddons', label: 'Remove unused add-ons', hint: `${[hasVisio && 'Visio', hasProject && 'Project', hasPBIPremium && 'Power BI Premium'].filter(Boolean).join(', ')} in non-relevant departments.`, scoped: false, hasThreshold: false } : null,
              hasRedundant ? { key: 'removeRedundantAddons', label: 'Remove redundant add-ons', hint: `${[hasDefenderAddon && 'Defender', hasOneDriveStandalone && 'OneDrive', hasEMSAddon && 'Intune/Entra ID'].filter(Boolean).join(', ')} overlap with suite licenses.`, scoped: false, hasThreshold: false } : null,
              hasOverlap ? { key: 'consolidateOverlap', label: 'Consolidate overlapping licenses', hint: `${[hasExchangeStandalone && 'Exchange standalone', hasTrials && 'trial/free licenses'].filter(Boolean).join(' and ')} alongside suites.`, scoped: false, hasThreshold: false } : null,
            ].filter(Boolean) as RuleDef[];

            const recommendations: { text: string; ruleKey: string; config?: Partial<ScopedRule> }[] = [];
            const basicUserCount = allLicenses.filter(l => l === "Microsoft 365 Business Basic").length;
            const highUsageBasicCount = data.filter(u => u.licenses.includes("Microsoft 365 Business Basic") && u.maxGB > 0 && (u.usageGB / u.maxGB) > 0.5).length;
            if (basicUserCount > 0 && highUsageBasicCount / basicUserCount > 0.2) {
              recommendations.push({ text: `${highUsageBasicCount} of ${basicUserCount} Business Basic users have >50% mailbox usage — upgrade to Standard for desktop apps.`, ruleKey: 'upgradeBasicToStandard', config: { enabled: true, scope: "all" } });
            }
            const secDeptsBizNonPrem = data.filter(u => SECURITY_DEPTS.has(u.department) && (u.licenses.includes("Microsoft 365 Business Basic") || u.licenses.includes("Microsoft 365 Business Standard")));
            if (secDeptsBizNonPrem.length > 0) {
              recommendations.push({ text: `${secDeptsBizNonPrem.length} security/IT team members lack Premium's endpoint protection and Conditional Access.`, ruleKey: 'upgradeToBizPremium', config: { enabled: true, scope: "security" } });
            }
            if (hasRedundant) {
              const redundantCount = data.filter(u => {
                const hasSuite = u.licenses.some(l => SUITE_LICENSES.has(l));
                return hasSuite && u.licenses.some(l => ["Defender for Office 365 P1", "Defender for Office 365 P2", "OneDrive for Business P1", "OneDrive for Business P2", "Entra ID P1", "Microsoft Intune Plan 1"].includes(l));
              }).length;
              recommendations.push({ text: `${redundantCount} users have add-ons already included in their suite license — remove for direct savings.`, ruleKey: 'removeRedundantAddons' });
            }
            const trialCount = data.filter(u => u.licenses.some(l => ["Teams Exploratory", "Power Automate Free", "Power Apps Trial", "Microsoft Stream", "Microsoft Teams (Free)", "Rights Management Adhoc"].includes(l))).length;
            if (trialCount > data.length * 0.3 && trialCount > 0) {
              recommendations.push({ text: `${trialCount} users have trial/free licenses alongside their suite — clean up for clarity.`, ruleKey: 'consolidateOverlap' });
            }
            const highTierLowUsage = data.filter(u => {
              const hasHighTier = u.licenses.some(l => ["Microsoft 365 E5", "Office 365 E5", "Microsoft 365 Business Premium"].includes(l));
              return hasHighTier && u.maxGB > 0 && (u.usageGB / u.maxGB) < 0.2 && !SECURITY_DEPTS.has(u.department);
            }).length;
            if (highTierLowUsage > 0) {
              recommendations.push({ text: `${highTierLowUsage} premium-tier users have <20% usage in non-security departments — potential downgrade candidates.`, ruleKey: 'downgradeUnderutilizedE5' });
            }

            const renderScopedConfig = (rule: RuleDef) => {
              if (!isScopedRule(rule.key)) return null;
              const scopedVal = (customRules as any)[rule.key] as ScopedRule;
              if (!scopedVal.enabled) return null;
              return (
                <div className="mt-2 pt-2 border-t border-border/30 space-y-2" data-testid={`config-${rule.key}`}>
                  <div className="text-[11px] font-medium text-muted-foreground uppercase tracking-wider">Department Scope</div>
                  <div className="flex flex-wrap gap-1.5">
                    {(["all", "security", "custom"] as RuleScope[]).map(s => (
                      <button key={s} type="button" onClick={() => setScopeForRule(rule.key, s)}
                        className={`text-xs px-2.5 py-1 rounded-md border transition-colors ${scopedVal.scope === s ? 'bg-primary/10 border-primary/30 text-primary font-medium' : 'border-border/60 text-muted-foreground hover:bg-muted/30'}`}
                        data-testid={`scope-${rule.key}-${s}`}
                      >
                        {s === "all" ? "All departments" : s === "security" ? "Security depts only" : "Select departments"}
                      </button>
                    ))}
                  </div>
                  {scopedVal.scope === "custom" && (
                    <div className="flex flex-wrap gap-1.5 mt-1">
                      {departments.map(dept => (
                        <button key={dept} type="button" onClick={() => setDepartmentsForRule(rule.key, dept)}
                          className={`text-[11px] px-2 py-0.5 rounded border transition-colors ${scopedVal.departments.includes(dept) ? 'bg-primary/10 border-primary/30 text-primary' : 'border-border/50 text-muted-foreground hover:bg-muted/20'}`}
                          data-testid={`dept-${rule.key}-${dept}`}
                        >
                          {dept}
                        </button>
                      ))}
                    </div>
                  )}
                  {rule.hasThreshold && (
                    <div className="flex items-center gap-3 mt-1">
                      <div className="text-xs text-muted-foreground">Threshold:</div>
                      <input type="range" min={5} max={50} step={5}
                        value={scopedVal.threshold ?? customRules.usageThreshold}
                        onChange={(e) => setRuleThreshold(rule.key, Number(e.target.value))}
                        className="w-24 accent-primary" data-testid={`threshold-${rule.key}`}
                      />
                      <span className="text-xs font-medium w-8">{scopedVal.threshold ?? customRules.usageThreshold}%</span>
                    </div>
                  )}
                </div>
              );
            };

            const renderRuleCard = (rule: RuleDef) => {
              const enabled = isRuleEnabled(rule.key);
              const impact = enabled ? getRuleImpact(rule.key) : null;
              return (
                <div key={rule.key} className={`rounded-lg border p-3 transition-colors ${enabled ? 'border-primary/30 bg-primary/5' : 'border-border/60'}`}>
                  <button type="button" onClick={() => toggleRule(rule.key)} className="w-full text-left" data-testid={`toggle-custom-${rule.key}`}>
                    <div className="flex items-start justify-between gap-3">
                      <div className="flex-1 min-w-0">
                        <div className="font-medium text-sm">{rule.label}</div>
                        <div className="text-xs text-muted-foreground mt-0.5">{rule.hint}</div>
                        {impact && impact.affected > 0 && (
                          <div className="flex items-center gap-2 mt-1.5 text-[11px]">
                            <span className="text-muted-foreground">{impact.affected} user{impact.affected !== 1 ? 's' : ''}</span>
                            <span className={impact.delta < 0 ? 'text-green-600 font-medium' : impact.delta > 0 ? 'text-amber-600 font-medium' : 'text-muted-foreground'}>
                              {impact.delta > 0 ? '+' : ''}{impact.delta < 0 ? '-' : ''}${Math.abs(impact.delta).toFixed(0)}/mo
                            </span>
                          </div>
                        )}
                      </div>
                      <div className={`mt-0.5 h-5 w-9 rounded-full border border-border/60 flex items-center px-0.5 shrink-0 ${enabled ? 'bg-primary/20' : 'bg-muted'}`}>
                        <div className={`h-4 w-4 rounded-full bg-background shadow-sm transition-transform ${enabled ? 'translate-x-4' : 'translate-x-0'}`} />
                      </div>
                    </div>
                  </button>
                  {rule.scoped && renderScopedConfig(rule)}
                </div>
              );
            };

            return (
            <Card className="border-border/60 bg-card shadow-sm">
              <CardHeader className="pb-3">
                <CardTitle className="text-base">Custom Analysis Rules</CardTitle>
                <CardDescription>Configure each rule's scope, thresholds, and target departments. Rules that don't apply to your license mix are hidden.</CardDescription>
              </CardHeader>
              <CardContent className="space-y-5">
                {recommendations.length > 0 && (
                  <div className="rounded-lg border border-blue-200 bg-blue-50/50 dark:bg-blue-950/20 dark:border-blue-900/30 p-4 space-y-2.5">
                    <div className="flex items-center gap-2 text-sm font-medium text-blue-700 dark:text-blue-300">
                      <Info className="h-4 w-4" />
                      Recommended Actions
                    </div>
                    {recommendations.map((rec, i) => (
                      <div key={i} className="flex items-start justify-between gap-3" data-testid={`recommendation-${i}`}>
                        <div className="text-xs text-blue-800 dark:text-blue-200 leading-relaxed flex-1">{rec.text}</div>
                        <Button size="sm" variant="outline" className="shrink-0 h-7 text-xs border-blue-300 text-blue-700 hover:bg-blue-100 dark:border-blue-800 dark:text-blue-300 dark:hover:bg-blue-900/30"
                          onClick={() => {
                            setCustomRules(prev => {
                              const val = (prev as any)[rec.ruleKey];
                              if (typeof val === 'boolean') return { ...prev, [rec.ruleKey]: true };
                              if (typeof val === 'object' && rec.config) return { ...prev, [rec.ruleKey]: { ...val, ...rec.config } };
                              if (typeof val === 'object') return { ...prev, [rec.ruleKey]: { ...val, enabled: true } };
                              return prev;
                            });
                          }}
                          data-testid={`apply-recommendation-${i}`}
                        >
                          Apply
                        </Button>
                      </div>
                    ))}
                  </div>
                )}

                {upgradeRules.length > 0 && (
                <div>
                  <div className="text-xs font-medium text-muted-foreground uppercase tracking-wider mb-2">Upgrades</div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    {upgradeRules.map(renderRuleCard)}
                  </div>
                </div>
                )}
                {downgradeRules.length > 0 && (
                <div>
                  <div className="text-xs font-medium text-muted-foreground uppercase tracking-wider mb-2">Downgrades</div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    {downgradeRules.map(renderRuleCard)}
                  </div>
                </div>
                )}
                {cleanupRules.length > 0 && (
                <div>
                  <div className="text-xs font-medium text-muted-foreground uppercase tracking-wider mb-2">Cleanup</div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    {cleanupRules.map(renderRuleCard)}
                  </div>
                </div>
                )}
                <div>
                  <div className="text-xs font-medium text-muted-foreground uppercase tracking-wider mb-2">Global Settings</div>
                  <div className="flex items-center gap-4 p-3 border border-border/60 rounded-lg">
                    <div className="flex-1">
                      <div className="font-medium text-sm">Default low-usage threshold: {customRules.usageThreshold}%</div>
                      <div className="text-xs text-muted-foreground mt-1">Fallback for downgrade rules without a custom threshold.</div>
                    </div>
                    <input type="range" min={5} max={50} step={5} value={customRules.usageThreshold}
                      onChange={(e) => setCustomRules(prev => ({ ...prev, usageThreshold: Number(e.target.value) }))}
                      className="w-32 accent-primary" data-testid="slider-usage-threshold"
                    />
                  </div>
                </div>
              </CardContent>
            </Card>
            );
          })()}
        </div>

        {/* Data Table */}
        <Card className="shadow-md border-border/50 overflow-hidden">
          <CardHeader className="border-b border-border/50 bg-muted/10 pb-4">
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
              <div>
                <CardTitle>Combined User Directory</CardTitle>
                <CardDescription>
                  {strategy === 'current' 
                    ? "Joined on UPN from Entra ID and Exchange Online." 
                    : "Showing projected licenses based on selected strategy."}
                </CardDescription>
              </div>
              <div className="flex items-center gap-2 w-full sm:w-auto">
                <div className="relative w-full sm:w-64">
                  <Search className="absolute left-2.5 top-2.5 h-4 w-4 text-muted-foreground" />
                  <Input 
                    type="search" 
                    placeholder="Search users..." 
                    className="pl-9 bg-background"
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    data-testid="input-search"
                  />
                </div>
                <Button
                  variant={activeFilterCount > 0 ? "default" : "outline"}
                  size="icon"
                  title="Filter"
                  onClick={() => setShowFilters(!showFilters)}
                  className="relative"
                  data-testid="button-filter"
                >
                  <Filter className="h-4 w-4" />
                  {activeFilterCount > 0 && (
                    <span className="absolute -top-1.5 -right-1.5 h-4 w-4 rounded-full bg-primary text-primary-foreground text-[10px] flex items-center justify-center font-bold">
                      {activeFilterCount}
                    </span>
                  )}
                </Button>
              </div>
            </div>
          </CardHeader>
          {showFilters && (
            <div className="px-6 py-3 border-b border-border/50 bg-muted/5 flex flex-wrap items-center gap-3">
              <div className="flex items-center gap-2">
                <span className="text-xs font-medium text-muted-foreground">Department</span>
                <Select value={filterDepartment} onValueChange={setFilterDepartment}>
                  <SelectTrigger className="h-8 w-[160px] text-xs bg-background" data-testid="filter-department">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="all">All departments</SelectItem>
                    {departments.map(d => (
                      <SelectItem key={d} value={d}>{d}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-medium text-muted-foreground">Status</span>
                <Select value={filterStatus} onValueChange={setFilterStatus}>
                  <SelectTrigger className="h-8 w-[130px] text-xs bg-background" data-testid="filter-status">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="all">All statuses</SelectItem>
                    <SelectItem value="Active">Active</SelectItem>
                    <SelectItem value="Warning">Warning</SelectItem>
                    <SelectItem value="Critical">Critical</SelectItem>
                  </SelectContent>
                </Select>
              </div>
              {strategy !== "current" && (
                <div className="flex items-center gap-2">
                  <span className="text-xs font-medium text-muted-foreground">Changes</span>
                  <Select value={filterModified} onValueChange={setFilterModified}>
                    <SelectTrigger className="h-8 w-[150px] text-xs bg-background" data-testid="filter-modified">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="all">All users</SelectItem>
                      <SelectItem value="changed">Changed only</SelectItem>
                      <SelectItem value="unchanged">Unchanged only</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
              )}
              {activeFilterCount > 0 && (
                <Button
                  variant="ghost"
                  size="sm"
                  className="h-8 text-xs text-muted-foreground"
                  onClick={() => { setFilterDepartment("all"); setFilterStatus("all"); setFilterModified("all"); }}
                  data-testid="button-clear-filters"
                >
                  <X className="h-3 w-3 mr-1" />
                  Clear filters
                </Button>
              )}
              <div className="ml-auto text-xs text-muted-foreground">
                {filteredData.length} of {optimizedData.length} users
              </div>
            </div>
          )}
          <div className="overflow-x-auto">
            <Table>
              <TableHeader className="bg-muted/30">
                <TableRow>
                  <TableHead className="w-[250px]">User & UPN</TableHead>
                  <TableHead>Department</TableHead>
                  <TableHead>Assigned Licenses</TableHead>
                  <TableHead className="text-right">Mailbox (GB)</TableHead>
                  <TableHead className="text-right">License Cost</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {isSyncing ? (
                  // Loading state rows
                  Array.from({ length: 5 }).map((_, i) => (
                    <TableRow key={i}>
                      <TableCell><Skeleton className="h-10 w-full" /></TableCell>
                      <TableCell><Skeleton className="h-6 w-24" /></TableCell>
                      <TableCell><Skeleton className="h-12 w-full" /></TableCell>
                      <TableCell><Skeleton className="h-6 w-16 ml-auto" /></TableCell>
                      <TableCell><Skeleton className="h-6 w-16 ml-auto" /></TableCell>
                    </TableRow>
                  ))
                ) : filteredData.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={5} className="h-32 text-center text-muted-foreground">
                      {searchTerm || activeFilterCount > 0
                        ? `No users found matching your ${searchTerm ? 'search' : ''}${searchTerm && activeFilterCount > 0 ? ' and ' : ''}${activeFilterCount > 0 ? 'filters' : ''}`
                        : "No users found"}
                    </TableCell>
                  </TableRow>
                ) : (
                  filteredData.map((user, rowIndex) => {
                    const originalUser = data.find(u => u.id === user.id)!;
                    const isModified = strategy !== 'current' && JSON.stringify(sortLicenses(originalUser.licenses)) !== JSON.stringify(user.licenses);
                    
                    return (
                      <TableRow 
                        key={`${user.id}-${strategyKey}`} 
                        className={`transition-colors group row-stagger ${isModified ? 'bg-primary/5 hover:bg-primary/10' : 'hover:bg-muted/20'}`} 
                        style={{ animationDelay: `${Math.min(rowIndex * 20, 400)}ms` }}
                        data-testid={`row-user-${user.id}`}
                      >
                        <TableCell>
                          <div className="font-medium">{user.displayName}</div>
                          <div className="text-sm text-muted-foreground truncate max-w-[200px]">{user.upn}</div>
                        </TableCell>
                        <TableCell>
                          <Badge variant="secondary" className="bg-secondary/50 font-normal">
                            {user.department}
                          </Badge>
                        </TableCell>
                        <TableCell>
                          <div className="flex flex-col gap-1">
                            {isModified ? (
                              <>
                                <div className="flex flex-col gap-1">
                                  {sortLicenses(originalUser.licenses).map((license, i) => (
                                    <div key={`old-${i}`} className="flex items-center gap-1.5">
                                      <Badge variant="outline" className="text-xs border-border/40 opacity-50 line-through w-fit">
                                        {license}
                                      </Badge>
                                    </div>
                                  ))}
                                </div>
                                <div className="flex items-center gap-1 my-0.5">
                                  <ArrowRight className="h-3 w-3 text-primary shrink-0" />
                                  <div className="h-px flex-1 bg-primary/20" />
                                </div>
                                <div className="flex flex-col gap-1">
                                  {user.licenses.map((license, i) => {
                                    const isSuite = SUITE_LICENSES.has(license);
                                    const isNew = !originalUser.licenses.includes(license);
                                    return (
                                      <Popover key={`new-${i}`}>
                                        <PopoverTrigger asChild>
                                          <Badge
                                            className={`text-xs w-fit cursor-pointer badge-hover ${isNew ? 'bg-primary/20 text-primary border-primary/20' : isSuite ? 'bg-blue-500/10 text-blue-700 dark:text-blue-300 border-blue-500/20' : 'bg-secondary/50 border-border/40'}`}
                                            data-testid={`badge-license-new-${i}`}
                                          >
                                            {license}
                                          </Badge>
                                        </PopoverTrigger>
                                        <PopoverContent className="w-80" side="right" align="start">
                                          <LicensePopoverContent licenseName={license} />
                                        </PopoverContent>
                                      </Popover>
                                    );
                                  })}
                                </div>
                                {user.reasons && user.reasons.length > 0 && (
                                  <div className="mt-1 space-y-0.5">
                                    {user.reasons.map((reason, i) => (
                                      <div key={i} className="text-[11px] text-muted-foreground leading-tight flex items-start gap-1">
                                        <Info className="h-3 w-3 shrink-0 mt-0.5 text-primary/60" />
                                        <span>{reason}</span>
                                      </div>
                                    ))}
                                  </div>
                                )}
                              </>
                            ) : (
                              <div className="flex flex-col gap-1">
                                {user.licenses.map((license, i) => {
                                  const isSuite = SUITE_LICENSES.has(license);
                                  return (
                                    <Popover key={i}>
                                      <PopoverTrigger asChild>
                                        <Badge
                                          variant="outline"
                                          className={`text-xs w-fit cursor-pointer badge-hover ${isSuite ? 'border-blue-500/30 bg-blue-500/5 text-blue-700 dark:text-blue-300' : 'border-border/60'}`}
                                          data-testid={`badge-license-${i}`}
                                        >
                                          {license}
                                        </Badge>
                                      </PopoverTrigger>
                                      <PopoverContent className="w-80" side="right" align="start">
                                        <LicensePopoverContent licenseName={license} />
                                      </PopoverContent>
                                    </Popover>
                                  );
                                })}
                              </div>
                            )}
                          </div>
                        </TableCell>
                        <TableCell className="text-right align-top pt-4">
                          <div className="flex flex-col items-end gap-1">
                            <div className="flex items-center gap-1.5">
                              {user.status === "Critical" && <AlertCircle className="h-3.5 w-3.5 text-destructive" />}
                              {user.status === "Warning" && <AlertCircle className="h-3.5 w-3.5 text-amber-500" />}
                              {user.status === "Active" && <CheckCircle2 className="h-3.5 w-3.5 text-green-500 opacity-0 group-hover:opacity-100 transition-opacity" />}
                              <span className={`font-medium ${user.status === 'Critical' ? 'text-destructive' : ''}`}>
                                {user.usageGB.toFixed(1)}
                              </span>
                              <span className="text-muted-foreground text-xs">/ {user.maxGB}</span>
                            </div>
                            {/* Progress bar for storage */}
                            <div className="w-24 h-1.5 bg-secondary rounded-full overflow-hidden">
                              <div 
                                className={`h-full rounded-full ${user.status === 'Critical' ? 'bg-destructive' : user.status === 'Warning' ? 'bg-amber-500' : 'bg-primary'}`} 
                                style={{ width: `${(user.usageGB / user.maxGB) * 100}%` }}
                              />
                            </div>
                          </div>
                        </TableCell>
                        <TableCell className="text-right align-top pt-4 font-medium">
                          {isModified ? (
                            <div className="flex flex-col items-end gap-1">
                              <span className="text-muted-foreground line-through text-xs">${originalUser.cost.toFixed(2)}</span>
                              <span className="text-primary">${user.cost.toFixed(2)}</span>
                            </div>
                          ) : (
                            <span>${user.cost.toFixed(2)}</span>
                          )}
                        </TableCell>
                      </TableRow>
                    );
                  })
                )}
              </TableBody>
            </Table>
          </div>
        </Card>
        {/* Industry Insights */}
        <Card className="border-border/50 shadow-sm">
          <CardHeader
            className="cursor-pointer py-3 px-6"
            onClick={() => { setShowNews(!showNews); if (!showNews && newsItems.length === 0) loadNews(); }}
          >
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="p-1.5 rounded-md bg-primary/10">
                  <Newspaper className="h-4 w-4 text-primary" />
                </div>
                <div>
                  <CardTitle className="text-base">Industry Insights</CardTitle>
                  <CardDescription className="text-xs">
                    Latest Microsoft 365 licensing news and updates
                  </CardDescription>
                </div>
              </div>
              <div className="flex items-center gap-2">
                {showNews && (
                  <Button
                    variant="ghost"
                    size="sm"
                    className="h-7 text-xs"
                    onClick={(e) => { e.stopPropagation(); loadNews(); }}
                    disabled={newsLoading}
                    data-testid="button-refresh-news"
                  >
                    <RefreshCcw className={`h-3 w-3 mr-1 ${newsLoading ? 'animate-spin' : ''}`} />
                    Refresh
                  </Button>
                )}
                {showNews ? <ChevronUp className="h-4 w-4 text-muted-foreground" /> : <ChevronDown className="h-4 w-4 text-muted-foreground" />}
              </div>
            </div>
          </CardHeader>
          {showNews && (
            <CardContent className="pt-0 px-6 pb-4">
              {newsLoading ? (
                <div className="space-y-3">
                  {Array.from({ length: 3 }).map((_, i) => (
                    <Skeleton key={i} className="h-16 w-full" />
                  ))}
                </div>
              ) : newsItems.length === 0 ? (
                <div className="text-sm text-muted-foreground text-center py-6">
                  No news items available. Click Refresh to try again.
                </div>
              ) : (
                <div className="space-y-3">
                  {newsItems.map((item, i) => (
                    <a
                      key={i}
                      href={item.link}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="block p-3 rounded-lg border border-border/50 hover:border-primary/30 hover:bg-primary/5 transition-colors"
                      data-testid={`news-item-${i}`}
                    >
                      <div className="flex items-start justify-between gap-3">
                        <div className="flex-1 min-w-0">
                          <div className="font-medium text-sm leading-snug">{item.title}</div>
                          {item.summary && (
                            <p className="text-xs text-muted-foreground mt-1 leading-relaxed line-clamp-2">{item.summary}</p>
                          )}
                          {item.date && (
                            <div className="text-[11px] text-muted-foreground mt-1.5">
                              {new Date(item.date).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}
                            </div>
                          )}
                        </div>
                        <ExternalLink className="h-3.5 w-3.5 text-muted-foreground shrink-0 mt-0.5" />
                      </div>
                    </a>
                  ))}
                  {newsCachedAt && (
                    <div className="text-[11px] text-muted-foreground text-right">
                      Last updated: {new Date(newsCachedAt).toLocaleTimeString()}
                    </div>
                  )}
                </div>
              )}
            </CardContent>
          )}
        </Card>
      </main>

      {/* Tutorial Overlay */}
      {tutorialStep >= 0 && tutorialStep < TUTORIAL_STEPS.length && (
        <>
          <div className="tutorial-backdrop" onClick={endTutorial} />
          {tutorialTooltipPos && (
            <div
              className="tutorial-tooltip bg-card border border-border rounded-xl shadow-xl p-5"
              style={{ top: tutorialTooltipPos.top, left: tutorialTooltipPos.left }}
            >
              <div className="flex items-center justify-between mb-2">
                <div className="text-xs font-medium text-primary">
                  Step {tutorialStep + 1} of {TUTORIAL_STEPS.length}
                </div>
                <button onClick={endTutorial} className="text-muted-foreground hover:text-foreground" data-testid="button-skip-tutorial">
                  <X className="h-4 w-4" />
                </button>
              </div>
              <h4 className="font-semibold text-base mb-1">{TUTORIAL_STEPS[tutorialStep].title}</h4>
              <p className="text-sm text-muted-foreground leading-relaxed mb-4">{TUTORIAL_STEPS[tutorialStep].body}</p>
              <div className="flex items-center justify-between">
                <div className="flex gap-1">
                  {TUTORIAL_STEPS.map((_, i) => (
                    <div key={i} className={`h-1.5 w-6 rounded-full ${i === tutorialStep ? 'bg-primary' : i < tutorialStep ? 'bg-primary/30' : 'bg-muted'}`} />
                  ))}
                </div>
                <div className="flex gap-2">
                  <Button variant="ghost" size="sm" onClick={endTutorial} data-testid="button-end-tutorial">Skip</Button>
                  <Button size="sm" onClick={nextTutorialStep} data-testid="button-next-tutorial">
                    {tutorialStep === TUTORIAL_STEPS.length - 1 ? "Finish" : "Next"}
                  </Button>
                </div>
              </div>
            </div>
          )}
        </>
      )}

      <footer className="py-4 text-center text-xs text-muted-foreground border-t border-border/30">
        &copy; 2026 Cavaridge, LLC. All rights reserved.
      </footer>
    </div>
  );
}
