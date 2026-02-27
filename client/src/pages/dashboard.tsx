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
} from "@/lib/api";
import { useToast } from "@/hooks/use-toast";

const mockData = [
  { id: "1", displayName: "Alex Johnson", upn: "alex.j@company.com", department: "Engineering", licenses: ["Microsoft 365 E5", "Visio Plan 2"], usageGB: 45.2, maxGB: 100, cost: 57.00, status: "Active" },
  { id: "2", displayName: "Sarah Smith", upn: "sarah.s@company.com", department: "Marketing", licenses: ["Microsoft 365 E3"], usageGB: 82.5, maxGB: 100, cost: 36.00, status: "Warning" },
  { id: "3", displayName: "Michael Chen", upn: "michael.c@company.com", department: "Sales", licenses: ["Microsoft 365 E3", "Power BI Pro"], usageGB: 12.1, maxGB: 100, cost: 46.00, status: "Active" },
  { id: "4", displayName: "Emily Davis", upn: "emily.d@company.com", department: "HR", licenses: ["Office 365 E1"], usageGB: 4.8, maxGB: 50, cost: 10.00, status: "Active" },
  { id: "5", displayName: "James Wilson", upn: "james.w@company.com", department: "Engineering", licenses: ["Microsoft 365 E5", "GitHub Copilot"], usageGB: 95.1, maxGB: 100, cost: 77.00, status: "Critical" },
  { id: "6", displayName: "Jessica Taylor", upn: "jessica.t@company.com", department: "Finance", licenses: ["Microsoft 365 E5"], usageGB: 22.4, maxGB: 100, cost: 57.00, status: "Active" },
  { id: "7", displayName: "David Anderson", upn: "david.a@company.com", department: "IT", licenses: ["Microsoft 365 E5", "Project Plan 3"], usageGB: 68.9, maxGB: 100, cost: 87.00, status: "Active" },
];

type Strategy = "current" | "security" | "cost" | "balanced" | "custom";

export default function Dashboard() {
  const [, navigate] = useLocation();
  const { toast } = useToast();
  const [isSyncing, setIsSyncing] = useState(false);
  const [isGeneratingSummary, setIsGeneratingSummary] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [data, setData] = useState<typeof mockData>([]);
  const [dataSource, setDataSource] = useState<"mock" | "uploaded" | "microsoft">("mock");
  const [showUploadPanel, setShowUploadPanel] = useState(false);
  const [uploadedUserFile, setUploadedUserFile] = useState<string | null>(null);
  const [uploadedMailboxFile, setUploadedMailboxFile] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState("");
  const [strategy, setStrategy] = useState<Strategy>("current");
  const [commitment, setCommitment] = useState<"monthly" | "annual">("monthly");
  const userFileRef = useRef<HTMLInputElement>(null);
  const mailboxFileRef = useRef<HTMLInputElement>(null);

  const [msAuth, setMsAuth] = useState<{
    configured: boolean;
    connected: boolean;
    user?: { displayName: string; email: string };
    tenantId?: string;
  }>({ configured: false, connected: false });
  const [msLoading, setMsLoading] = useState(false);

  const [customRules, setCustomRules] = useState({
    upgradeE1ToE3: true,
    upgradeE3ToE5: false,
    downgradeE5ToE3NonCore: false,
    removeVisio: false,
    removeProject: false,
    addCopilotEngineering: false,
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
    setIsSyncing(true);
    setTimeout(() => {
      setData(mockData);
      setDataSource("mock");
      setIsSyncing(false);
    }, 1500);
  }, []);

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
    setData(mockData);
    setDataSource("mock");
    setUploadedUserFile(null);
    setUploadedMailboxFile(null);
    setShowUploadPanel(false);
    toast({ title: "Reset to demo data", description: "Upload your M365 reports to analyze real data." });
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
        setData(mockData);
        setDataSource("mock");
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

  // Apply Optimization Logic
  const optimizedData = useMemo(() => {
    if (strategy === "current") return data;

    const rules =
      strategy === "custom"
        ? customRules
        : {
            upgradeE1ToE3: strategy === "security" || strategy === "balanced",
            upgradeE3ToE5: strategy === "security",
            downgradeE5ToE3NonCore: strategy === "cost" || strategy === "balanced",
            removeVisio: strategy === "cost",
            removeProject: strategy === "cost",
            addCopilotEngineering: strategy === "security",
          };

    return data.map((user) => {
      let newLicenses = [...user.licenses];
      let newCost = user.cost;

      if (rules.upgradeE1ToE3 && newLicenses.includes("Office 365 E1")) {
        newLicenses = newLicenses.filter((l) => l !== "Office 365 E1");
        newLicenses.push("Microsoft 365 E3");
        newCost += 26;
      }

      if (rules.upgradeE3ToE5 && newLicenses.includes("Microsoft 365 E3")) {
        newLicenses = newLicenses.filter((l) => l !== "Microsoft 365 E3");
        newLicenses.push("Microsoft 365 E5");
        newCost += 21;
      }

      if (
        rules.downgradeE5ToE3NonCore &&
        newLicenses.includes("Microsoft 365 E5") &&
        !["IT", "Engineering"].includes(user.department)
      ) {
        newLicenses = newLicenses.filter((l) => l !== "Microsoft 365 E5");
        newLicenses.push("Microsoft 365 E3");
        newCost -= 21;
      }

      if (rules.removeVisio && newLicenses.includes("Visio Plan 2")) {
        newLicenses = newLicenses.filter((l) => l !== "Visio Plan 2");
        newCost -= 15;
      }

      if (rules.removeProject && newLicenses.includes("Project Plan 3")) {
        newLicenses = newLicenses.filter((l) => l !== "Project Plan 3");
        newCost -= 30;
      }

      if (
        rules.addCopilotEngineering &&
        user.department === "Engineering" &&
        !newLicenses.includes("GitHub Copilot")
      ) {
        newLicenses.push("GitHub Copilot");
        newCost += 20;
      }

      return { ...user, licenses: newLicenses, cost: newCost };
    });
  }, [customRules, data, strategy]);

  const filteredData = optimizedData.filter(item => 
    item.displayName.toLowerCase().includes(searchTerm.toLowerCase()) || 
    item.upn.toLowerCase().includes(searchTerm.toLowerCase()) ||
    item.department.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const applyRules = (rules: typeof customRules) => {
    return data.map((user) => {
      let newLicenses = [...user.licenses];
      let newCost = user.cost;
      if (rules.upgradeE1ToE3 && newLicenses.includes("Office 365 E1")) {
        newLicenses = newLicenses.filter((l) => l !== "Office 365 E1");
        newLicenses.push("Microsoft 365 E3");
        newCost += 26;
      }
      if (rules.upgradeE3ToE5 && newLicenses.includes("Microsoft 365 E3")) {
        newLicenses = newLicenses.filter((l) => l !== "Microsoft 365 E3");
        newLicenses.push("Microsoft 365 E5");
        newCost += 21;
      }
      if (rules.downgradeE5ToE3NonCore && newLicenses.includes("Microsoft 365 E5") && !["IT", "Engineering"].includes(user.department)) {
        newLicenses = newLicenses.filter((l) => l !== "Microsoft 365 E5");
        newLicenses.push("Microsoft 365 E3");
        newCost -= 21;
      }
      if (rules.removeVisio && newLicenses.includes("Visio Plan 2")) {
        newLicenses = newLicenses.filter((l) => l !== "Visio Plan 2");
        newCost -= 15;
      }
      if (rules.removeProject && newLicenses.includes("Project Plan 3")) {
        newLicenses = newLicenses.filter((l) => l !== "Project Plan 3");
        newCost -= 30;
      }
      if (rules.addCopilotEngineering && user.department === "Engineering" && !newLicenses.includes("GitHub Copilot")) {
        newLicenses.push("GitHub Copilot");
        newCost += 20;
      }
      return { ...user, licenses: newLicenses, cost: newCost };
    });
  };

  const costForStrategy = (strat: Strategy) => {
    if (strat === "current") return data.reduce((a, c) => a + c.cost, 0);
    const rules = strat === "custom" ? customRules : {
      upgradeE1ToE3: strat === "security" || strat === "balanced",
      upgradeE3ToE5: strat === "security",
      downgradeE5ToE3NonCore: strat === "cost" || strat === "balanced",
      removeVisio: strat === "cost",
      removeProject: strat === "cost",
      addCopilotEngineering: strat === "security",
    };
    return applyRules(rules).reduce((a, c) => a + c.cost, 0);
  };

  const baseTotalCost = data.reduce((acc, curr) => acc + curr.cost, 0);
  const projectedTotalCost = optimizedData.reduce((acc, curr) => acc + curr.cost, 0);
  const costDiff = projectedTotalCost - baseTotalCost;

  const commitmentMultiplier = commitment === "annual" ? 0.85 : 1;
  const totalCost = projectedTotalCost * commitmentMultiplier;

  const baseTotalCostCommitted = baseTotalCost * commitmentMultiplier;
  const costDiffCommitted = totalCost - baseTotalCostCommitted;

  const totalUsers = optimizedData.length;
  const totalStorage = optimizedData.reduce((acc, curr) => acc + curr.usageGB, 0);

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
        <div className="flex items-center gap-2">
          <div className="h-8 w-8 rounded-md bg-primary flex items-center justify-center text-primary-foreground font-bold">
            A
          </div>
          <h1 className="text-xl font-semibold tracking-tight">Astra</h1>
        </div>
        <div className="flex items-center gap-3">
          {dataSource === "mock" && data.length > 0 && (
            <Badge variant="outline" className="text-xs font-normal text-amber-600 border-amber-300">Sample Data</Badge>
          )}
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
          <Button size="sm" className="gap-2" onClick={handleExportXlsx} data-testid="button-export">
            <Download className="h-4 w-4" />
            Export XLSX
          </Button>
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
                  disabled={isUploading || dataSource === "mock"}
                  data-testid="button-upload-mailbox"
                >
                  {isUploading ? <Loader2 className="h-4 w-4 animate-spin" /> : <Upload className="h-4 w-4" />}
                  {uploadedMailboxFile ? "Replace file" : "Upload Mailbox Usage CSV/XLSX"}
                </Button>
                {dataSource === "mock" && (
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
      <main className="flex-1 p-8 max-w-7xl mx-auto w-full space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
        
        <div className="flex flex-col gap-4">
          <div className="flex flex-col gap-2">
            <h2 className="text-3xl font-display font-semibold">M365 Insights</h2>
            <p className="text-muted-foreground">Automated merge of Active Users and Mailbox Usage reports with actionable insights.</p>
          </div>

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
        </div>

        {/* KPIs */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <Card className="shadow-sm border-border/50">
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium text-muted-foreground">Total Active Users</CardTitle>
              <Users className="h-4 w-4 text-muted-foreground" />
            </CardHeader>
            <CardContent>
              {isSyncing ? (
                <Skeleton className="h-8 w-20" />
              ) : (
                <div className="text-3xl font-bold font-display" data-testid="text-total-users">{totalUsers}</div>
              )}
            </CardContent>
          </Card>
          
          <Card className="shadow-sm border-border/50">
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium text-muted-foreground">Total Mailbox Usage</CardTitle>
              <Database className="h-4 w-4 text-muted-foreground" />
            </CardHeader>
            <CardContent>
               {isSyncing ? (
                <Skeleton className="h-8 w-24" />
              ) : (
                <div className="text-3xl font-bold font-display" data-testid="text-total-storage">{totalStorage.toFixed(1)} GB</div>
              )}
            </CardContent>
          </Card>

          <Card className={`shadow-sm border-border/50 transition-colors ${strategy !== 'current' ? 'bg-primary/5 border-primary/20' : ''}`}>
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
                    ${totalCost.toFixed(2)}
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

        {/* Strategy Selector */}
        <div className="space-y-4">
          <div className="flex items-center gap-2">
            <h3 className="text-lg font-semibold font-display">Optimization Strategy</h3>
            <Badge variant="secondary" className="font-normal text-xs">AI Recommended</Badge>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <Card 
              className={`cursor-pointer transition-all hover:border-primary/50 ${strategy === 'current' ? 'ring-2 ring-primary border-primary' : 'border-border/50'}`}
              onClick={() => setStrategy('current')}
              data-testid="strategy-current"
            >
              <CardHeader className="p-4 pb-2">
                <CardTitle className="text-base flex items-center gap-2">
                  <div className={`p-1.5 rounded-md ${strategy === 'current' ? 'bg-primary/10 text-primary' : 'bg-muted text-muted-foreground'}`}>
                    <CheckCircle2 className="h-4 w-4" />
                  </div>
                  Current State
                </CardTitle>
              </CardHeader>
              <CardContent className="p-4 pt-2">
                <CardDescription>Keep existing license assignments with no changes.</CardDescription>
              </CardContent>
            </Card>

            <Card 
              className={`cursor-pointer transition-all hover:border-primary/50 ${strategy === 'security' ? 'ring-2 ring-primary border-primary' : 'border-border/50'}`}
              onClick={() => setStrategy('security')}
              data-testid="strategy-security"
            >
              <CardHeader className="p-4 pb-2">
                <CardTitle className="text-base flex items-center gap-2">
                  <div className={`p-1.5 rounded-md ${strategy === 'security' ? 'bg-primary/10 text-primary' : 'bg-muted text-muted-foreground'}`}>
                    <Shield className="h-4 w-4" />
                  </div>
                  Maximize Security
                </CardTitle>
              </CardHeader>
              <CardContent className="p-4 pt-2">
                <CardDescription>Upgrade users to E5 and add Copilot where applicable.</CardDescription>
              </CardContent>
            </Card>

            <Card 
              className={`cursor-pointer transition-all hover:border-primary/50 ${strategy === 'cost' ? 'ring-2 ring-primary border-primary' : 'border-border/50'}`}
              onClick={() => setStrategy('cost')}
              data-testid="strategy-cost"
            >
              <CardHeader className="p-4 pb-2">
                <CardTitle className="text-base flex items-center gap-2">
                  <div className={`p-1.5 rounded-md ${strategy === 'cost' ? 'bg-primary/10 text-primary' : 'bg-muted text-muted-foreground'}`}>
                    <TrendingDown className="h-4 w-4" />
                  </div>
                  Minimize Cost
                </CardTitle>
              </CardHeader>
              <CardContent className="p-4 pt-2">
                <CardDescription>Downgrade unused E5s and remove expensive add-ons.</CardDescription>
              </CardContent>
            </Card>

            <Card 
              className={`cursor-pointer transition-all hover:border-primary/50 ${strategy === 'balanced' ? 'ring-2 ring-primary border-primary' : 'border-border/50'}`}
              onClick={() => setStrategy('balanced')}
              data-testid="strategy-balanced"
            >
              <CardHeader className="p-4 pb-2">
                <CardTitle className="text-base flex items-center gap-2">
                  <div className={`p-1.5 rounded-md ${strategy === 'balanced' ? 'bg-primary/10 text-primary' : 'bg-muted text-muted-foreground'}`}>
                    <Scale className="h-4 w-4" />
                  </div>
                  Balanced Approach
                </CardTitle>
              </CardHeader>
              <CardContent className="p-4 pt-2">
                <CardDescription>Upgrade E1s for baseline security, downgrade non-core E5s.</CardDescription>
              </CardContent>
            </Card>

            <Card 
              className={`cursor-pointer transition-all hover:border-primary/50 ${strategy === 'custom' ? 'ring-2 ring-primary border-primary' : 'border-border/50'}`}
              onClick={() => setStrategy('custom')}
              data-testid="strategy-custom"
            >
              <CardHeader className="p-4 pb-2">
                <CardTitle className="text-base flex items-center gap-2">
                  <div className={`p-1.5 rounded-md ${strategy === 'custom' ? 'bg-primary/10 text-primary' : 'bg-muted text-muted-foreground'}`}>
                    <Filter className="h-4 w-4" />
                  </div>
                  Custom
                </CardTitle>
              </CardHeader>
              <CardContent className="p-4 pt-2">
                <CardDescription>Pick your own upgrade/downgrade rules.</CardDescription>
              </CardContent>
            </Card>
          </div>

          {strategy === 'custom' && (
            <Card className="border-border/60 bg-card shadow-sm">
              <CardHeader className="pb-3">
                <CardTitle className="text-base">Custom rules</CardTitle>
                <CardDescription>Toggle what this recommendation engine is allowed to change.</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  {[
                    { key: 'upgradeE1ToE3', label: 'Upgrade E1 → E3 (baseline security)', hint: 'Improves baseline security posture for low-tier users.' },
                    { key: 'upgradeE3ToE5', label: 'Upgrade E3 → E5 (max security)', hint: 'Adds advanced security/compliance capabilities.' },
                    { key: 'downgradeE5ToE3NonCore', label: 'Downgrade E5 → E3 for non-core depts', hint: 'Keeps E5 for IT/Engineering; reduces spend elsewhere.' },
                    { key: 'removeVisio', label: 'Remove Visio Plan 2', hint: 'Eliminates a common high-cost add-on.' },
                    { key: 'removeProject', label: 'Remove Project Plan 3', hint: 'Eliminates a common high-cost add-on.' },
                    { key: 'addCopilotEngineering', label: 'Add GitHub Copilot for Engineering', hint: 'Productivity add-on for engineering teams.' },
                  ].map((r) => (
                    <button
                      key={r.key}
                      type="button"
                      onClick={() =>
                        setCustomRules((prev) => ({
                          ...prev,
                          [r.key]: !(prev as any)[r.key],
                        }))
                      }
                      className={`text-left rounded-lg border p-3 transition-colors hover:bg-muted/30 ${
                        (customRules as any)[r.key] ? 'border-primary/30 bg-primary/5' : 'border-border/60'
                      }`}
                      data-testid={`toggle-custom-${r.key}`}
                    >
                      <div className="flex items-start justify-between gap-3">
                        <div>
                          <div className="font-medium text-sm">{r.label}</div>
                          <div className="text-xs text-muted-foreground mt-1">{r.hint}</div>
                        </div>
                        <div className={`mt-0.5 h-5 w-9 rounded-full border border-border/60 flex items-center px-0.5 ${
                          (customRules as any)[r.key] ? 'bg-primary/20' : 'bg-muted'
                        }`}>
                          <div className={`h-4 w-4 rounded-full bg-background shadow-sm transition-transform ${
                            (customRules as any)[r.key] ? 'translate-x-4' : 'translate-x-0'
                          }`} />
                        </div>
                      </div>
                    </button>
                  ))}
                </div>
              </CardContent>
            </Card>
          )}
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
                <Button variant="outline" size="icon" title="Filter" data-testid="button-filter">
                  <Filter className="h-4 w-4" />
                </Button>
              </div>
            </div>
          </CardHeader>
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
                      No users found matching "{searchTerm}"
                    </TableCell>
                  </TableRow>
                ) : (
                  filteredData.map((user) => {
                    const originalUser = data.find(u => u.id === user.id)!;
                    const isModified = strategy !== 'current' && JSON.stringify(originalUser.licenses) !== JSON.stringify(user.licenses);
                    
                    return (
                      <TableRow 
                        key={user.id} 
                        className={`transition-colors group ${isModified ? 'bg-primary/5 hover:bg-primary/10' : 'hover:bg-muted/20'}`} 
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
                          <div className="flex flex-col gap-2">
                            {isModified ? (
                              <div className="flex flex-col gap-1.5">
                                <div className="flex flex-wrap gap-1.5 opacity-50 line-through">
                                  {originalUser.licenses.map((license, i) => (
                                    <span key={i} className="text-xs">{license}</span>
                                  ))}
                                </div>
                                <div className="flex items-center gap-1.5">
                                  <ArrowRight className="h-3 w-3 text-primary" />
                                  <div className="flex flex-wrap gap-1.5">
                                    {user.licenses.map((license, i) => (
                                      <Badge key={i} className="text-xs bg-primary/20 text-primary border-primary/20 hover:bg-primary/30">
                                        {license}
                                      </Badge>
                                    ))}
                                  </div>
                                </div>
                              </div>
                            ) : (
                              <div className="flex flex-wrap gap-1.5">
                                {user.licenses.map((license, i) => (
                                  <Badge key={i} variant="outline" className="text-xs border-border/60">
                                    {license}
                                  </Badge>
                                ))}
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
      </main>
      <footer className="py-4 text-center text-xs text-muted-foreground border-t border-border/30">
        &copy; 2026 Cavaridge, LLC. All rights reserved.
      </footer>
    </div>
  );
}
