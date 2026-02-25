import { useState, useEffect, useMemo } from "react";
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
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";
import * as XLSX from "xlsx";

// Mock data generation based on the provided excel instructions
const mockData = [
  { id: "1", displayName: "Alex Johnson", upn: "alex.j@company.com", department: "Engineering", licenses: ["Microsoft 365 E5", "Visio Plan 2"], usageGB: 45.2, maxGB: 100, cost: 57.00, status: "Active" },
  { id: "2", displayName: "Sarah Smith", upn: "sarah.s@company.com", department: "Marketing", licenses: ["Microsoft 365 E3"], usageGB: 82.5, maxGB: 100, cost: 36.00, status: "Warning" },
  { id: "3", displayName: "Michael Chen", upn: "michael.c@company.com", department: "Sales", licenses: ["Microsoft 365 E3", "Power BI Pro"], usageGB: 12.1, maxGB: 100, cost: 46.00, status: "Active" },
  { id: "4", displayName: "Emily Davis", upn: "emily.d@company.com", department: "HR", licenses: ["Office 365 E1"], usageGB: 4.8, maxGB: 50, cost: 10.00, status: "Active" },
  { id: "5", displayName: "James Wilson", upn: "james.w@company.com", department: "Engineering", licenses: ["Microsoft 365 E5", "GitHub Copilot"], usageGB: 95.1, maxGB: 100, cost: 77.00, status: "Critical" },
  { id: "6", displayName: "Jessica Taylor", upn: "jessica.t@company.com", department: "Finance", licenses: ["Microsoft 365 E5"], usageGB: 22.4, maxGB: 100, cost: 57.00, status: "Active" },
  { id: "7", displayName: "David Anderson", upn: "david.a@company.com", department: "IT", licenses: ["Microsoft 365 E5", "Project Plan 3"], usageGB: 68.9, maxGB: 100, cost: 87.00, status: "Active" },
];

type Strategy = "current" | "security" | "cost" | "balanced";

export default function Dashboard() {
  const [isSyncing, setIsSyncing] = useState(false);
  const [data, setData] = useState<typeof mockData>([]);
  const [searchTerm, setSearchTerm] = useState("");
  const [strategy, setStrategy] = useState<Strategy>("current");

  useEffect(() => {
    // Initial mock load
    setIsSyncing(true);
    setTimeout(() => {
      setData(mockData);
      setIsSyncing(false);
    }, 1500);
  }, []);

  const handleSync = () => {
    setIsSyncing(true);
    setTimeout(() => {
      setData([...mockData].sort(() => Math.random() - 0.5)); // shuffle to simulate refresh
      setIsSyncing(false);
    }, 2000);
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
    
    return data.map(user => {
      let newLicenses = [...user.licenses];
      let newCost = user.cost;
      
      if (strategy === "security") {
        // Upgrade E1 to E3, E3 to E5
        if (newLicenses.includes("Office 365 E1")) {
          newLicenses = newLicenses.filter(l => l !== "Office 365 E1");
          newLicenses.push("Microsoft 365 E3");
          newCost += 26;
        } else if (newLicenses.includes("Microsoft 365 E3")) {
          newLicenses = newLicenses.filter(l => l !== "Microsoft 365 E3");
          newLicenses.push("Microsoft 365 E5");
          newCost += 21;
        }
        // Add Copilot to Engineering
        if (user.department === "Engineering" && !newLicenses.includes("GitHub Copilot")) {
          newLicenses.push("GitHub Copilot");
          newCost += 20;
        }
      } else if (strategy === "cost") {
        // Downgrade E5 to E3 where not IT/Engineering
        if (newLicenses.includes("Microsoft 365 E5") && !["IT", "Engineering"].includes(user.department)) {
          newLicenses = newLicenses.filter(l => l !== "Microsoft 365 E5");
          newLicenses.push("Microsoft 365 E3");
          newCost -= 21;
        }
        // Remove add-ons for cost saving
        if (newLicenses.includes("Visio Plan 2")) {
          newLicenses = newLicenses.filter(l => l !== "Visio Plan 2");
          newCost -= 15; 
        }
        if (newLicenses.includes("Project Plan 3")) {
          newLicenses = newLicenses.filter(l => l !== "Project Plan 3");
          newCost -= 30;
        }
      } else if (strategy === "balanced") {
        // Downgrade non-core E5s, but upgrade E1s to E3 for baseline security
        if (newLicenses.includes("Microsoft 365 E5") && ["Marketing", "HR", "Sales"].includes(user.department)) {
           newLicenses = newLicenses.filter(l => l !== "Microsoft 365 E5");
           newLicenses.push("Microsoft 365 E3");
           newCost -= 21;
        }
        if (newLicenses.includes("Office 365 E1")) {
          newLicenses = newLicenses.filter(l => l !== "Office 365 E1");
          newLicenses.push("Microsoft 365 E3");
          newCost += 26;
        }
      }
      
      return { ...user, licenses: newLicenses, cost: newCost };
    });
  }, [data, strategy]);

  const filteredData = optimizedData.filter(item => 
    item.displayName.toLowerCase().includes(searchTerm.toLowerCase()) || 
    item.upn.toLowerCase().includes(searchTerm.toLowerCase()) ||
    item.department.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const baseTotalCost = data.reduce((acc, curr) => acc + curr.cost, 0);
  const totalCost = optimizedData.reduce((acc, curr) => acc + curr.cost, 0);
  const costDiff = totalCost - baseTotalCost;
  
  const totalUsers = optimizedData.length;
  const totalStorage = optimizedData.reduce((acc, curr) => acc + curr.usageGB, 0);

  return (
    <div className="min-h-screen bg-background flex flex-col font-sans text-foreground">
      {/* Top Navigation */}
      <header className="sticky top-0 z-10 bg-card/80 backdrop-blur-md border-b border-border px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-2">
          <div className="h-8 w-8 rounded-md bg-primary flex items-center justify-center text-primary-foreground font-bold">
            M
          </div>
          <h1 className="text-xl font-semibold tracking-tight">M365 Insights</h1>
        </div>
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 text-sm text-muted-foreground mr-4">
            <span className="relative flex h-2 w-2">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-green-400 opacity-75"></span>
              <span className="relative inline-flex rounded-full h-2 w-2 bg-green-500"></span>
            </span>
            Connected to Microsoft 365
          </div>
          <Button 
            variant="outline" 
            size="sm" 
            onClick={handleSync}
            disabled={isSyncing}
            className="gap-2"
            data-testid="button-sync"
          >
            <RefreshCcw className={`h-4 w-4 ${isSyncing ? 'animate-spin' : ''}`} />
            {isSyncing ? 'Syncing...' : 'Sync Graph Data'}
          </Button>
          <Button size="sm" className="gap-2" onClick={handleExportXlsx} data-testid="button-export">
            <Download className="h-4 w-4" />
            Export XLSX
          </Button>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 p-8 max-w-7xl mx-auto w-full space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
        
        <div className="flex flex-col gap-2">
          <h2 className="text-3xl font-display font-semibold">Usage & Licensing Dashboard</h2>
          <p className="text-muted-foreground">Automated merge of Active Users and Mailbox Usage reports with actionable insights.</p>
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
                  {strategy !== 'current' && (
                    <div className={`text-sm font-medium mb-1 ${costDiff > 0 ? 'text-amber-500' : 'text-green-500'}`}>
                      {costDiff > 0 ? '+' : ''}{costDiff.toFixed(2)} /mo
                    </div>
                  )}
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
          </div>
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
    </div>
  );
}
