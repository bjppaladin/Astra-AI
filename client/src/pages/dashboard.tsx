import { useState, useEffect } from "react";
import { 
  RefreshCcw, 
  Download, 
  Users, 
  Database, 
  CreditCard,
  Search,
  Filter,
  CheckCircle2,
  AlertCircle
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";

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

export default function Dashboard() {
  const [isSyncing, setIsSyncing] = useState(false);
  const [data, setData] = useState<typeof mockData>([]);
  const [searchTerm, setSearchTerm] = useState("");

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

  const filteredData = data.filter(item => 
    item.displayName.toLowerCase().includes(searchTerm.toLowerCase()) || 
    item.upn.toLowerCase().includes(searchTerm.toLowerCase()) ||
    item.department.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const totalCost = data.reduce((acc, curr) => acc + curr.cost, 0);
  const totalUsers = data.length;
  const totalStorage = data.reduce((acc, curr) => acc + curr.usageGB, 0);

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
          <Button size="sm" className="gap-2" data-testid="button-export">
            <Download className="h-4 w-4" />
            Export Combined CSV
          </Button>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 p-8 max-w-7xl mx-auto w-full space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
        
        <div className="flex flex-col gap-2">
          <h2 className="text-3xl font-display font-semibold">Usage & Licensing Dashboard</h2>
          <p className="text-muted-foreground">Automated merge of Active Users and Mailbox Usage reports.</p>
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

          <Card className="shadow-sm border-border/50">
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium text-muted-foreground">Est. Monthly License Cost</CardTitle>
              <CreditCard className="h-4 w-4 text-muted-foreground" />
            </CardHeader>
            <CardContent>
               {isSyncing ? (
                <Skeleton className="h-8 w-28" />
              ) : (
                <div className="text-3xl font-bold font-display text-primary" data-testid="text-total-cost">${totalCost.toFixed(2)}</div>
              )}
            </CardContent>
          </Card>
        </div>

        {/* Data Table */}
        <Card className="shadow-md border-border/50 overflow-hidden">
          <CardHeader className="border-b border-border/50 bg-muted/20 pb-4">
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
              <div>
                <CardTitle>Combined User Directory</CardTitle>
                <CardDescription>Joined on UPN from Entra ID and Exchange Online.</CardDescription>
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
                  filteredData.map((user) => (
                    <TableRow key={user.id} className="hover:bg-muted/20 transition-colors group" data-testid={`row-user-${user.id}`}>
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
                        <div className="flex flex-wrap gap-1.5">
                          {user.licenses.map((license, i) => (
                            <Badge key={i} variant="outline" className="text-xs border-border/60">
                              {license}
                            </Badge>
                          ))}
                        </div>
                      </TableCell>
                      <TableCell className="text-right">
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
                      <TableCell className="text-right font-medium">
                        ${user.cost.toFixed(2)}
                      </TableCell>
                    </TableRow>
                  ))
                )}
              </TableBody>
            </Table>
          </div>
        </Card>
      </main>
    </div>
  );
}
