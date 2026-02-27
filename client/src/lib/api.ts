import { queryClient } from "./queryClient";

export async function getMicrosoftAuthStatus(): Promise<{
  configured: boolean;
  connected: boolean;
  user?: { displayName: string; email: string };
  tenantId?: string;
}> {
  const res = await fetch("/api/auth/microsoft/status");
  if (!res.ok) throw new Error("Failed to check auth status");
  return res.json();
}

export async function getMicrosoftLoginUrl(): Promise<string> {
  const res = await fetch("/api/auth/microsoft/login");
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.error || "Failed to start login");
  }
  const data = await res.json();
  return data.authUrl;
}

export async function disconnectMicrosoft(): Promise<void> {
  const res = await fetch("/api/auth/microsoft/disconnect", { method: "POST" });
  if (!res.ok) throw new Error("Failed to disconnect");
}

export async function syncMicrosoftData(): Promise<{
  users: any[];
  source: string;
  syncedAt: string;
}> {
  const res = await fetch("/api/microsoft/sync");
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.error || "Failed to sync data");
  }
  return res.json();
}

export async function uploadUsersFile(file: File): Promise<{
  users: any[];
  source: string;
  fileName: string;
  totalParsed: number;
  licensedUsers: number;
}> {
  const formData = new FormData();
  formData.append("file", file);
  const res = await fetch("/api/upload/users", {
    method: "POST",
    body: formData,
  });
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.error || "Failed to parse user file");
  }
  return res.json();
}

export async function uploadMailboxFile(file: File): Promise<{
  mailboxData: Record<string, { usageGB: number; maxGB: number }>;
  source: string;
  fileName: string;
  totalMailboxes: number;
}> {
  const formData = new FormData();
  formData.append("file", file);
  const res = await fetch("/api/upload/mailbox", {
    method: "POST",
    body: formData,
  });
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.error || "Failed to parse mailbox file");
  }
  return res.json();
}

export async function fetchSubscriptions(): Promise<{
  subscriptions: {
    skuId: string;
    skuPartNumber: string;
    displayName: string;
    costPerUser: number;
    enabled: number;
    consumed: number;
    available: number;
    capabilityStatus: string;
    appliesTo: string;
  }[];
  tenantId?: string;
}> {
  const res = await fetch("/api/microsoft/subscriptions");
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.error || "Failed to fetch subscriptions");
  }
  return res.json();
}

export async function saveReport(data: {
  name: string;
  strategy: string;
  commitment: string;
  userData: any[];
  customRules?: any;
}) {
  const res = await fetch("/api/reports", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data),
  });
  if (!res.ok) throw new Error("Failed to save report");
  queryClient.invalidateQueries({ queryKey: ["/api/reports"] });
  return res.json();
}

export async function fetchReports() {
  const res = await fetch("/api/reports");
  if (!res.ok) throw new Error("Failed to fetch reports");
  return res.json();
}

export async function deleteReport(id: number) {
  const res = await fetch(`/api/reports/${id}`, { method: "DELETE" });
  if (!res.ok) throw new Error("Failed to delete report");
  queryClient.invalidateQueries({ queryKey: ["/api/reports"] });
}

export interface UserActivity {
  exchangeActive: boolean;
  oneDriveActive: boolean;
  sharePointActive: boolean;
  teamsActive: boolean;
  yammerActive: boolean;
  skypeActive: boolean;
  exchangeLastDate: string | null;
  oneDriveLastDate: string | null;
  sharePointLastDate: string | null;
  teamsLastDate: string | null;
  yammerLastDate: string | null;
  skypeLastDate: string | null;
  activeServiceCount: number;
  totalServiceCount: number;
  daysSinceLastActivity: number | null;
}

export async function uploadActivityFile(file: File): Promise<{
  activityData: Record<string, UserActivity>;
  source: string;
  fileName: string;
  totalUsers: number;
}> {
  const formData = new FormData();
  formData.append("file", file);
  const res = await fetch("/api/upload/activity", {
    method: "POST",
    body: formData,
  });
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.error || "Failed to parse activity file");
  }
  return res.json();
}

export async function fetchSummary(reportId: number) {
  const res = await fetch(`/api/reports/${reportId}/summary`);
  if (!res.ok) return null;
  return res.json();
}

export async function fetchGreeting(): Promise<{
  greeting: {
    message: string;
    subtitle: string;
    loginCount: number;
    firstName: string;
  } | null;
}> {
  const res = await fetch("/api/user/greeting");
  if (!res.ok) return { greeting: null };
  return res.json();
}

export async function fetchNews(): Promise<{
  items: { title: string; link: string; date: string; summary: string }[];
  cachedAt?: string;
  stale?: boolean;
  error?: string;
}> {
  const res = await fetch("/api/insights/news");
  if (!res.ok) return { items: [] };
  return res.json();
}

export async function generateSummaryStream(
  reportId: number,
  payload: {
    costCurrent: number;
    costSecurity: number;
    costSaving: number;
    costBalanced: number;
    costCustom?: number;
    commitment: string;
    userData: any[];
  },
  onChunk: (text: string) => void,
  onDone: () => void,
  onError: (err: string) => void
) {
  const res = await fetch(`/api/reports/${reportId}/summary`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    onError("Failed to generate summary");
    return;
  }

  const reader = res.body?.getReader();
  if (!reader) { onError("No response body"); return; }

  const decoder = new TextDecoder();
  let buffer = "";

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split("\n");
    buffer = lines.pop() || "";

    for (const line of lines) {
      if (!line.startsWith("data: ")) continue;
      try {
        const event = JSON.parse(line.slice(6));
        if (event.content) onChunk(event.content);
        if (event.done) onDone();
        if (event.error) onError(event.error);
      } catch {}
    }
  }
}
