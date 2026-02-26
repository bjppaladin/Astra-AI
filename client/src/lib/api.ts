import { queryClient } from "./queryClient";

export async function getMicrosoftAuthStatus(): Promise<{
  connected: boolean;
  userEmail?: string;
  userName?: string;
  tenantId?: string;
}> {
  const res = await fetch("/api/auth/microsoft/status");
  if (!res.ok) return { connected: false };
  return res.json();
}

export async function getMicrosoftLoginUrl(): Promise<string> {
  const res = await fetch("/api/auth/microsoft/login");
  if (!res.ok) throw new Error("Failed to get login URL");
  const data = await res.json();
  return data.authUrl;
}

export async function microsoftLogout(): Promise<void> {
  await fetch("/api/auth/microsoft/logout", { method: "POST" });
  queryClient.invalidateQueries({ queryKey: ["/api/auth/microsoft/status"] });
}

export async function syncGraphData(): Promise<{ users: any[]; source: string; tenant?: string }> {
  const res = await fetch("/api/graph/sync");
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.error || "Failed to sync data");
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

export async function fetchSummary(reportId: number) {
  const res = await fetch(`/api/reports/${reportId}/summary`);
  if (!res.ok) return null;
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
