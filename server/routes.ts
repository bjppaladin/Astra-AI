import type { Express } from "express";
import { createServer, type Server } from "http";
import { storage } from "./storage";
import OpenAI from "openai";
import { z } from "zod";
import multer from "multer";
import * as XLSX from "xlsx";
import crypto from "crypto";
import {
  isOAuthConfigured,
  getAuthUrl,
  exchangeCodeForTokens,
  refreshAccessToken,
  getCurrentUser,
  fetchM365Data,
  fetchActiveUserDetailReport,
} from "./microsoft-graph";

const openrouter = new OpenAI({
  apiKey: process.env.OPENROUTER_API_KEY,
  baseURL: "https://openrouter.ai/api/v1",
});

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

const SKU_COST_MAP: Record<string, { name: string; cost: number }> = {
  "SPE_E5": { name: "Microsoft 365 E5", cost: 57.00 },
  "MICROSOFT 365 E5": { name: "Microsoft 365 E5", cost: 57.00 },
  "SPE_E3": { name: "Microsoft 365 E3", cost: 36.00 },
  "MICROSOFT 365 E3": { name: "Microsoft 365 E3", cost: 36.00 },
  "STANDARDPACK": { name: "Office 365 E1", cost: 10.00 },
  "OFFICE 365 E1": { name: "Office 365 E1", cost: 10.00 },
  "SPE_F1": { name: "Microsoft 365 F1", cost: 2.25 },
  "MICROSOFT 365 F1": { name: "Microsoft 365 F1", cost: 2.25 },
  "ENTERPRISEPREMIUM": { name: "Office 365 E5", cost: 38.00 },
  "OFFICE 365 E5": { name: "Office 365 E5", cost: 38.00 },
  "ENTERPRISEPACK": { name: "Office 365 E3", cost: 23.00 },
  "OFFICE 365 E3": { name: "Office 365 E3", cost: 23.00 },
  "VISIOCLIENT": { name: "Visio Plan 2", cost: 15.00 },
  "VISIO PLAN 2": { name: "Visio Plan 2", cost: 15.00 },
  "VISIO ONLINE PLAN 2": { name: "Visio Plan 2", cost: 15.00 },
  "PROJECTPREMIUM": { name: "Project Plan 5", cost: 55.00 },
  "PROJECT PLAN 5": { name: "Project Plan 5", cost: 55.00 },
  "PROJECTPROFESSIONAL": { name: "Project Plan 3", cost: 30.00 },
  "PROJECT PLAN 3": { name: "Project Plan 3", cost: 30.00 },
  "POWER_BI_PRO": { name: "Power BI Pro", cost: 10.00 },
  "POWER BI PRO": { name: "Power BI Pro", cost: 10.00 },
  "POWER_BI_PREMIUM_PER_USER": { name: "Power BI Premium Per User", cost: 20.00 },
  "MICROSOFT 365 COPILOT": { name: "Microsoft 365 Copilot", cost: 30.00 },
  "MICROSOFT_365_COPILOT": { name: "Microsoft 365 Copilot", cost: 30.00 },
  "GITHUB COPILOT": { name: "GitHub Copilot", cost: 20.00 },
  "EXCHANGESTANDARD": { name: "Exchange Online Plan 1", cost: 4.00 },
  "EXCHANGE ONLINE (PLAN 1)": { name: "Exchange Online Plan 1", cost: 4.00 },
  "EXCHANGEENTERPRISE": { name: "Exchange Online Plan 2", cost: 8.00 },
  "EXCHANGE ONLINE (PLAN 2)": { name: "Exchange Online Plan 2", cost: 8.00 },
  "MICROSOFT 365 BUSINESS BASIC": { name: "Microsoft 365 Business Basic", cost: 6.00 },
  "O365_BUSINESS_ESSENTIALS": { name: "Microsoft 365 Business Basic", cost: 6.00 },
  "MICROSOFT 365 BUSINESS STANDARD": { name: "Microsoft 365 Business Standard", cost: 12.50 },
  "O365_BUSINESS_PREMIUM": { name: "Microsoft 365 Business Standard", cost: 12.50 },
  "MICROSOFT 365 BUSINESS PREMIUM": { name: "Microsoft 365 Business Premium", cost: 22.00 },
  "SPB": { name: "Microsoft 365 Business Premium", cost: 22.00 },
  "TEAMS_EXPLORATORY": { name: "Teams Exploratory", cost: 0 },
  "FLOW_FREE": { name: "Power Automate Free", cost: 0 },
  "POWERAPPS_VIRAL": { name: "Power Apps Trial", cost: 0 },
  "STREAM": { name: "Microsoft Stream", cost: 0 },
};

function findLicenseInfo(licenseName: string): { name: string; cost: number } {
  const upper = licenseName.trim().toUpperCase();
  if (SKU_COST_MAP[upper]) return SKU_COST_MAP[upper];
  for (const [key, val] of Object.entries(SKU_COST_MAP)) {
    if (upper.includes(key) || key.includes(upper)) return val;
  }
  return { name: licenseName.trim(), cost: 0 };
}

function parseCSVContent(content: string): string[][] {
  const rows: string[][] = [];
  let current = "";
  let inQuotes = false;
  let row: string[] = [];

  for (let i = 0; i < content.length; i++) {
    const char = content[i];
    if (char === '"') {
      if (inQuotes && content[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      row.push(current.trim());
      current = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && content[i + 1] === "\n") i++;
      row.push(current.trim());
      if (row.some((c) => c)) rows.push(row);
      row = [];
      current = "";
    } else {
      current += char;
    }
  }
  if (current || row.length) {
    row.push(current.trim());
    if (row.some((c) => c)) rows.push(row);
  }
  return rows;
}

function findColumnIndex(headers: string[], ...candidates: string[]): number {
  for (const candidate of candidates) {
    const lower = candidate.toLowerCase();
    const idx = headers.findIndex((h) => h.toLowerCase().includes(lower));
    if (idx !== -1) return idx;
  }
  return -1;
}

function parseFileToRows(buffer: Buffer, filename: string): string[][] {
  const ext = filename.toLowerCase().split(".").pop();
  let rows: string[][];
  if (ext === "csv") {
    let content = buffer.toString("utf-8");
    if (content.charCodeAt(0) === 0xfeff) content = content.slice(1);
    rows = parseCSVContent(content);
  } else {
    const workbook = XLSX.read(buffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data: string[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    rows = data.map((row) => row.map((cell) => String(cell).trim()));
  }

  const headerKeywords = ["user principal name", "display name", "displayname", "upn", "assigned products", "email"];
  const headerIdx = rows.findIndex((row) =>
    row.some((cell) => headerKeywords.some((kw) => cell.toLowerCase().includes(kw)))
  );

  if (headerIdx > 0) {
    return rows.slice(headerIdx);
  }

  return rows;
}

const tokenStore = new Map<string, {
  accessToken: string;
  refreshToken?: string;
  expiresAt: Date;
  tenantId: string;
  userName?: string;
  userEmail?: string;
}>();

async function getValidToken(sessionId: string): Promise<string | null> {
  const stored = tokenStore.get(sessionId);
  if (!stored) return null;

  if (stored.expiresAt > new Date(Date.now() + 5 * 60 * 1000)) {
    return stored.accessToken;
  }

  if (stored.refreshToken) {
    try {
      const refreshed = await refreshAccessToken(stored.refreshToken);
      stored.accessToken = refreshed.accessToken;
      stored.refreshToken = refreshed.refreshToken || stored.refreshToken;
      stored.expiresAt = refreshed.expiresAt;
      tokenStore.set(sessionId, stored);
      return stored.accessToken;
    } catch {
      tokenStore.delete(sessionId);
    }
  }

  return null;
}

function getRedirectUri(req: any): string {
  const proto = req.headers["x-forwarded-proto"] || req.protocol;
  const host = req.headers["x-forwarded-host"] || req.headers.host;
  return `${proto}://${host}/api/auth/microsoft/callback`;
}

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {

  app.get("/api/auth/microsoft/status", async (req, res) => {
    const configured = isOAuthConfigured();
    const sessionId = req.session?.microsoftSessionId;

    if (!sessionId) {
      return res.json({ configured, connected: false });
    }

    const token = await getValidToken(sessionId);
    if (!token) {
      return res.json({ configured, connected: false });
    }

    const stored = tokenStore.get(sessionId);
    return res.json({
      configured,
      connected: true,
      user: { displayName: stored?.userName, email: stored?.userEmail },
      tenantId: stored?.tenantId,
    });
  });

  app.get("/api/auth/microsoft/login", (req, res) => {
    try {
      if (!isOAuthConfigured()) {
        return res.status(500).json({ error: "Microsoft OAuth is not configured on this server." });
      }

      const state = crypto.randomBytes(16).toString("hex");
      req.session.oauthState = state;

      const redirectUri = getRedirectUri(req);
      const authUrl = getAuthUrl(redirectUri, state);
      res.json({ authUrl });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.get("/api/auth/microsoft/callback", async (req, res) => {
    try {
      const { code, state, error, error_description } = req.query;

      if (error) {
        return res.redirect(`/?auth_error=${encodeURIComponent(String(error_description || error))}`);
      }

      if (!code || state !== req.session.oauthState) {
        return res.redirect("/?auth_error=Invalid+OAuth+state");
      }

      const redirectUri = getRedirectUri(req);
      const tokens = await exchangeCodeForTokens(String(code), redirectUri);

      const user = await getCurrentUser(tokens.accessToken);

      const sessionId = crypto.randomBytes(16).toString("hex");
      tokenStore.set(sessionId, {
        accessToken: tokens.accessToken,
        refreshToken: tokens.refreshToken,
        expiresAt: tokens.expiresAt,
        tenantId: tokens.tenantId,
        userName: user.displayName,
        userEmail: user.mail,
      });
      req.session.microsoftSessionId = sessionId;

      res.redirect("/?auth_success=true");
    } catch (err: any) {
      console.error("OAuth callback error:", err.message);
      res.redirect(`/?auth_error=${encodeURIComponent(err.message)}`);
    }
  });

  app.post("/api/auth/microsoft/disconnect", (req, res) => {
    const sessionId = req.session?.microsoftSessionId;
    if (sessionId) {
      tokenStore.delete(sessionId);
      delete req.session.microsoftSessionId;
    }
    res.json({ success: true });
  });

  app.get("/api/microsoft/sync", async (req, res) => {
    const sessionId = req.session?.microsoftSessionId;
    if (!sessionId) return res.status(401).json({ error: "Not connected to Microsoft 365" });

    const token = await getValidToken(sessionId);
    if (!token) return res.status(401).json({ error: "Session expired. Please reconnect." });

    const stored = tokenStore.get(sessionId);
    try {
      const users = await fetchM365Data(token);
      res.json({
        users,
        source: "microsoft365",
        tenantId: stored?.tenantId,
        syncedAt: new Date().toISOString(),
      });
    } catch (err: any) {
      console.error("M365 sync error:", err.message);
      res.status(500).json({ error: `Failed to sync: ${err.message}` });
    }
  });

  app.get("/api/microsoft/report/active-users", async (req, res) => {
    const sessionId = req.session?.microsoftSessionId;
    if (!sessionId) return res.status(401).json({ error: "Not connected to Microsoft 365" });

    const token = await getValidToken(sessionId);
    if (!token) return res.status(401).json({ error: "Session expired. Please reconnect." });

    const stored = tokenStore.get(sessionId);
    try {
      const report = await fetchActiveUserDetailReport(token);
      res.json({
        report,
        tenantId: stored?.tenantId,
        totalUsers: report.length,
        generatedAt: new Date().toISOString(),
      });
    } catch (err: any) {
      console.error("Report error:", err.message);
      res.status(500).json({ error: `Failed to get report: ${err.message}` });
    }
  });

  const logoUpload = multer({
    storage: multer.diskStorage({
      destination: "uploads/logos",
      filename: (_req, file, cb) => {
        const ext = file.originalname.split(".").pop();
        cb(null, `logo-${Date.now()}.${ext}`);
      },
    }),
    limits: { fileSize: 2 * 1024 * 1024 },
    fileFilter: (_req, file, cb) => {
      const allowed = ["image/png", "image/jpeg", "image/svg+xml"];
      cb(null, allowed.includes(file.mimetype));
    },
  });

  app.get("/api/settings/branding", async (req, res) => {
    try {
      const tenantId = Number(req.query.tenantId) || 1;
      const branding = await storage.getBranding(tenantId);
      res.json(branding || null);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.put("/api/settings/branding", async (req, res) => {
    try {
      const tenantId = Number(req.body.tenantId) || 1;
      const { tenantId: _tid, ...data } = req.body;
      const branding = await storage.upsertBranding(tenantId, data);
      res.json(branding);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post("/api/settings/branding/logo", logoUpload.single("logo"), (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ error: "No file uploaded or invalid file type. Accepted: PNG, JPG, SVG (max 2MB)." });
      const logoUrl = `/uploads/logos/${req.file.filename}`;
      res.json({ logoUrl });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.post("/api/upload/users", upload.single("file"), (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ error: "No file uploaded" });

      const rows = parseFileToRows(req.file.buffer, req.file.originalname);
      if (rows.length < 2) return res.status(400).json({ error: "File appears empty or has no data rows" });

      const headers = rows[0];

      const displayNameIdx = findColumnIndex(headers, "display name", "displayname", "full name", "name");
      const upnIdx = findColumnIndex(headers, "user principal name", "upn", "email", "username");
      const deptIdx = findColumnIndex(headers, "department", "dept");
      const licenseIdx = findColumnIndex(headers, "assigned products", "licenses", "assigned licenses", "products", "product");
      const deletedIdx = findColumnIndex(headers, "is deleted", "deleted");
      const enabledIdx = findColumnIndex(headers, "account enabled");

      if (displayNameIdx === -1 && upnIdx === -1) {
        return res.status(400).json({
          error: "Could not find user identity columns. Expected columns like 'Display Name' or 'User Principal Name'. Please upload the Active Users report from M365 Admin Center.",
          detectedColumns: headers,
        });
      }

      const users: any[] = [];

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length < 2) continue;

        if (deletedIdx >= 0) {
          const val = row[deletedIdx]?.toLowerCase();
          if (val === "true" || val === "yes") continue;
        }

        if (enabledIdx >= 0) {
          const val = row[enabledIdx]?.toLowerCase();
          if (val === "false" || val === "no") continue;
        }

        const displayName = displayNameIdx >= 0 ? row[displayNameIdx] : row[upnIdx]?.split("@")[0] || `User ${i}`;
        const upn = upnIdx >= 0 ? row[upnIdx] : `user${i}@unknown.com`;
        const department = deptIdx >= 0 ? row[deptIdx] || "Unassigned" : "Unassigned";

        const rawLicenses = licenseIdx >= 0 ? row[licenseIdx] : "";
        const licenseList = rawLicenses
          .split(/[+;,]/)
          .map((l: string) => l.trim())
          .filter((l: string) => l && l !== "-");

        const licenses: string[] = [];
        let cost = 0;
        for (const lic of licenseList) {
          const info = findLicenseInfo(lic);
          licenses.push(info.name);
          cost += info.cost;
        }

        if (licenses.length === 0) continue;

        users.push({
          id: String(i),
          displayName,
          upn,
          department,
          licenses,
          cost,
          usageGB: 0,
          maxGB: 50,
          status: "Active",
        });
      }

      if (users.length === 0) {
        return res.status(400).json({ error: "No licensed users found in the file. Make sure you're uploading the Active Users report." });
      }

      res.json({
        users,
        source: "uploaded",
        fileName: req.file.originalname,
        totalParsed: rows.length - 1,
        licensedUsers: users.length,
      });
    } catch (err: any) {
      console.error("User upload parse error:", err);
      res.status(400).json({ error: `Failed to parse file: ${err.message}` });
    }
  });

  app.post("/api/upload/mailbox", upload.single("file"), (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ error: "No file uploaded" });

      const rows = parseFileToRows(req.file.buffer, req.file.originalname);
      if (rows.length < 2) return res.status(400).json({ error: "File appears empty or has no data rows" });

      const headers = rows[0];

      const upnIdx = findColumnIndex(headers, "user principal name", "upn", "email", "owner upn");
      const storageIdx = findColumnIndex(headers, "storage used", "total item size", "mailbox size");
      const quotaIdx = findColumnIndex(headers, "prohibit send/receive quota", "issue warning quota", "prohibit send quota", "quota");

      if (upnIdx === -1) {
        return res.status(400).json({
          error: "Could not find 'User Principal Name' column. Please upload the Mailbox Usage report from M365 Admin Center.",
          detectedColumns: headers,
        });
      }

      const mailboxData: Record<string, { usageGB: number; maxGB: number }> = {};

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length < 2) continue;

        const upn = row[upnIdx]?.toLowerCase();
        if (!upn) continue;

        let storageBytes = storageIdx >= 0 ? parseFloat(row[storageIdx]) || 0 : 0;
        let quotaBytes = quotaIdx >= 0 ? parseFloat(row[quotaIdx]) || 0 : 0;

        let usageGB: number;
        let maxGB: number;

        if (storageBytes > 1_000_000) {
          usageGB = storageBytes / (1024 * 1024 * 1024);
          maxGB = quotaBytes > 0 ? quotaBytes / (1024 * 1024 * 1024) : 50;
        } else if (storageBytes > 1000) {
          usageGB = storageBytes / 1024;
          maxGB = quotaBytes > 0 ? quotaBytes / 1024 : 50;
        } else {
          usageGB = storageBytes;
          maxGB = quotaBytes > 0 ? quotaBytes : 50;
        }

        mailboxData[upn] = {
          usageGB: Math.round(usageGB * 10) / 10,
          maxGB: Math.round(maxGB),
        };
      }

      res.json({
        mailboxData,
        source: "uploaded",
        fileName: req.file.originalname,
        totalMailboxes: Object.keys(mailboxData).length,
      });
    } catch (err: any) {
      console.error("Mailbox upload parse error:", err);
      res.status(400).json({ error: `Failed to parse file: ${err.message}` });
    }
  });

  app.get("/api/reports", async (_req, res) => {
    const reports = await storage.getReports();
    res.json(reports);
  });

  app.get("/api/reports/:id", async (req, res) => {
    const report = await storage.getReport(Number(req.params.id));
    if (!report) return res.status(404).json({ error: "Report not found" });
    res.json(report);
  });

  app.post("/api/reports", async (req, res) => {
    try {
      const report = await storage.createReport(req.body);
      res.status(201).json(report);
    } catch (err: any) {
      res.status(400).json({ error: err.message });
    }
  });

  app.delete("/api/reports/:id", async (req, res) => {
    await storage.deleteReport(Number(req.params.id));
    res.status(204).send();
  });

  app.get("/api/reports/:id/summary", async (req, res) => {
    const summary = await storage.getExecutiveSummary(Number(req.params.id));
    if (!summary) return res.status(404).json({ error: "Summary not found" });
    res.json(summary);
  });

  app.post("/api/reports/:id/summary", async (req, res) => {
    try {
      const reportId = Number(req.params.id);
      const report = await storage.getReport(reportId);
      if (!report) return res.status(404).json({ error: "Report not found" });

      const { costCurrent, costSecurity, costSaving, costBalanced, costCustom, commitment, userData } = req.body;

      const commitmentLabel = commitment === "annual" ? "Annual Commitment" : "Monthly Commitment";

      const prompt = `You are a senior virtual CIO (vCIO) preparing an executive summary for a C-Suite audience about Microsoft 365 licensing optimization. You must be authoritative, data-driven, and persuasive. This document needs to withstand the scrutiny of executives who will challenge every recommendation.

Here is the data:

BILLING BASIS: ${commitmentLabel}
CURRENT MONTHLY SPEND: $${costCurrent.toFixed(2)}
OPTION 1 — MAXIMIZE SECURITY: $${costSecurity.toFixed(2)} (${costSecurity > costCurrent ? '+' : ''}$${(costSecurity - costCurrent).toFixed(2)}/mo delta)
OPTION 2 — MINIMIZE COST: $${costSaving.toFixed(2)} (${costSaving > costCurrent ? '+' : ''}$${(costSaving - costCurrent).toFixed(2)}/mo delta)
OPTION 3 — BALANCED APPROACH: $${costBalanced.toFixed(2)} (${costBalanced > costCurrent ? '+' : ''}$${(costBalanced - costCurrent).toFixed(2)}/mo delta)
${costCustom !== undefined ? `OPTION 4 — CUSTOM STRATEGY: $${costCustom.toFixed(2)} (${costCustom > costCurrent ? '+' : ''}$${(costCustom - costCurrent).toFixed(2)}/mo delta)` : ''}

USER DIRECTORY (${(userData as any[]).length} users):
${(userData as any[]).map((u: any) => `- ${u.displayName} (${u.department}): Current licenses: ${u.licenses.join(', ')}; Mailbox: ${u.usageGB}GB/${u.maxGB}GB; Current cost: $${u.cost}/mo`).join('\n')}

Write a polished executive summary in Markdown that includes:
1. **Executive Overview** — A 2-3 sentence summary of the current licensing posture and why action is needed.
2. **Cost Comparison Table** — A clear comparison of all options (Current State, Maximize Security, Minimize Cost, Balanced${costCustom !== undefined ? ', Custom' : ''}) showing monthly cost, annual projected cost, delta vs current, and a one-line rationale.
3. **Risk Assessment** — For each option, highlight key risks (security gaps, compliance exposure, productivity impact, budget impact). Be specific — reference actual user counts and license tiers from the data.
4. **Recommendation** — Your professional recommendation as vCIO with a clear rationale. Consider the balance of security posture, cost efficiency, and operational impact. Be decisive.
5. **Implementation Roadmap** — A phased approach (30/60/90 day) for executing the recommended strategy.
6. **Next Steps** — 3-4 concrete action items for leadership.

Write with confidence and authority. Use precise dollar figures. Reference specific license tiers (E1, E3, E5) and their security implications. Do not hedge excessively — executives want clear direction.`;

      res.setHeader("Content-Type", "text/event-stream");
      res.setHeader("Cache-Control", "no-cache");
      res.setHeader("Connection", "keep-alive");

      const stream = await openrouter.chat.completions.create({
        model: "anthropic/claude-sonnet-4",
        messages: [{ role: "user", content: prompt }],
        stream: true,
        max_tokens: 4096,
      });

      let fullContent = "";

      for await (const chunk of stream) {
        const content = chunk.choices[0]?.delta?.content || "";
        if (content) {
          fullContent += content;
          res.write(`data: ${JSON.stringify({ content })}\n\n`);
        }
      }

      const summary = await storage.createExecutiveSummary({
        reportId,
        content: fullContent,
        costCurrent,
        costSecurity,
        costSaving,
        costBalanced,
        costCustom: costCustom ?? null,
        commitment,
      });

      res.write(`data: ${JSON.stringify({ done: true, summaryId: summary.id })}\n\n`);
      res.end();
    } catch (err: any) {
      console.error("Error generating summary:", err);
      if (res.headersSent) {
        res.write(`data: ${JSON.stringify({ error: err.message })}\n\n`);
        res.end();
      } else {
        res.status(500).json({ error: err.message });
      }
    }
  });

  return httpServer;
}
