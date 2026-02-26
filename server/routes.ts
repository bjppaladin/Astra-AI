import type { Express } from "express";
import { createServer, type Server } from "http";
import { storage } from "./storage";
import OpenAI from "openai";
import { z } from "zod";
import crypto from "crypto";
import {
  getAuthUrl,
  exchangeCodeForTokens,
  getCurrentUser,
  fetchM365Data,
  refreshAccessToken,
} from "./microsoft-graph";

const openrouter = new OpenAI({
  apiKey: process.env.OPENROUTER_API_KEY,
  baseURL: "https://openrouter.ai/api/v1",
});

function getRedirectUri(req: any): string {
  const protocol = req.headers["x-forwarded-proto"] || req.protocol;
  const host = req.headers["x-forwarded-host"] || req.headers.host;
  return `${protocol}://${host}/api/auth/microsoft/callback`;
}

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {

  app.get("/api/auth/microsoft/login", (req, res) => {
    try {
      const state = crypto.randomBytes(16).toString("hex");
      req.session.oauthState = state;
      const redirectUri = getRedirectUri(req);
      const authUrl = getAuthUrl(redirectUri, state);
      res.json({ authUrl });
    } catch (err: any) {
      console.error("Auth login error:", err);
      res.status(500).json({ error: err.message });
    }
  });

  app.get("/api/auth/microsoft/callback", async (req, res) => {
    try {
      const { code, state, error, error_description } = req.query;

      if (error) {
        return res.redirect(`/?auth_error=${encodeURIComponent(error_description as string || error as string)}`);
      }

      if (!code || typeof code !== "string") {
        return res.redirect("/?auth_error=Missing+authorization+code");
      }

      if (state !== req.session.oauthState) {
        return res.redirect("/?auth_error=Invalid+state+parameter");
      }

      req.session.oauthState = undefined;

      const redirectUri = getRedirectUri(req);
      const tokens = await exchangeCodeForTokens(code, redirectUri);

      const user = await getCurrentUser(tokens.accessToken);

      const sessionId = crypto.randomBytes(32).toString("hex");
      req.session.microsoftSessionId = sessionId;

      await storage.upsertMicrosoftToken({
        sessionId,
        accessToken: tokens.accessToken,
        refreshToken: tokens.refreshToken || null,
        expiresAt: tokens.expiresAt,
        tenantId: tokens.tenantId || null,
        userEmail: user.mail,
        userName: user.displayName,
      });

      res.redirect("/?auth_success=true");
    } catch (err: any) {
      console.error("OAuth callback error:", err);
      res.redirect(`/?auth_error=${encodeURIComponent(err.message)}`);
    }
  });

  app.get("/api/auth/microsoft/status", async (req, res) => {
    try {
      const sessionId = req.session.microsoftSessionId;
      if (!sessionId) {
        return res.json({ connected: false });
      }

      const token = await storage.getMicrosoftToken(sessionId);
      if (!token) {
        return res.json({ connected: false });
      }

      const isExpired = new Date() > new Date(token.expiresAt);

      res.json({
        connected: !isExpired,
        userEmail: token.userEmail,
        userName: token.userName,
        tenantId: token.tenantId,
        expiresAt: token.expiresAt,
      });
    } catch (err: any) {
      res.json({ connected: false });
    }
  });

  app.post("/api/auth/microsoft/logout", async (req, res) => {
    try {
      const sessionId = req.session.microsoftSessionId;
      if (sessionId) {
        await storage.deleteMicrosoftToken(sessionId);
      }
      req.session.microsoftSessionId = undefined;
      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  });

  app.get("/api/graph/sync", async (req, res) => {
    try {
      const sessionId = req.session.microsoftSessionId;
      if (!sessionId) {
        return res.status(401).json({ error: "Not connected to Microsoft 365. Please sign in first." });
      }

      let token = await storage.getMicrosoftToken(sessionId);
      if (!token) {
        return res.status(401).json({ error: "Session expired. Please sign in again." });
      }

      if (new Date() > new Date(token.expiresAt)) {
        if (token.refreshToken) {
          try {
            const refreshed = await refreshAccessToken(token.refreshToken);
            token = await storage.upsertMicrosoftToken({
              sessionId,
              accessToken: refreshed.accessToken,
              refreshToken: refreshed.refreshToken || token.refreshToken,
              expiresAt: refreshed.expiresAt,
              tenantId: token.tenantId,
              userEmail: token.userEmail,
              userName: token.userName,
            });
          } catch {
            return res.status(401).json({ error: "Token expired. Please sign in again." });
          }
        } else {
          return res.status(401).json({ error: "Token expired. Please sign in again." });
        }
      }

      const data = await fetchM365Data(token.accessToken);
      res.json({ users: data, source: "live", tenant: token.tenantId });
    } catch (err: any) {
      console.error("Graph sync error:", err);
      const message = err.message?.includes("Graph API error")
        ? "Failed to fetch data from Microsoft 365. Please check your permissions and try again."
        : err.message;
      res.status(500).json({ error: message });
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
