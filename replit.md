# M365 License & Usage Insights

## Overview
A full-stack web application for automated Microsoft 365 user/license management and mailbox usage reporting. Supports live connection to Microsoft 365 via OAuth2, with licensing optimization strategies and AI-generated executive summaries for C-Suite presentations.

## Architecture
- **Frontend**: React + TypeScript + Vite + Tailwind CSS v4 + shadcn/ui components
- **Backend**: Express.js + TypeScript
- **Database**: PostgreSQL with Drizzle ORM
- **AI**: OpenRouter (anthropic/claude-sonnet-4 for executive summaries)
- **Auth**: Microsoft Entra ID OAuth2 (authorization code flow with refresh tokens)
- **Routing**: wouter (frontend), Express (backend API)
- **Export**: xlsx library for Excel exports
- **Sessions**: express-session with connect-pg-simple (PostgreSQL-backed)

## Key Features
1. **Microsoft 365 Login** — OAuth2 sign-in to connect live M365 tenant data
2. **Live Graph API Sync** — Pull real users, licenses, and mailbox usage from Microsoft Graph
3. **Dashboard** — Combined user directory with licenses + mailbox usage (demo mode when not connected)
4. **Strategy Selector** — Current / Security / Cost / Balanced / Custom optimization
5. **Billing Commitment** — Monthly vs Annual cost comparison
6. **XLSX Export** — Download combined report as Excel
7. **Executive Summary** — AI-generated vCIO analysis streamed in real-time

## Data Model (shared/schema.ts)
- `users` — Auth table (placeholder)
- `reports` — Saved report snapshots (strategy, commitment, user data as JSONB)
- `executiveSummaries` — AI-generated summaries linked to reports
- `microsoftTokens` — OAuth access/refresh tokens per session
- `user_sessions` — Express session store (auto-created by connect-pg-simple)

## File Structure
```
client/src/
  pages/dashboard.tsx        — Main dashboard with KPIs, strategy selector, data table, M365 login
  pages/executive-summary.tsx — AI summary viewer with streaming
  lib/api.ts                 — API client functions (auth, sync, reports)
server/
  db.ts                      — Database connection (Drizzle + pg)
  index.ts                   — Express app setup with session middleware
  routes.ts                  — API routes (auth, graph sync, reports, summaries)
  storage.ts                 — Database storage interface
  microsoft-graph.ts         — Microsoft Graph API service (OAuth, users, licenses, mailbox)
shared/
  schema.ts                  — Drizzle schema definitions
```

## API Routes
- `GET /api/auth/microsoft/login` — Get Microsoft OAuth login URL
- `GET /api/auth/microsoft/callback` — OAuth callback handler
- `GET /api/auth/microsoft/status` — Check connection status
- `POST /api/auth/microsoft/logout` — Disconnect Microsoft account
- `GET /api/graph/sync` — Pull live data from Microsoft Graph
- `GET /api/reports` — List saved reports
- `POST /api/reports` — Save a report snapshot
- `DELETE /api/reports/:id` — Delete a report
- `GET /api/reports/:id/summary` — Get saved executive summary
- `POST /api/reports/:id/summary` — Generate AI executive summary (SSE streaming)

## Environment Variables
- `DATABASE_URL` — PostgreSQL connection string
- `OPENROUTER_API_KEY` — OpenRouter API key for AI summaries
- `MICROSOFT_CLIENT_ID` — Azure AD app registration client ID
- `MICROSOFT_CLIENT_SECRET` — Azure AD app registration client secret
- `MICROSOFT_TENANT_ID` — Azure AD tenant ID (optional, defaults to "common" for multi-tenant)
- `SESSION_SECRET` — Express session secret (auto-generated if not set)

## Azure App Registration Setup
1. Go to Azure Portal > Entra ID > App Registrations > New Registration
2. Set redirect URI to: `https://<your-domain>/api/auth/microsoft/callback`
3. Under API Permissions, add Microsoft Graph delegated permissions:
   - User.Read.All
   - Reports.Read.All
   - Organization.Read.All
4. Grant admin consent for the permissions
5. Under Certificates & Secrets, create a client secret
6. Copy Client ID, Client Secret, and Tenant ID into environment variables
