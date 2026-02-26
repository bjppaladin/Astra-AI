# M365 License & Usage Insights

## Overview
A full-stack web application for Microsoft 365 license and mailbox usage analysis. Users can either sign in via Microsoft OAuth consent screen (providing their own Azure AD app credentials) or upload CSV/XLSX exports from the M365 Admin Center. The app merges data, provides licensing optimization strategies, and generates AI-powered executive summaries for C-Suite presentations.

## Architecture
- **Frontend**: React + TypeScript + Vite + Tailwind CSS v4 + shadcn/ui components
- **Backend**: Express.js + TypeScript
- **Database**: PostgreSQL with Drizzle ORM
- **AI**: OpenRouter (anthropic/claude-sonnet-4 for executive summaries)
- **Auth**: Microsoft OAuth2 consent flow (user-provided Azure AD credentials, session-scoped)
- **Sessions**: express-session + connect-pg-simple (PostgreSQL-backed)
- **Routing**: wouter (frontend), Express (backend API)
- **Export**: xlsx library for Excel exports
- **File Upload**: multer for CSV/XLSX file parsing

## Key Features
1. **Microsoft OAuth Sign-In** — Users enter their own Azure AD app credentials (Client ID, Secret, Tenant ID), then go through the Microsoft consent screen to authorize. No server-stored credentials needed.
2. **CSV/XLSX Upload** — Alternative method: upload Active Users and Mailbox Usage reports exported from M365 Admin Center.
3. **Graph API Sync** — After OAuth, pulls licensed users and mailbox usage data via Microsoft Graph API.
4. **Smart Parsing** — Automatic column detection, preamble-row skipping, and license-to-cost mapping for 30+ M365 SKUs.
5. **Data Merging** — Joins user and mailbox data on UPN for a unified view.
6. **Dashboard** — Combined user directory with licenses + mailbox usage (demo mode with sample data).
7. **Strategy Selector** — Current / Security / Cost / Balanced / Custom optimization.
8. **Billing Commitment** — Monthly vs Annual cost comparison (0.85 multiplier for annual).
9. **XLSX Export** — Download combined report as Excel.
10. **Executive Summary** — AI-generated vCIO analysis streamed in real-time via SSE.

## Data Model (shared/schema.ts)
- `users` — Auth table (placeholder)
- `reports` — Saved report snapshots (strategy, commitment, user data as JSONB)
- `executiveSummaries` — AI-generated summaries linked to reports
- `user_sessions` — Session store (auto-created by connect-pg-simple)

## File Structure
```
client/src/
  pages/dashboard.tsx         — Main dashboard with KPIs, strategy selector, data table, OAuth + file upload
  pages/executive-summary.tsx — AI summary viewer with streaming
  lib/api.ts                  — API client functions (auth, upload, reports, sync)
server/
  db.ts                       — Database connection (Drizzle + pg)
  index.ts                    — Express app setup with session middleware
  microsoft-graph.ts          — Microsoft Graph API client (OAuth, user fetch, mailbox reports)
  routes.ts                   — API routes (auth, file upload, reports, summaries)
  storage.ts                  — Database storage interface
shared/
  schema.ts                   — Drizzle schema definitions
```

## API Routes
- `GET /api/auth/microsoft/status` — Check OAuth connection status
- `POST /api/auth/microsoft/login` — Start OAuth flow (accepts clientId, clientSecret, tenantId in body)
- `GET /api/auth/microsoft/callback` — OAuth redirect handler
- `POST /api/auth/microsoft/disconnect` — Clear OAuth session
- `GET /api/microsoft/sync` — Fetch users + mailbox data via Graph API
- `POST /api/upload/users` — Parse uploaded Active Users CSV/XLSX
- `POST /api/upload/mailbox` — Parse uploaded Mailbox Usage CSV/XLSX
- `GET /api/reports` — List saved reports
- `POST /api/reports` — Save a report snapshot
- `DELETE /api/reports/:id` — Delete a report
- `GET /api/reports/:id/summary` — Get saved executive summary
- `POST /api/reports/:id/summary` — Generate AI executive summary (SSE streaming)

## Environment Variables
- `DATABASE_URL` — PostgreSQL connection string
- `OPENROUTER_API_KEY` — OpenRouter API key for AI summaries

## How to Use
### Option A: Microsoft OAuth (Recommended)
1. Click "Import Data" in the app header
2. Enter your Azure AD Client ID, Client Secret, and Tenant ID
3. Click "Sign in with Microsoft" → go through the consent screen
4. Click "Sync Data" to pull users, licenses, and mailbox usage

### Option B: File Upload
1. Go to M365 Admin Center → Reports → Usage → Active Users → Export CSV
2. Go to M365 Admin Center → Reports → Usage → Email activity → Export CSV
3. Click "Import Data" in the app header
4. Upload Active Users file first, then Mailbox Usage file
5. The app automatically parses, maps licenses to costs, and merges the data
