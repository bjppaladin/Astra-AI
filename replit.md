# M365 License & Usage Insights

## Overview
A full-stack web application for automated Microsoft 365 user/license management and mailbox usage reporting. Users import CSV/XLSX exports from the M365 Admin Center, and the app merges the data, provides licensing optimization strategies, and generates AI-powered executive summaries for C-Suite presentations.

## Architecture
- **Frontend**: React + TypeScript + Vite + Tailwind CSS v4 + shadcn/ui components
- **Backend**: Express.js + TypeScript
- **Database**: PostgreSQL with Drizzle ORM
- **AI**: OpenRouter (anthropic/claude-sonnet-4 for executive summaries)
- **Routing**: wouter (frontend), Express (backend API)
- **Export**: xlsx library for Excel exports
- **File Upload**: multer for CSV/XLSX file parsing

## Key Features
1. **Data Import** — Upload Active Users and Mailbox Usage CSV/XLSX exports from M365 Admin Center
2. **Smart Parsing** — Automatic column detection and license-to-cost mapping for all M365 SKUs
3. **Data Merging** — Joins user and mailbox data on UPN for a unified view
4. **Dashboard** — Combined user directory with licenses + mailbox usage (demo mode with sample data)
5. **Strategy Selector** — Current / Security / Cost / Balanced / Custom optimization
6. **Billing Commitment** — Monthly vs Annual cost comparison
7. **XLSX Export** — Download combined report as Excel
8. **Executive Summary** — AI-generated vCIO analysis streamed in real-time

## Data Model (shared/schema.ts)
- `users` — Auth table (placeholder)
- `reports` — Saved report snapshots (strategy, commitment, user data as JSONB)
- `executiveSummaries` — AI-generated summaries linked to reports

## File Structure
```
client/src/
  pages/dashboard.tsx        — Main dashboard with KPIs, strategy selector, data table, file upload
  pages/executive-summary.tsx — AI summary viewer with streaming
  lib/api.ts                 — API client functions (upload, reports)
server/
  db.ts                      — Database connection (Drizzle + pg)
  index.ts                   — Express app setup
  routes.ts                  — API routes (file upload, reports, summaries)
  storage.ts                 — Database storage interface
shared/
  schema.ts                  — Drizzle schema definitions
```

## API Routes
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

## How to Import Real Data
1. Go to M365 Admin Center → Reports → Usage → Active Users → Export CSV
2. Go to M365 Admin Center → Reports → Usage → Email activity → Export CSV
3. Click "Import M365 Data" in the app header
4. Upload Active Users file first, then Mailbox Usage file
5. The app automatically parses, maps licenses to costs, and merges the data
