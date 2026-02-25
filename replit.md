# M365 License & Usage Insights

## Overview
A full-stack web application for automated Microsoft 365 user/license management and mailbox usage reporting. Features licensing optimization strategies with AI-generated executive summaries for C-Suite presentations.

## Architecture
- **Frontend**: React + TypeScript + Vite + Tailwind CSS v4 + shadcn/ui components
- **Backend**: Express.js + TypeScript
- **Database**: PostgreSQL with Drizzle ORM
- **AI**: OpenAI via Replit AI Integrations (gpt-5.2 for executive summaries)
- **Routing**: wouter (frontend), Express (backend API)
- **Export**: xlsx library for Excel exports

## Key Features
1. **Dashboard** — Combined user directory with licenses + mailbox usage
2. **Strategy Selector** — Current / Security / Cost / Balanced / Custom optimization
3. **Billing Commitment** — Monthly vs Annual cost comparison
4. **XLSX Export** — Download combined report as Excel
5. **Executive Summary** — AI-generated vCIO analysis streamed in real-time

## Data Model (shared/schema.ts)
- `users` — Auth table (placeholder)
- `reports` — Saved report snapshots (strategy, commitment, user data as JSONB)
- `executiveSummaries` — AI-generated summaries linked to reports

## File Structure
```
client/src/
  pages/dashboard.tsx      — Main dashboard with KPIs, strategy selector, data table
  pages/executive-summary.tsx — AI summary viewer with streaming
  lib/api.ts               — API client functions
server/
  db.ts                    — Database connection (Drizzle + pg)
  routes.ts                — API routes (/api/reports, /api/reports/:id/summary)
  storage.ts               — Database storage interface
shared/
  schema.ts                — Drizzle schema definitions
```

## API Routes
- `GET /api/reports` — List saved reports
- `POST /api/reports` — Save a report snapshot
- `DELETE /api/reports/:id` — Delete a report
- `GET /api/reports/:id/summary` — Get saved executive summary
- `POST /api/reports/:id/summary` — Generate AI executive summary (SSE streaming)

## Environment Variables
- `DATABASE_URL` — PostgreSQL connection string
- `AI_INTEGRATIONS_OPENAI_API_KEY` — OpenAI API key (via Replit AI Integrations)
- `AI_INTEGRATIONS_OPENAI_BASE_URL` — OpenAI base URL (via Replit AI Integrations)
