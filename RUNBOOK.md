# Astra — Operations Runbook v1.0

---

## Table of Contents

1. [System Overview](#1-system-overview)
2. [File Map](#2-file-map)
3. [Database Schema](#3-database-schema)
4. [API Routes](#4-api-routes)
5. [Authentication & Authorization](#5-authentication--authorization)
6. [Environment Variables & Secrets](#6-environment-variables--secrets)
7. [Frontend Navigation](#7-frontend-navigation)
8. [Operational Troubleshooting](#8-operational-troubleshooting)
9. [AI Pipeline](#9-ai-pipeline)
10. [Design System](#10-design-system)

---

## 1. System Overview

Astra is a full-stack Microsoft 365 license and mailbox usage analysis platform built by Cavaridge, LLC. Users sign in via Microsoft OAuth (multi-tenant, `/common` endpoint) or upload CSV/XLSX exports from the M365 Admin Center. The app merges user and mailbox data, provides usage-aware licensing optimization strategies with per-user recommendations, offers a comprehensive license feature comparison guide, and generates AI-powered executive briefings in a vCIO style with PDF/PNG export.

### Architecture

| Layer | Technology |
|-------|-----------|
| Frontend | React 18 + TypeScript + Vite + Tailwind CSS v4 + shadcn/ui |
| Backend | Express.js + TypeScript (tsx) |
| Database | PostgreSQL + Drizzle ORM |
| Auth | Multi-tenant Microsoft OAuth2 (`/common` endpoint) |
| Sessions | express-session + connect-pg-simple (PostgreSQL-backed) |
| AI | OpenRouter (anthropic/claude-sonnet-4) |
| Routing | wouter (frontend), Express (backend) |
| Export | xlsx (Excel), html2canvas + jsPDF (PDF/PNG, lazy-loaded) |
| File Upload | multer (CSV/XLSX parsing) |

---

## 2. File Map

### Server Files (`server/`)

| File | Lines | Purpose |
|------|-------|---------|
| `index.ts` | ~130 | Express app bootstrap, session middleware (trust proxy, secure cookies, PG store), Vite dev middleware, HTTP server |
| `routes.ts` | ~797 | All API route handlers: OAuth flow, Graph API sync, file upload/parsing, reports CRUD, AI summary generation (SSE) |
| `microsoft-graph.ts` | ~318 | Microsoft Graph API client: OAuth token exchange/refresh, user fetch, mailbox usage reports, active user detail, subscribedSkus, SKU-to-name/cost mapping |
| `storage.ts` | ~66 | IStorage interface + DatabaseStorage implementation (reports, summaries, tokens CRUD via Drizzle) |
| `db.ts` | ~13 | Drizzle + pg database connection pool |

### Client Files (`client/src/`)

| File | Lines | Purpose |
|------|-------|---------|
| `App.tsx` | ~33 | Root component, route registration (`/`, `/licenses`, `/report/:id/summary`), React Query + Tooltip providers |
| `pages/dashboard.tsx` | ~1436 | Main dashboard: KPIs, billing selector, tenant subscriptions panel, strategy selector with 5 modes, custom rules panel, combined user directory table with filters, OAuth + file upload panels |
| `pages/executive-summary.tsx` | ~598 | AI executive briefing viewer: SSE streaming with word count/timer, markdown renderer (tables, blockquotes, HR, emoji), PDF/PNG/print export |
| `pages/license-comparison.tsx` | ~446 | License comparison guide: up to 3 licenses side-by-side, 8 feature categories, URL query param pre-selection, grouped dropdowns |
| `pages/not-found.tsx` | ~21 | 404 fallback page |
| `lib/api.ts` | ~188 | API client functions: auth status/login/disconnect, sync, subscriptions, file upload, reports CRUD, AI summary streaming |
| `lib/license-data.ts` | ~643 | Comprehensive M365 license feature dataset: 19 licenses, 8 categories, 50+ features per license, helper functions |
| `lib/queryClient.ts` | ~57 | React Query client configuration |
| `hooks/use-toast.ts` | ~191 | Toast notification hook (shadcn/ui) |

### Shared Files (`shared/`)

| File | Lines | Purpose |
|------|-------|---------|
| `schema.ts` | ~79 | Drizzle schema: users, reports, executiveSummaries, microsoftTokens tables; Zod insert schemas; TypeScript types |

### Config Files (root)

| File | Lines | Purpose |
|------|-------|---------|
| `package.json` | ~115 | Dependencies, scripts (`dev`, `build`, `db:push`) |
| `tsconfig.json` | ~23 | TypeScript configuration with path aliases |
| `vite.config.ts` | ~51 | Vite config with React plugin, path aliases, proxy to Express backend |
| `drizzle.config.ts` | ~14 | Drizzle Kit config pointing to shared/schema.ts |
| `client/index.html` | ~26 | HTML shell with OG/Twitter meta tags, font imports |

---

## 3. Database Schema (4 tables + 1 auto-managed)

### `users`

| Column | Type | Constraints |
|--------|------|------------|
| `id` | varchar | PK, DEFAULT `gen_random_uuid()` |
| `username` | text | NOT NULL, UNIQUE |
| `password` | text | NOT NULL |

### `reports`

| Column | Type | Constraints |
|--------|------|------------|
| `id` | serial | PK |
| `name` | text | NOT NULL |
| `strategy` | text | NOT NULL, DEFAULT `'current'` |
| `commitment` | text | NOT NULL, DEFAULT `'monthly'` |
| `user_data` | jsonb | NOT NULL |
| `custom_rules` | jsonb | nullable |
| `created_at` | timestamp | NOT NULL, DEFAULT `CURRENT_TIMESTAMP` |

### `executive_summaries`

| Column | Type | Constraints |
|--------|------|------------|
| `id` | serial | PK |
| `report_id` | integer | NOT NULL, FK → `reports.id` ON DELETE CASCADE |
| `content` | text | NOT NULL |
| `cost_current` | real | NOT NULL |
| `cost_security` | real | NOT NULL |
| `cost_saving` | real | NOT NULL |
| `cost_balanced` | real | NOT NULL |
| `cost_custom` | real | nullable |
| `commitment` | text | NOT NULL |
| `created_at` | timestamp | NOT NULL, DEFAULT `CURRENT_TIMESTAMP` |

### `microsoft_tokens`

| Column | Type | Constraints |
|--------|------|------------|
| `id` | serial | PK |
| `session_id` | text | NOT NULL, UNIQUE |
| `access_token` | text | NOT NULL |
| `refresh_token` | text | nullable |
| `expires_at` | timestamp | NOT NULL |
| `tenant_id` | text | nullable |
| `user_email` | text | nullable |
| `user_name` | text | nullable |
| `created_at` | timestamp | NOT NULL, DEFAULT `CURRENT_TIMESTAMP` |

### `user_sessions` (auto-managed by connect-pg-simple)

| Column | Type | Constraints |
|--------|------|------------|
| `sid` | varchar | PK |
| `sess` | json | NOT NULL |
| `expire` | timestamp | NOT NULL, indexed |

---

## 4. API Routes

### 4.1 Authentication (Microsoft OAuth)

| Method | Path | Purpose | Auth |
|--------|------|---------|------|
| GET | `/api/auth/microsoft/status` | Check OAuth connection status (configured, connected, user, tenantId) | No |
| GET | `/api/auth/microsoft/login` | Start OAuth flow, returns auth URL with CSRF state param | No |
| GET | `/api/auth/microsoft/callback` | OAuth redirect handler — exchanges code for tokens, extracts tenant ID from JWT | No |
| POST | `/api/auth/microsoft/disconnect` | Clear OAuth session and token store | Session |

### 4.2 Microsoft Graph API

| Method | Path | Purpose | Auth |
|--------|------|---------|------|
| GET | `/api/microsoft/sync` | Fetch licensed users + mailbox usage via Graph API | OAuth session |
| GET | `/api/microsoft/report/active-users` | Fetch Office 365 Active User Detail report (30-day) | OAuth session |
| GET | `/api/microsoft/subscriptions` | Fetch tenant subscribed SKUs (license inventory) | OAuth session |

### 4.3 File Upload

| Method | Path | Purpose | Auth |
|--------|------|---------|------|
| POST | `/api/upload/users` | Parse uploaded Active Users CSV/XLSX, map licenses to costs | No |
| POST | `/api/upload/mailbox` | Parse uploaded Mailbox Usage CSV/XLSX, extract storage data | No |

### 4.4 Reports

| Method | Path | Purpose | Auth |
|--------|------|---------|------|
| GET | `/api/reports` | List all saved reports | No |
| GET | `/api/reports/:id` | Get a single report by ID | No |
| POST | `/api/reports` | Save a report snapshot (strategy, commitment, user data) | No |
| DELETE | `/api/reports/:id` | Delete a report (cascades to summary) | No |

### 4.5 Executive Summaries

| Method | Path | Purpose | Auth |
|--------|------|---------|------|
| GET | `/api/reports/:id/summary` | Get saved executive summary for a report | No |
| POST | `/api/reports/:id/summary` | Generate AI executive summary via SSE streaming | No |

---

## 5. Authentication & Authorization

- **Method**: Multi-tenant Microsoft OAuth2 via `https://login.microsoftonline.com/common/oauth2/v2.0`
- **Credentials**: Server-stored Azure AD app credentials (`MICROSOFT_CLIENT_ID`, `MICROSOFT_CLIENT_SECRET`) — end users just click "Sign in with Microsoft"
- **Session**: express-session with PostgreSQL store (connect-pg-simple)
  - `trust proxy` enabled for production (Replit reverse proxy)
  - `secure: true` always (HTTPS in production)
  - `saveUninitialized: true` to ensure session exists before OAuth redirect
  - Explicit `session.save()` before returning auth URL to prevent race conditions
- **Token storage**: In-memory `Map` keyed by session ID, with tokens also persisted to `microsoft_tokens` table
- **Token refresh**: Automatic refresh via `refreshAccessToken()` when tokens expire
- **CSRF protection**: Random state parameter stored in session, validated on callback
- **Scopes**: `User.Read`, `User.Read.All`, `Directory.Read.All`, `Reports.Read.All`
- **Protected routes**: Only `/api/microsoft/*` (sync, report, subscriptions) require an active OAuth session
- **Public routes**: File upload, reports CRUD, and summary generation do not require auth

---

## 6. Environment Variables & Secrets

| Variable | Required | Description |
|----------|----------|-------------|
| `DATABASE_URL` | Yes | PostgreSQL connection string (auto-provided by Replit) |
| `MICROSOFT_CLIENT_ID` | Yes | Azure AD Application (Client) ID |
| `MICROSOFT_CLIENT_SECRET` | Yes | Azure AD Client Secret |
| `MICROSOFT_TENANT_ID` | No | Azure AD Tenant ID (stored but unused — app uses `/common` endpoint) |
| `OPENROUTER_API_KEY` | Yes | OpenRouter API key for AI executive summary generation |

---

## 7. Frontend Navigation

| Path | Page | Auth Required | Description |
|------|------|--------------|-------------|
| `/` | Dashboard | No | Main dashboard with KPIs, strategy engine, user directory, upload/sync |
| `/licenses` | License Comparison Guide | No | Side-by-side feature comparison for up to 3 M365 licenses |
| `/report/:id/summary` | Executive Briefing | No | AI-generated vCIO summary viewer with PDF/PNG export |

All pages share a consistent header with the Astra logo ("A" icon), app name, and nav links (Dashboard, License Guide). Footer displays "© 2026 Cavaridge, LLC. All rights reserved."

---

## 8. Operational Troubleshooting

### 8.1 Authentication

| Symptom | Cause | Fix |
|---------|-------|-----|
| "Invalid OAuth state" error after Microsoft sign-in | Session not persisted before redirect due to async save | Fixed: `saveUninitialized: true` + explicit `session.save()` before returning auth URL |
| OAuth works in dev but fails in production | Missing `trust proxy` or `secure: false` behind Replit reverse proxy | Fixed: `trust proxy` enabled, `secure: true` always |
| Token expired errors during sync | Access token expired and refresh failed | Check `refreshAccessToken()` — may need user to re-authenticate |

### 8.2 Data & Upload

| Symptom | Cause | Fix |
|---------|-------|-----|
| CSV upload returns 0 licensed users | File has preamble rows or unexpected column names | Smart parser skips preamble rows and auto-detects columns — check server logs for parsing details |
| Mailbox data doesn't merge with users | UPN case mismatch between user and mailbox files | Merging normalizes to lowercase — check if UPNs actually match |

### 8.3 PDF/PNG Export

| Symptom | Cause | Fix |
|---------|-------|-----|
| Box cutoffs in exported PDF/PNG | CSS overflow hidden or fixed heights not flattened | Fixed: computed styles flattened to RGB, `overflow: visible` forced, `height: auto` forced, dimensions measured post-flatten |

### 8.4 Frontend

| Symptom | Cause | Fix |
|---------|-------|-----|
| Strategy cards show no impact preview | No data loaded (mock data hasn't initialized) | Wait for initial data load or upload/sync data first |
| License badge click doesn't navigate | Badge onClick missing on some rendering paths | Both modified and unmodified badge paths now have onClick handlers with `navigate()` |

---

## 9. AI Pipeline

| Setting | Value |
|---------|-------|
| Provider | OpenRouter (`https://openrouter.ai/api/v1`) |
| Model | `anthropic/claude-sonnet-4` |
| Temperature | 0.4 |
| Max tokens | 8,192 |
| Streaming | SSE (Server-Sent Events) |

The executive summary endpoint pre-computes analytics before sending to the AI:
- Department breakdowns with license distribution
- License tier distribution and cost analysis
- Mailbox usage analytics (quartiles, high-risk users)
- Risk signals (over-provisioned, under-provisioned, security gaps)
- Cost comparison across all 4 strategies

The AI generates an 8-section vCIO brief: Executive Summary, Current State Assessment, Security Strategy, Cost Strategy, Balanced Strategy, Risk Matrix, Implementation Roadmap, and Financial Summary.

---

## 10. Design System

| Element | Value |
|---------|-------|
| Display font | Plus Jakarta Sans |
| Body font | Inter |
| Primary color | `hsl(212, 100%, 48%)` / #0078d4 (Microsoft blue) |
| Component library | shadcn/ui (Radix UI primitives) |
| Icons | Lucide React |
| Charts | Recharts |
| Dark mode | Supported (HSL-based CSS variables) |
| Style | Enterprise Minimalist — soft shadows, glassmorphism header, muted borders |

---

*Astra — Operations Runbook v1.0*
*Updated February 2026 — Initial runbook covering full application*
