# Astra — M365 License & Usage Insights

## Overview
Astra is a full-stack web application for Microsoft 365 license and mailbox usage analysis. Built by Cavaridge, LLC. Users from any M365 organization sign in via the Microsoft OAuth consent screen (one-click "Sign in with Microsoft" — no Azure setup required for end users). Alternatively, users can upload CSV/XLSX exports from the M365 Admin Center. The app merges data, provides usage-aware licensing optimization strategies with per-user recommendations, and generates AI-powered executive summaries in a vCIO style with PDF/PNG export.

## Architecture
- **Frontend**: React + TypeScript + Vite + Tailwind CSS v4 + shadcn/ui components
- **Backend**: Express.js + TypeScript
- **Database**: PostgreSQL with Drizzle ORM
- **AI**: OpenRouter (anthropic/claude-sonnet-4, temp 0.4, max_tokens 8192)
- **Auth**: Multi-tenant Microsoft OAuth2 via `https://login.microsoftonline.com/common` — server-stored app credentials, users just click sign-in
- **Sessions**: express-session + connect-pg-simple (PostgreSQL-backed), `trust proxy` enabled, `secure: true` always, `saveUninitialized: true`, explicit `session.save()` before returning auth URL
- **Routing**: wouter (frontend), Express (backend API)
- **Export**: xlsx for Excel, html2canvas + jsPDF for PDF/PNG (lazy-loaded)
- **File Upload**: multer for CSV/XLSX file parsing
- **Design**: Enterprise Minimalist — Plus Jakarta Sans (display) + Inter (body), Microsoft blue (#0078d4), shadcn/ui

## Key Features
1. **Microsoft OAuth Sign-In** — One-click "Sign in with Microsoft" for any M365 tenant. No Azure setup for end users.
2. **CSV/XLSX Upload** — Alternative: upload Active Users and Mailbox Usage reports from M365 Admin Center.
3. **Graph API Sync** — Pulls licensed users, SKU mappings, and mailbox usage via Microsoft Graph API.
4. **Active User Detail Report** — Dedicated endpoint for Office 365 Active User Detail report (30-day).
5. **Smart Parsing** — Automatic column detection, preamble-row skipping, and license-to-cost mapping for 100+ M365 SKUs (including Enterprise, Business, Frontline, Security, EMS, telephony, and add-on SKUs). `findLicenseInfo` normalizes underscores/spaces for robust matching.
6. **Data Merging** — Joins user and mailbox data on UPN for a unified view.
7. **Dashboard** — Combined user directory with licenses + mailbox usage (demo mode with sample data). Filterable by department, status, and modification state. License badges are clickable and link to the License Comparison Guide.
8. **Usage-Aware Strategy Engine** — Current / Security / Cost / Balanced / Custom optimization. Per-user analysis based on mailbox usage ratios, department classification, license tier, and add-on relevance. Supports Enterprise (E1/E3/E5), Business (Basic/Standard/Premium), and Frontline (F1/F3) tiers with tier-specific upgrade/downgrade rules. Includes redundant add-on detection (Defender, OneDrive, Intune, Entra ID included in suites). Each recommendation includes specific reasoning. Costs recomputed from final license set (not deltas). Security-dept awareness (IT, Engineering, Compliance, Security, InfoSec). Custom mode has grouped rules (Upgrades, Downgrades & Savings, Redundant Add-on Removal) with usage threshold slider. Strategy cards preview impact (users affected, net cost delta, upgrade/downgrade counts). Custom rules: `upgradeUnderprovisioned`, `upgradeBasicToStandard`, `upgradeToE5Security`, `upgradeToBizPremiumSecurity`, `downgradeUnderutilizedE5`, `downgradeOverprovisionedE3`, `downgradeUnderutilizedBizPremium`, `downgradeBizStandardToBasic`, `removeUnusedAddons`, `removeRedundantAddons`, `consolidateOverlap`, `addCopilotPowerUsers`, `usageThreshold`.
9. **Billing Commitment** — Monthly vs Annual cost comparison (0.85 multiplier for annual). Defaults to Annual.
10. **XLSX Export** — Download combined report as Excel.
11. **Tenant Subscriptions** — Collapsible panel showing live M365 subscription data via Graph API (`subscribedSkus`). Displays subscription name, status, purchased/assigned/available counts, utilization bar, cost per user, and total monthly spend with a totals row. Auto-loads when connected, manual refresh button.
12. **License Comparison Guide** — Dedicated `/licenses` page with side-by-side feature comparison for up to 3 M365 licenses. Comprehensive feature dataset (19 licenses) covering 8 categories: Core Apps, Email & Calendar, Communication & Collaboration, Storage, Security & Identity, Compliance & Data Governance, Automation & Development, AI & Advanced. License badges in the dashboard are clickable and navigate to the comparison page with that license pre-selected via URL query params. Supports suite licenses (E1, E3, E5, F1, Business Basic/Standard/Premium), standalone (Exchange Online Plan 1/2), and add-ons (Copilot, Visio, Project, Power BI, GitHub Copilot).
13. **Executive Briefing** — Comprehensive AI-generated vCIO analysis (8 sections: executive summary, current state assessment, strategy deep-dives, risk matrix with severity ratings, implementation roadmap, financial summary, next steps). Pre-computes dept breakdowns, license distribution, mailbox analytics, and risk signals before sending to AI. Uses system + user message prompting with temperature 0.4. Polished line-by-line markdown renderer with styled tables (auto-detected headers, color-coded deltas), blockquotes, HR rules, emoji support. Real-time word count + elapsed time during SSE streaming. Print-optimized CSS.
14. **Export to PDF/PNG** — Export executive briefing as a multi-page PDF (A4) or full-resolution PNG image. Uses html2canvas for rendering + jsPDF for PDF pagination. Lazy-loaded via dynamic imports.

## Navigation
- All pages share a consistent header with the Astra logo ("A" icon), app name, and nav links: Dashboard, License Guide.
- Executive Summary page adds contextual nav back to Dashboard.
- Footer on all pages: "© 2026 Cavaridge, LLC. All rights reserved."

## Data Model (shared/schema.ts)
- `users` — Auth table (placeholder)
- `reports` — Saved report snapshots (strategy, commitment, user data as JSONB)
- `executiveSummaries` — AI-generated summaries linked to reports
- `microsoftTokens` — Microsoft OAuth tokens (session-scoped)
- `user_sessions` — Session store (auto-created by connect-pg-simple)

## File Structure
```
client/src/
  App.tsx                       — Route registration (/, /licenses, /report/:id/summary)
  pages/dashboard.tsx           — Main dashboard with KPIs, strategy selector, subscriptions panel, data table, OAuth + file upload
  pages/executive-summary.tsx   — AI summary viewer with streaming, PDF/PNG export
  pages/license-comparison.tsx  — License comparison guide (up to 3 side-by-side, URL param pre-selection)
  pages/not-found.tsx           — 404 page
  lib/api.ts                    — API client functions (auth, upload, reports, sync, subscriptions)
  lib/license-data.ts           — Comprehensive M365 license feature dataset (19 licenses, 8 feature categories)
  lib/queryClient.ts            — React Query client
  hooks/use-toast.ts            — Toast notification hook
  components/ui/                — shadcn/ui components
server/
  db.ts                         — Database connection (Drizzle + pg)
  index.ts                      — Express app setup with session middleware
  microsoft-graph.ts            — Microsoft Graph API client (OAuth, user fetch, mailbox reports, active user report, subscribedSkus)
  routes.ts                     — API routes (auth, file upload, reports, summaries, subscriptions)
  storage.ts                    — Database storage interface (IStorage + DatabaseStorage)
shared/
  schema.ts                     — Drizzle schema definitions
```

## API Routes
- `GET /api/auth/microsoft/status` — Check OAuth connection status (includes tenantId)
- `GET /api/auth/microsoft/login` — Start OAuth flow → returns auth URL for consent screen
- `GET /api/auth/microsoft/callback` — OAuth redirect handler (extracts tid from JWT)
- `POST /api/auth/microsoft/disconnect` — Clear OAuth session
- `GET /api/microsoft/sync` — Fetch users + mailbox data via Graph API (includes tenantId)
- `GET /api/microsoft/report/active-users` — Office 365 Active User Detail report (30-day)
- `GET /api/microsoft/subscriptions` — Fetch tenant subscribed SKUs via Graph API
- `POST /api/upload/users` — Parse uploaded Active Users CSV/XLSX
- `POST /api/upload/mailbox` — Parse uploaded Mailbox Usage CSV/XLSX
- `GET /api/reports` — List saved reports
- `POST /api/reports` — Save a report snapshot
- `DELETE /api/reports/:id` — Delete a report
- `GET /api/reports/:id/summary` — Get saved executive summary
- `POST /api/reports/:id/summary` — Generate AI executive summary (SSE streaming)

## Environment Variables
- `DATABASE_URL` — PostgreSQL connection string
- `MICROSOFT_CLIENT_ID` — Azure AD Application (Client) ID
- `MICROSOFT_CLIENT_SECRET` — Azure AD Client Secret
- `MICROSOFT_TENANT_ID` — Azure AD Tenant ID (unused in code, hardcoded to "common")
- `OPENROUTER_API_KEY` — OpenRouter API key for AI summaries

## How to Use
### Option A: Microsoft OAuth (Recommended)
1. Click "Import Data" in the app header
2. Click "Sign in with Microsoft"
3. Go through the Microsoft consent screen (admin can consent for the whole org)
4. Click "Sync Data" to pull users, licenses, and mailbox usage
5. Tenant subscriptions auto-load in the collapsible panel

### Option B: File Upload
1. Go to M365 Admin Center → Reports → Usage → Active Users → Export CSV
2. Go to M365 Admin Center → Reports → Usage → Email activity → Export CSV
3. Click "Import Data" in the app header
4. Upload Active Users file first, then Mailbox Usage file
5. The app automatically parses, maps licenses to costs, and merges the data

### License Comparison
1. Click "License Guide" in the navigation bar, or click any license badge in the user directory
2. Select up to 3 licenses from the dropdowns (grouped by Suite vs Add-on/Standalone)
3. Compare features across 8 categories with included/excluded/partial indicators
