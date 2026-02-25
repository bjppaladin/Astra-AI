# PROJECT_NAME — Operations Runbook v1.0

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

---

## 1. System Overview

<!-- One paragraph: what does this app do? -->

PROJECT_NAME is ...

### Architecture

| Layer | Technology |
|-------|-----------|
| Frontend | |
| Backend | |
| Database | |
| Auth | |

---

## 2. File Map

### Server Files (`server/`)

| File | Lines | Purpose |
|------|-------|---------|
| `index.ts` | ~XX | App bootstrap, middleware, HTTP server |

### Client Files (`client/src/`)

| File | Lines | Purpose |
|------|-------|---------|
| `App.tsx` | ~XX | Root component, routing |

### Shared Files (`shared/`)

| File | Lines | Purpose |
|------|-------|---------|
| `schema.ts` | ~XX | Drizzle schema, Zod validators, TypeScript types |

### Config Files (root)

| File | Lines | Purpose |
|------|-------|---------|
| `package.json` | ~XX | Dependencies, scripts |
| `tsconfig.json` | ~XX | TypeScript configuration |
| `tailwind.config.ts` | ~XX | Tailwind theme |

---

## 3. Database Schema

<!-- One subsection per table. Include column name, type, and constraints. -->

### `users`

| Column | Type | Constraints |
|--------|------|------------|
| `id` | serial | PK |
| `email` | varchar(255) | NOT NULL, UNIQUE |
| `created_at` | timestamp | DEFAULT now() |

---

## 4. API Routes

### 4.1 Authentication

| Method | Path | Purpose | Auth |
|--------|------|---------|------|
| POST | `/api/auth/login` | Log in | No |
| POST | `/api/auth/logout` | Log out | Yes |
| GET | `/api/auth/me` | Current user | Yes |

### 4.2 Resource Name

| Method | Path | Purpose | Auth |
|--------|------|---------|------|
| GET | `/api/resource` | List all | Yes |
| GET | `/api/resource/:id` | Get one | Yes |
| POST | `/api/resource` | Create | Yes |
| PATCH | `/api/resource/:id` | Update | Yes |
| DELETE | `/api/resource/:id` | Delete | Yes |

---

## 5. Authentication & Authorization

<!-- How does auth work? Session, JWT, OAuth? What roles exist? -->

- **Method**: express-session with PostgreSQL store
- **Roles**: user, admin
- **Protected routes**: All `/api/*` except auth endpoints

---

## 6. Environment Variables & Secrets

| Variable | Required | Description |
|----------|----------|-------------|
| `DATABASE_URL` | Yes | PostgreSQL connection string (auto-provided) |
| `SESSION_SECRET` | Yes | Session encryption key |

---

## 7. Frontend Navigation

| Path | Page | Auth Required |
|------|------|--------------|
| `/` | Home / Dashboard | No |
| `/login` | Login | No |
| `/dashboard` | Main dashboard | Yes |

---

## 8. Operational Troubleshooting

### 8.1 Common Issues

| Symptom | Cause | Fix |
|---------|-------|-----|
| 500 on startup | Missing DATABASE_URL | Check env vars |
| Login fails silently | Session secret not set | Add SESSION_SECRET to secrets |

---

*PROJECT_NAME — Operations Runbook v1.0*
*Updated MONTH YEAR — Initial runbook*
