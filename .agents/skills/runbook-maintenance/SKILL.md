---
name: runbook-maintenance
description: Build and maintain a comprehensive operations runbook (RUNBOOK.md) for any project. Use when starting a new project, after adding/changing routes, files, schema, or features, or when the user asks to create or update a runbook. Triggers: runbook, operations doc, project documentation, RUNBOOK.md, document the project.
---

# Runbook Maintenance

Build and keep a living `RUNBOOK.md` at the project root that documents everything an operator or future AI session needs to understand and run the application.

## When to Create

- At project start, once the app has at least a backend and a database
- When the user asks for project documentation or a runbook
- When you notice `RUNBOOK.md` does not exist

## When to Update

Update the runbook **in every session** that changes any of the following:

- Files added, removed, or significantly changed (update File Map)
- API routes added, changed, or removed (update API Routes)
- Database tables or columns changed (update Schema)
- Environment variables or secrets added (update Env Vars)
- New bugs discovered and fixed (add Troubleshooting entry)
- Architecture or dependency changes (update System Overview)

## Required Sections

Every runbook must include these sections. See `template.md` in this skill folder for a ready-to-use starter.

### 1. System Overview
- One-paragraph description of what the app does
- Architecture table: layer → technology (Frontend, Backend, Database, etc.)

### 2. File Map
- One table per directory group (server, client, shared, config)
- Columns: **File** | **Lines** | **Purpose**
- Line counts must reflect reality — run `wc -l` to verify
- Update line counts whenever a file changes significantly (±50 lines)

### 3. Database Schema
- One table per DB table showing columns, types, and constraints
- Note foreign keys and indexes
- Include the table count in the section heading

### 4. API Routes
- Group by resource (e.g., Authentication, Deals, Documents)
- Columns: **Method** | **Path** | **Purpose** | **Auth**
- Every route in the codebase must appear here — grep for `app.get`, `app.post`, `router.get`, etc. to verify completeness

### 5. Authentication & Authorization
- How auth works (session, JWT, OAuth, etc.)
- Role/permission model
- Protected vs public routes

### 6. Environment Variables & Secrets
- Table with: **Variable** | **Required** | **Description**
- Never include actual secret values
- Note which are auto-provided by the platform vs user-supplied

### 7. Frontend Navigation
- List of pages/routes with their paths and purposes
- Note which require authentication

### 8. Operational Troubleshooting
- Table with: **Symptom** | **Cause** | **Fix**
- Add an entry every time you fix a non-trivial bug
- Group by area (Database, API, Frontend, PDF, Auth, etc.)

### 9. Version Footer
- Both the title heading and the last line must show the version
- Format: `v{major}.{minor}` (e.g., v1.0, v1.1, v2.0)

## Version Bumping Rules

| Change type | Version bump |
|------------|-------------|
| New section added, major feature documented | Major (v1.0 → v2.0) |
| Entries added/updated within existing sections | Minor (v1.0 → v1.1) |
| Typo or formatting fix only | No bump needed |

Always include a change summary in the footer:

```
*ProjectName — Operations Runbook v1.2*
*Updated Month Year — Brief description of what changed*
```

## Quality Checklist

Before finishing any runbook update, verify:

- [ ] Line counts in File Map match reality (`wc -l filename`)
- [ ] All API routes are listed (grep the codebase for route definitions)
- [ ] No stale entries for deleted files or removed routes
- [ ] Version number matches in both the title and the footer
- [ ] Every recent bug fix has a Troubleshooting entry
- [ ] No secret values appear anywhere in the document

## Tips

- **Keep it scannable**: Use tables over prose wherever possible
- **Be precise**: "~1722 lines" is better than "large file"
- **Group logically**: API routes by resource, files by directory, env vars by source
- **Date your updates**: The footer should always say when it was last changed
- **Use the template**: Read `template.md` in this skill folder for a starting structure you can copy and adapt

## Optional Sections

Add these when the project warrants them:

- **Usage Limits / Rate Limiting** — if the app has plan-based tiers or rate limits
- **AI Pipelines** — if the app uses LLMs, embeddings, or vector search
- **Design System** — if there's a defined color palette, component library, or theme
- **Deployment & Workflows** — if there are multiple workflows or deployment steps
- **Scoring / Business Logic** — if there's complex domain logic worth documenting
