---
name: pakplan-generator
description: >-
  Generates and emails FFE weekly pack plans. Downloads Pakvolumes and
  Markpryse Excel files from SharePoint, builds per-producer Excel files,
  and sends results via Microsoft Graph. Use when the user asks to run
  pakplan, generate pack plans, or send the weekly FFE pack plans.
allowed-tools: Bash(node:*)
---

# FFE Pakplan Generator

Runs `scripts/pakplan/run.mjs` — zero npm dependencies, uses Node.js built-ins only.

## Prerequisites

The following env vars must be set (add to `.env.local`):

```bash
AZURE_TENANT_ID=<your-tenant-id>
AZURE_CLIENT_ID=<app-client-id>
AZURE_CLIENT_SECRET=<client-secret>
PAKPLAN_SEND_AS=tjaart@ffesa.co.za   # mailbox to send from
```

The Azure AD app needs **Microsoft Graph Application** permissions (admin consent required):
- `Sites.Read.All` (or `Sites.ReadWrite.All`) — downloads files from SharePoint via Graph
- `Mail.Send` — sends email

## Steps

1. **Check env vars** — confirm `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, and `AZURE_CLIENT_SECRET` are set. If not, show the `.env.local` snippet above and stop.

2. **Run the generator**:

   With `.env.local`:
   ```bash
   pnpm exec dotenvx run -f .env.local -- node scripts/pakplan/run.mjs
   ```

   If vars are already exported in the shell:
   ```bash
   node scripts/pakplan/run.mjs
   ```

3. **Report results** — from the script output, summarise: export week, number of pack plans, producer list, and whether the email was sent.

## Dry run (no email)

```bash
pnpm exec dotenvx run -f .env.local -- \
  PAKPLAN_DRY_RUN=1 node scripts/pakplan/run.mjs
```

## Troubleshooting

| Error | Fix |
|-------|-----|
| `AADSTS700016` | App not found — verify `AZURE_CLIENT_ID` and `AZURE_TENANT_ID` |
| `AADSTS7000215` | Invalid client secret — regenerate in Azure portal |
| `HTTP 401` on SharePoint | Token scope wrong or app not consented for `Sites.Read.All` |
| `HTTP 403` on SharePoint | Grant `Sites.Read.All` Application permission + admin consent |
| `HTTP 403` on sendMail | Grant `Mail.Send` Application permission + admin consent |
| `HTTP 400: Cannot access mailbox` | `PAKPLAN_SEND_AS` must be a licensed Exchange user or shared mailbox |
| `Missing required environment variable` | Set the three `AZURE_*` vars in `.env.local` |
