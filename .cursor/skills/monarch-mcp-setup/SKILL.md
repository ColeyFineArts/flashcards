# Monarch MCP — Option 2 (official)

Official connector: `https://api.monarch.com/mcp`  
Help: https://help.monarch.com/hc/en-us/articles/50207234679956-Monarch-MCP-Connector

## Verified from this Cloud Agent (2026-09-17)

| Check | Result |
|-------|--------|
| Host reachable | Yes — POST → `401 invalid_token` (needs OAuth); GET → `405` |
| Monarch already in this run’s tool catalog | **No** |
| Repo config | `.cursor/mcp.json` has the official URL (no secrets) |

## Path A — Desktop Cursor (uses repo `mcp.json`)

1. Open this repo in **Cursor Desktop** (loads `.cursor/mcp.json`), or paste the same block into `~/.cursor/mcp.json`.
2. Fully quit + reopen Cursor.
3. **Settings → Tools & MCP** → **monarch** → **Connect**.
4. Browser → Monarch sign-in → **Authorize** (password stays with Monarch, not in config).
5. Confirm Monarch tools appear in the chat tool list.

If the browser doesn’t open, copy the authorize URL from MCP status/logs and open it manually. Desktop OAuth redirect is typically `http://localhost:8787/callback`.

## Path B — Cloud Agents (dashboard — not the repo file)

Cloud Agents **do not** load project `.cursor/mcp.json`. Add Monarch under the Cloud Agents MCP UI:

1. Open [cursor.com/agents](https://cursor.com/agents) (or Dashboard → Cloud Agents → Plugins & MCPs).
2. Add a custom **HTTP** MCP server with URL exactly: `https://api.monarch.com/mcp`
3. Complete **OAuth** when prompted (Cloud redirect: `https://www.cursor.com/agents/mcp/oauth/callback`).
4. Start a **new** Cloud Agent run (or reconnect MCP) and confirm Monarch tools show in the catalog.

Until Path B is done, this cloud session keeps using **CSV drops** in Drive `Monarch Money`.

## After connect — useful asks

- “List accounts synced in Monarch”
- “TY2025 transactions tagged Business / missing TAX:STR”
- “Airbnb/VRBO payouts this year vs category Travel & Vacation”

Hybrid rule unchanged: Monarch = operating P&L; CapEx + sale CDs stay on registers.

## Do not use

Unofficial MCP servers that store `MONARCH_EMAIL` / `MONARCH_PASSWORD` / MFA secrets in config. Official OAuth only.
