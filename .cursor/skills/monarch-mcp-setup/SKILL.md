# Monarch MCP — Option 2 (official)

Official connector: `https://api.monarch.com/mcp`  
Help: https://help.monarch.com/hc/en-us/articles/50207234679956-Monarch-MCP-Connector

## What we verified from this Cloud Agent (2026-09-17)

| Check | Result |
|-------|--------|
| Host reachable | Yes — POST → `401 invalid_token` (OAuth required); GET → `405` |
| Monarch in cloud tool catalog | **No** — cannot `mcp_auth` Monarch from this session |
| Config in repo | `.cursor/mcp.json` points at the official URL (no secrets) |

**OAuth needs a browser.** Complete connection in **Desktop Cursor** (or another Cursor surface that can open Monarch’s authorize page). This headless cloud VM cannot finish the login redirect.

## Desktop setup (do this on your machine)

1. Open this repo in **Cursor Desktop** (so it picks up `.cursor/mcp.json`), **or** add the same block to `~/.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "monarch": {
      "url": "https://api.monarch.com/mcp"
    }
  }
}
```

2. Fully quit and reopen Cursor (MCP configs often load only on restart).
3. **Settings → Tools & MCP** (or Customize → MCPs) → find **monarch** → **Connect**.
4. Browser opens Monarch → sign in → **Authorize**. You never put your Monarch password in `mcp.json`.
5. Confirm Monarch tools appear (accounts, transactions, etc.).

If the browser doesn’t open: copy the authorize URL from MCP logs / status and open it manually. If authorize succeeds but Cursor stays disconnected: check firewall/loopback for `http://localhost:8787/callback`.

## After you’re connected

- Use a **Desktop** Agent chat to query Monarch live (tags, categories, “show STR payouts”).
- Cloud Agents may still lack Monarch until Cursor exposes that MCP to the cloud tool catalog — keep CSV drops as the cloud fallback (`Documents/Monarch Money`).

## Do not use

Unofficial servers that store `MONARCH_EMAIL` / `MONARCH_PASSWORD` / MFA secrets in `mcp.json`. Prefer official OAuth only.
