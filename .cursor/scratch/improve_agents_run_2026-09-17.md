# improve-agents-md — 2026-09-17 (CLOUD) — polished

## Findings → mechanisms (not exhortations)

| Finding | Mechanism shipped |
|---------|-------------------|
| xlsx base64 Drive thrash (4 children, corrupt stubs) | `drive_create_verify.py` + alwaysApply `drive-uploads.mdc` + AGENTS.md §2 |
| Topics skills without Topics | CLOUD ADAPTER on all 4 skills + AGENTS.md §0/§5 |
| Monarch re-ask | alwaysApply `monarch-cloud.mdc` + AGENTS.md §4 |
| Premature packet build | alwaysApply `plan-lock.mdc` + AGENTS.md §3 |
| `search_files` free-text | AGENTS.md §2 query examples |

## Verify

```bash
python .cursor/tools/drive_create_verify.py --self-test
```

## Next-run backlog

- Wire Drive MCP wrapper that calls verify automatically (needs MCP-side change; not in this repo)
- Bench: blind session that tries base64 upload and must refuse
