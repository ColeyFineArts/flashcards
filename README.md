# flashcards

## Agent operating card

Cloud agents: read [`AGENTS.md`](AGENTS.md) first (Drive upload policy, Monarch CSV path, Topics CLOUD ADAPTER).
Skills live under [`.cursor/skills/`](.cursor/skills/). After Drive `create_file`, run:

```bash
python .cursor/tools/drive_create_verify.py --self-test
python .cursor/tools/drive_create_verify.py --got-filesize F --expected-min-bytes N --title "…"
```

