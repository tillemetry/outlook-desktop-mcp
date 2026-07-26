# AGENTS.md

## Scope

Maintain the macOS Outlook Desktop MCP compatibility layer. Keep Windows behavior unchanged unless a task explicitly targets it.

## Guardrails

- Never commit Outlook data, credentials, logs, or local project tracking files.
- Keep macOS UI-scripting fallbacks narrow and preserve older working paths.
- Add one focused regression test for non-trivial behavior changes.

## Verification

```bash
.venv/bin/python -m unittest discover -s tests -p 'test_*.py' -v
```
