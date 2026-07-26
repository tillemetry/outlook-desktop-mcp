## Session 2026-07-26

### Maintain a dedicated macOS compatibility fork
**Context:** Outlook 16.111 changed its Accessibility hierarchy, breaking the inherited inbox reader; a newer third-party fork only contained Windows changes.
**Options considered:** Switch to the newer fork; retain a local-only patch; fork upstream and maintain macOS fixes.
**Chosen:** Fork upstream publicly and commit the macOS compatibility fallback with a regression test.
**Rationale:** Keeps the working fix available to other Mac users while preserving an upstream tracking path.
