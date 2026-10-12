# UNPUBLISHED release notes template

This blank English template is an input, not a released version or a passed
validation record. Complete an external copy after the exact final commit is
selected under [the release procedure](../releasing.md). Replace every blank;
remove this template notice only in the completed publication notes. The GitHub
Release title must be the chosen tag alone.

- Maintainer-approved version: **[blank]** (root VERSION at the selected commit).
- Tag (`v` plus that version): **[blank]**.
- Exact reviewed and validated main commit (full 40-character SHA): **[blank]**.
- Version/commit decision and separate publication authorization in #11: **[blank]**.
- Reviewed PRs and actual merge/evidence links: **[blank]**.
- Confirmed non-draft Release URL: **[blank; record after creation/readback]**.
- Release channel / prerelease readback: **[blank]**.
- Annotated tag-object SHA and peeled commit readback: **[blank]**.

## Changes and distribution

**[blank: concrete resulting behavior and scope]**

Distribution: complete Git source snapshot and GitHub automatic source archives.
No standalone binary, extra release asset, package registry or marketplace.

## Compatibility and installation

- Tested Windows / Classic Outlook / PowerShell / OpenCode versions: **[blank]**.
- Caller/interface/receipt compatibility and migration (or justified None): **[blank]**.
- Fixed installation guide pinned to this commit: **[blank]**.
- Default user/global and explicit Project lifecycle guide pinned to this commit: **[blank]**.
- Status/update/deliberate older-source rollback/removal and recovery links: **[blank]**.
- Usage/privacy and read-only/untrusted-input limits: **[blank]**.

## Validation and remaining boundaries

| Required evidence | Actual exact-commit result and link |
| --- | --- |
| `windows-tests`, enforced on main | [blank: SHA, command/run, counts, failures/skips] |
| `opencode-discovery`, enforced on main | [blank: SHA, pinned native CLI version, scenarios, failures/skips] |
| Native Windows/Classic Outlook #9 signoff | [blank: exact candidate, sanitized before/after outcomes and signoff; private data stays private] |
| Installation/update/rollback/removal | [blank: environment, assertions, scope and evidence] |
| Model assessment | [blank: actual human assessment scope and limits; no offline/correctness guarantee] |
| Documentation/commands/links | [blank: actual rendered/semantic checks and untested commands] |

**[blank: known issues, coverage gaps, untested boundaries and their owners]**

Synthetic CI/discovery results do not certify live Outlook or model behavior.
Local COM retrieval uses existing Outlook authentication and OpenCode's chosen
model may receive selected data. No mail/attachment-content write path or
security-policy change is authorized. Publish only sanitized evidence; never
mail, addresses, mailbox/item identifiers, credentials or raw query results.
