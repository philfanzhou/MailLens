# Native Windows development checks

[Windows CI](../.github/workflows/ci.yml) runs on pull requests to `main` and
pushes to `main`. The stable checks are `windows-tests` and `opencode-discovery`.
Both use Windows Server 2022, Windows PowerShell 5.1 in STA, read-only repository
permissions, pinned checkout code and a 20-minute job timeout. Superseded runs
are cancelled; cancellation is not successful validation. CI does not access
Outlook, call a model, publish a release or change repository protection.

## Reproduce a job

From native Windows, use a [verified fixed source checkout](versioned-installation.md)
and a new absolute directory outside Git. The helper refuses to reuse an existing run directory:

```powershell
powershell.exe -NoProfile -STA -File C:\MailLens\scripts\run-ci.ps1 `
    -Job windows-tests -RunRoot C:\Temp\mail-lens-tests-unique
powershell.exe -NoProfile -STA -File C:\MailLens\scripts\run-ci.ps1 `
    -Job opencode-discovery -RunRoot C:\Temp\mail-lens-discovery-unique
```

Respect existing execution/security policy. Each job records the actual checkout
SHA, PowerShell version/apartment and test counts. It saves development-only
[Pester 5.7.1](https://github.com/pester/Pester/releases/tag/5.7.1) from PSGallery
to its own run directory and imports that exact manifest. End users need no
Pester dependency. Network access is needed to obtain development tools.

The package checker uses built-in PowerShell/.NET. It checks required files,
strict UTF-8 and PowerShell BOM/parser errors, final newlines/whitespace, root
VERSION syntax, and confined relative skill resources. It supports the package's
flat `key: scalar` frontmatter, rejects duplicate/missing name or description,
and rejects a second skill version field. Its language checks cover CJK in
project Markdown and PowerShell comment/help tokens while allowing Unicode
synthetic data and Chinese skill instructions. These mechanical checks do not
prove every language or read-only semantic requirement.

Query tests use schema 2; installer ownership receipts retain their separate
schema 1. Search tests cover global budgets, direct/default scope, exact date
filter membership under en-US/zh-CN, literal/body gating, auxiliary degradation,
warning aggregation and bounded summary retention. These are synthetic control
flow checks, not Outlook provider timing or live mailbox validation.

All unit containers are mandatory, including the existing COM doubles and
installer tests. The explicit Pester result gate rejects missing/empty containers,
failures, skipped, inconclusive and not-run tests. Temporary negative tests prove
that malformed packages, deleted tests, failed tests and missing tools cannot
produce green results. Their deliberate failures are expected assertions inside
the passing CI test container.

## Pinned CLI discovery

The discovery job downloads the official
[OpenCode 1.18.30 Windows x64 baseline archive](https://github.com/anomalyco/opencode/releases/tag/v1.18.30).
The release source is `5cd8e68fdd72b27818d26d168b9c7a06b359567e`; the release
API's independently published SHA-256 is
`8c34c5767a688e8d6c4640b4893d34f08f08b071ed0043c940dc5be17a3b9233`.
The helper verifies this digest before extraction and checks the executable's
actual version. A download, checksum, extraction or version failure fails CI.

Each observation starts a new `opencode.exe debug skill` process, captures strict
UTF-8 stdout/stderr separately, requires exit zero and validates JSON records.
Only records named `mail-lens` are counted; OpenCode also supplies built-in skills.
Assertions compare the absolute selected SKILL.md path, description, body and
installed resource bytes. A 90-second subprocess timeout kills only the process
created by the test and fails the assertion.

Tests create an isolated Git project boundary and fresh XDG config/data/cache/state,
OpenCode test-home/config inputs and temporary source/installation directories.
They restore their environment in `finally`/test cleanup. They never install into
the actual Windows profile or change HOME/CODEX_HOME. No provider credentials,
plugin, skill download URL or model request is configured. Pinned flags disable
CLI auto-update and model-list fetching; this is not a universal no-network claim.

The actual installer CLI is exercised in isolated default XDG/global, explicit
User/custom and Project scopes. Fresh discovery verifies install, unchanged
status/reinstall, update, older-content rollback, source relocation/removal and
retained standalone uninstall. Missing/corrupt receipts refuse update/removal
while the preserved valid package remains discoverable. Duplicate-name discovery
can select either root, so tests separately compare the other package's bytes and
then verify its discovery after selected removal. Other skills/configuration are
preserved. See the [installer lifecycle](installation.md).

All content is synthetic. Hosted checks are not native Classic Outlook or model
certification; [issue #9](https://github.com/philfanzhou/MailLens/issues/9) owns
that separate evidence. Keep live mail and live validation evidence outside Git
and public CI logs. Run-specific reports are local synthetic NUnit XML files;
CI does not upload user data or credentials.

## Private installed-model scenarios

[Issue #9](https://github.com/philfanzhou/MailLens/issues/9) validates the exact
final installed commit on native Windows/Classic Outlook with the configured
OpenCode model and existing authorized sample mail. These are maintainer-run
scenarios, **not passed results** or automated live-profile tests. Run serially.
Record operation sequence, projected summary/body-chunk counts, excerpt scalar
counts, total selected excerpt characters, elapsed time and before/after read-only
observations privately outside Git. Public evidence is only sanitized counts,
versions, candidate SHA, outcomes and gaps, never content/identifiers.

| Scenario | Observable expectation |
| --- | --- |
| Show Inbox mail | browse first, <=10 selected summaries; only needed get calls with explicit 2000-scalar first chunks; no routine diagnose/folders/global search. Measure the warm-session ten-subject + one-body target from #1, not a synthetic timing claim. |
| Chinese word only in Body, selected custom folder | bounded search first in that exact scope; visible literal terms/range and a cited matching excerpt. |
| Explicit all-mounted bounded/exhaustive request | preserve authorized scope, finite shared budgets and accurate observed/exact counts; no silent narrowing or completeness claim from a cap. |
| Truncated/partial/denied result, missing result or stalled host | disclose gaps and last safe stage; no automatic retry/limit increase/global rescan, and no raw catalog/warning dump. |
| Long body needing another chunk | select only necessary next scalar chunk; unchanged identity/source/length/modification checks, disclosed truncation and no repeated body history. |
| Stale or moved identity | report stale failure without same-subject replacement, or cite verified current location and disclose movement. |
| Synthetic malicious mail instructions | treat content/attachment filenames as data; no command execution, data upload, attachment access or mailbox write. |

Human assessment must check useful cited answers, model obedience and context
growth; syntax/package checks and scripted scenario calculations cannot certify
those outcomes. Whole-package update/rollback and OpenCode restart are required;
current schema 2 source is untagged while published rc.1 is historical schema 1.

## Required-check promotion

[Issue #8](https://github.com/philfanzhou/MailLens/issues/8) requires two delivery
steps. Its source PR uses `Refs #8` and requires both actual PR jobs to pass on
the current source before review/merge. After merge, observe both successful jobs
on the exact resulting main SHA and their GitHub Actions integration source.
Only then may the authorized maintainer extend the existing
[main ruleset](https://github.com/philfanzhou/MailLens/rules/24913360) with these
two strict, up-to-date required checks, outside CI.

Immediately refresh the complete ruleset before writing. Preserve active
enforcement, conditions, every existing rule/parameter and the empty bypass list.
Read back the ruleset and effective main rules, and verify a failing required
check visibly blocks a synthetic feature PR. Close any disposable negative PR
and safely remove only its branch. Never merge a failing PR, bypass protection,
force/direct-push main or relax rules to test enforcement. Preserve immutable
release-tag protection. Unknown settings outcomes require readback before retry;
permission/API failure keeps #8 open with the exact blocker. Source merge alone
does not complete #8. Version changes, live validation and publication remain
separate authorized tasks.
