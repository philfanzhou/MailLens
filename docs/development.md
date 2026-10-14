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

The discovery job obtains the pinned official npm package
[`@opencode/cli-windows-x64-baseline@2.0.28`](https://registry.npmjs.org/@opencode%2fcli-windows-x64-baseline/2.0.28).
Its exact tarball URL and published SHA-512 integrity are fixed in
[`run-ci.ps1`](../scripts/run-ci.ps1); the immutable upstream source is
`026b106bc5d41b679ffc413264229c340871e4d3`. Before executing any bytes, the
helper verifies the digest, validates tar checksums/types/names/sizes and requires
one exact `package/bin/opencode.exe` alongside package metadata. It extracts only
that executable into its new owned tool directory, runs no package scripts or
global install, and requires the actual `opencode v2.0.28` version. Download,
checksum, unsafe/ambiguous extraction and version errors fail CI.

Each observation starts a foreground owned service with explicit `127.0.0.1`
and an available port, never the managed background service or its default port.
A generated in-memory password authenticates HTTP; it is supplied only to the
child environment and never printed. Both service streams are drained and
discarded. HTTP errors expose only a safe category, not credentials or response
bodies. Requests reject redirects, bound response bytes, explicitly decode UTF-8
with optional BOM, and have request/read deadlines. Process commands have a
90-second deadline; readiness has a 30-second deadline and checks child liveness.
A port collision fails closed instead of adopting another listener.

V2 [`GET /api/skill`](https://github.com/anomalyco/opencode/blob/026b106bc5d41b679ffc413264229c340871e4d3/packages/protocol/src/groups/skill.ts)
returns `{ location, data }`; records use `id` and absolute `path`. Before any
positive **or negative** assertion, the helper polls `GET /api/plugin` at the
explicit project location for `opencode.config.skill` with `state.status == active`.
An early empty list is inconclusive. A ready missing target fails a positive
check, and proves absence after uninstall. Malformed responses, plugin failure,
authentication/HTTP failure, service exit and timeout fail explicitly. Each
installer observation uses a fresh service, avoiding stale cached instructions.
Unit scenarios cover initialization, true absence, wrong/duplicate IDs and paths,
content, malformed records, occupied ports, HTTP failures and owned cleanup.

Tests isolate XDG config/data/cache/state, supported test-home/custom config,
project and temporary package paths. They restore environment variables and
remove only their owned files/processes in cleanup, including assertion failure
and cooperative cancellation. Force-kill/power-loss cleanup is not guaranteed.
They never modify the actual Windows profile or HOME/CODEX_HOME. No provider
credentials, plugin/skill download configuration or model request is introduced.
Pinned flags disable auto-update and model-list fetching; this is not a universal
no-network claim. See the [V2 migration guide](https://opencode.ai/v2/docs/migrate-v1/)
and [V2 skill guide](https://opencode.ai/v2/docs/skills/).

The actual installer exercises all six existing cases in default XDG/global,
User/custom and Project scopes: install/status/unchanged reinstall, update,
older-content rollback, source relocation/deletion and retained uninstall;
missing/corrupt receipt refusals; selected duplicate removal. V2 project discovery
wins the duplicate id; removing it exposes the exact preserved global package.
Assertions compare exactly one `id == mail-lens`, selected absolute path, name,
complete Chinese description/body and every installed supporting resource's bytes.
Other skills and configuration are preserved. See [installer lifecycle](installation.md)
and the runnable [installed-copy verification](../README.md#windows-quickstart).

For every package check, an owned session receives
`POST /api/experimental/session/:sessionID/skill` with
`{"id":"mail-lens","resume":false}`. Stored skill-message `text` must equal the
complete installed body; the owned session is then deleted. This admits the skill
without a prompt/model loop; it does not certify autonomous selection/obedience.
Stored messages are body-only. The separate model tool wrapper is established by
pinned [`Skill.prepare` / `toModelOutput`](https://github.com/anomalyco/opencode/blob/026b106bc5d41b679ffc413264229c340871e4d3/packages/core/src/skill.ts):
it derives the base directory from `dirname(skill.path)`, resolves supporting
paths there, and samples at most ten files in the rendered `skill` output.
The native test checks all supporting bytes separately; it does not pretend that
wrapper appears in the stored-message API.

Historical exact-main run [38412384951](https://github.com/philfanzhou/MailLens/actions/runs/38412384951)
validated 432 unit cases and six OpenCode **1.18.30** discovery cases at
`e357f79b159407d723533e53ee5014d1cd7e1a52`. Those results remain historical;
[#37](https://github.com/philfanzhou/MailLens/issues/37) records exact-source V2
Windows evidence. V1 `debug skill` and its top-level array are obsolete for V2.

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
