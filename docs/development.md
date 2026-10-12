# Native Windows development checks

[Windows CI](../.github/workflows/ci.yml) runs on pull requests to `main` and
pushes to `main`. The stable checks are `windows-tests` and `opencode-discovery`.
Both use Windows Server 2022, Windows PowerShell 5.1 in STA, read-only repository
permissions, pinned checkout code and a 20-minute job timeout. Superseded runs
are cancelled; cancellation is not successful validation. CI does not access
Outlook, call a model, publish a release or change repository protection.

## Reproduce a job

From native Windows, use a fixed source checkout and a new absolute directory
outside Git. The helper refuses to reuse an existing run directory:

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
