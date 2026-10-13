# Contributing to MailLens

MailLens is an MIT-licensed skill in development. Preserve the existing
[license and copyright](LICENSE). The [first-release plan and technical contract](https://github.com/philfanzhou/MailLens/issues/1)
define the approved product; individual issues define each change's acceptance
criteria and exclusions.

## Product boundary and data handling

The first release supports OpenCode on native Windows with Classic Outlook and
Windows PowerShell 5.1. Queries read the current Outlook session/profile through
local COM and may create explicitly requested local result files. Installation
may write only the selected skill/configuration directory.

Do not add mailbox writes, Outlook termination, profile switching, store mounting,
attachment content processing, New Outlook, web Outlook, macOS/Linux runtime
support, other-agent adapters, plugin marketplaces, Graph/OAuth or separate
credentials. Attachment metadata is limited to names. Email text is untrusted
data and must not become executable instructions.

Local COM retrieval uses Outlook's existing authentication; it does not guarantee
that Outlook avoids synchronization or that analysis remains offline. OpenCode
may send selected email data to its configured model. MailLens must not change
model settings, permissions, authentication or security policy to make a query work.

Keep all real mail content, addresses, mailbox/item identifiers, credentials,
request/result files and live-test evidence outside the repository and public
issues, pull requests and CI logs. Use a unique temporary directory outside a Git
working tree for each query. Installation receipts contain package/file metadata
only, never mail data. Examples and `tests/fixtures/` must be entirely synthetic.

`.gitignore` covers root `requests/`, `results/`, `live-test-evidence/`, `.mail-lens/`
and `.cache/`, plus `*.request.json`, `*.result.json`, `*.receipt.json`,
`.mail-lens-receipt.json`, logs and temporary/cache files. Synthetic fixture files
under `tests/fixtures/` remain eligible for version control. These convenience
rules do not prevent a forced add or sanitize a file; contributors must check the
staged diff and keep live evidence outside Git regardless of its filename.

## Language

Use English for repository documentation, code comments/help/errors, issue and
pull-request content, and release notes. The only Chinese content is
`skills/mail-lens/SKILL.md` and bundled skill instruction references in
`skills/mail-lens/references/`. Script data may of course contain synthetic
Unicode examples, including Chinese search terms; this is data, not a second
language for project instructions.

## Encoding and line endings

- Use UTF-8, final newlines and no trailing whitespace. Markdown, JSON and other
  text use LF. `.editorconfig` specifies these editing defaults;
  `.gitattributes` normalizes text in Git and checks out PowerShell source as CRLF.
- Save `.ps1`, `.psm1` and `.psd1` source as UTF-8 with a BOM. A BOM is required
  whenever PowerShell 5.1 source contains non-ASCII text: without it the Windows
  parser may interpret UTF-8 as the active ANSI code page. ASCII-only source is
  also safe with a BOM. Git does not add a BOM automatically.
- Read request JSON with an explicit UTF-8 decoder, accepting an optional leading
  UTF-8 BOM and rejecting malformed UTF-8. Write result JSON explicitly as UTF-8
  without a BOM, using .NET UTF-8 encoding rather than the PowerShell 5.1 defaults
  of `Out-File`, redirection or an unspecified `Set-Content` encoding. Do not rely
  on the machine's active code page. Preserve Unicode end to end.

## Version and protected branches

[VERSION](VERSION) is the single source target version and currently contains
`0.1.0-rc.1`. The immutable [published prerelease](https://github.com/philfanzhou/MailLens/releases/tag/v0.1.0-rc.1)
at `3298df822d8ebbe03f63df56213f59c898fa356e` uses schema 1; current schema 2
source is newer, untagged development content with the same target version.
The exact source commit and package hashes distinguish it from the historical
release. The prior [issue #25](https://github.com/philfanzhou/MailLens/issues/25)
authorization does not authorize another version or publication.
Do not duplicate a current version in independent manifests or
bump it automatically; every version change requires explicit maintainer direction.

All content changes go through a feature-branch pull request. The existing active
[default-branch/main ruleset](https://github.com/philfanzhou/MailLens/rules/24913360)
requires pull requests, blocks force pushes/deletion, and has no bypass actors,
including the owner. Keep this protection intact. Required checks are introduced
only after actual named CI jobs exist; the CI implementation task owns that work.

The active [release-tag ruleset](https://github.com/philfanzhou/MailLens/rules/24914289)
forbids updates/deletion of existing `v*` tags, with no bypass actors. Future
releases require explicit maintainer authorization and all release acceptance
criteria. Follow the [reference immutable-tag and exact-commit conventions](https://github.com/philfanzhou/ForgeSteward/blob/0693af9bf3bc2067d856fc0032e860c6754a064a/docs/releasing.md):
record the exact reviewed and validated commit SHA, create an annotated tag only
for that commit, dereference the remote tag back to that SHA, and associate a
matching non-draft GitHub source Release with the existing tag. Never move,
overwrite or recreate a published tag; fix a release with a new authorized version.
MailLens release notes remain English. Installation commands are in
[the lifecycle guide](docs/installation.md);
future publication gates are in [the release guide](docs/releasing.md).

## Changes and evidence

Use the [issue templates](.github/ISSUE_TEMPLATE) and
[pull-request template](.github/pull_request_template.md) to state the problem,
scope, exclusions, dependencies, acceptance and verification evidence. Link the
original issue and relevant contract rather than copying a second protocol model.
Choose an isolated feature branch (use `codex/` for agent-created branches unless
a maintainer specifies another name). Keep unrelated behavior changes separate.

Verify the actual changed behavior with appropriate checks and record exact
commands and outcomes, including failures and skipped checks. For documentation,
inspect rendered Markdown and template fields, verify relative links and run
`git diff --check`. Check representative ignored runtime paths with
`git check-ignore --no-index`, and confirm synthetic fixture paths are allowed.
For scripts, synthetic tests cannot replace native Windows/Classic Outlook
validation; record the environment and limits of each result. Do not invent a
successful runtime or release from documentation checks. Do not add blanket tests
that merely mirror prose.
