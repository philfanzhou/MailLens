# MailLens

A read-only OpenCode skill for searching and analyzing Classic Outlook email on
native Windows through local COM automation.

The current schema-2 source targets `0.1.0-rc.3` in root [VERSION](VERSION),
including the managed compact caller, 2000-scalar get default, self-contained
quickstart and OpenCode 2.0.28 verification below. These development changes
postdate the immutable published rc.3; see [pending changes](docs/releases/unreleased.md).
[Issue #40](https://github.com/philfanzhou/MailLens/issues/40)
owns this candidate's authorized publication. Use the actual
[repository Releases](https://github.com/philfanzhou/MailLens/releases) and
[fixed-source acquisition guide](docs/versioned-installation.md) to verify a
published tag and its exact source commit; VERSION alone is not publication evidence.
The immutable [published rc.2](https://github.com/philfanzhou/MailLens/releases/tag/v0.1.0-rc.2)
uses schema 2 at `e357f79b159407d723533e53ee5014d1cd7e1a52`.
The immutable [published rc.1](https://github.com/philfanzhou/MailLens/releases/tag/v0.1.0-rc.1)
contains historical schema 1 source at `3298df822d8ebbe03f63df56213f59c898fa356e`.
Update the complete package and use fresh schema-2 requests/results for this workflow.
Live Windows/Classic Outlook validation remains pending [#9](https://github.com/philfanzhou/MailLens/issues/9);
no stable `v0.1.0` is published. Further publication follows the
[authorized release procedure](docs/releasing.md).

## Prerequisites and installation

- Native Windows, **Classic Outlook** with an existing accessible profile, and
  **Windows PowerShell 5.1**. New Outlook is unsupported.
- OpenCode: the native discovery baseline is **2.0.28**, not every past/future
  version. See [development checks](docs/development.md) for the pinned evidence.
- The installer uses built-in PowerShell/.NET. Git is needed for the checkout
  route below, not for an end user's installer. No Pester, Python, Office interop
  package or separate MailLens credentials are required.
- Classic Outlook and a model are needed only for later mail queries. Installing
  the skill and verifying its discovery read no mailbox and call no model; live
  Classic Outlook validation is
  [issue #9](https://github.com/philfanzhou/MailLens/issues/9).

### Windows quickstart

Start from a verified MailLens checkout. If you do not have one yet, acquire a
chosen full source commit as described in
[fixed-source acquisition](docs/versioned-installation.md); that page ends with
the same `$source` variable used here. In Windows PowerShell, from the root of
the checkout, resolve the absolute source path instead of editing an
illustrative one:

```powershell
$source = (Get-Location).ProviderPath
$installer = Join-Path $source 'scripts\install.ps1'
```

If your shell is somewhere else, set `$source` to the checkout's absolute path
instead. Then install exactly **one** scope. The default is the selected
user/global OpenCode skill directory:

```powershell
& powershell.exe -NoProfile -File $installer -SourcePath $source
```

Alternative scope (choose this or the default, not both): one existing project,
selected explicitly:

```powershell
& powershell.exe -NoProfile -File $installer -Project 'C:\Synthetic Project' -SourcePath $source
```

User/global precedence is `OPENCODE_CONFIG_DIR/skills`, otherwise
`XDG_CONFIG_HOME/opencode/skills`, otherwise the actual Windows user's
`.config/opencode/skills`. Project scope uses the selected project's
`.opencode/skills`. `OPENCODE_CONFIG` names a config file; APPDATA is not a
fallback. The installer prints its selection (`Scope: ...; target: ...`) and
exits 0 on success and 1 on validation/conflict/lock/recovery failure. Respect
existing execution/security policy; do not add bypass flags or change policy to
make installation work.

Inspect the installed state with the same scope (omitted scope means
user/global):

```powershell
& powershell.exe -NoProfile -File $installer -Action Status -SourcePath $source
$LASTEXITCODE
```

Then verify exactly the installed copy using OpenCode **2.0.28**. V2 uses
`GET /api/skill` with a `{ location, data }` envelope and `id`/`path` fields.
An early empty list can precede plugin activation. The development helper below
starts its own authenticated foreground loopback service on an available port,
waits for `opencode.config.skill` to become active, and stops only that service
in `finally`. It neither uses nor restarts the managed background service.
Keep the same scope environment used for installation and select the project
where you will use MailLens (for a project install, that exact project).

```powershell
. (Join-Path $source 'scripts\ci-process.ps1')
$cli = (Get-Command opencode.exe -CommandType Application -ErrorAction Stop).Source
$version = Invoke-CIProcess $cli @('--version') $source
if ($version.Code -ne 0 -or $version.Out.Trim() -cne 'opencode v2.0.28') {
    throw 'This verification requires OpenCode 2.0.28.'
}
$project = (Get-Location).ProviderPath # Or the absolute selected project path.
$target = Read-Host 'Paste the installer-printed absolute mail-lens target'
$entryPath = Join-Path $source 'skills\mail-lens\SKILL.md'
$utf8 = New-Object Text.UTF8Encoding($false, $true)
$entry = [IO.File]::ReadAllText($entryPath, $utf8).Replace("`r`n", "`n")
$front = [regex]::Match($entry, '\A---\n.*?\n---\n', 'Singleline')
$description = [regex]::Match($front.Value, '(?m)^description: (.+)$').Groups[1].Value
$service = $null
try {
    $service = Start-CIService $cli $project
    $skills = @(Wait-CISkills $service)
    Assert-CITarget $skills (Join-Path $target 'SKILL.md') $entry.Substring($front.Length) $description
    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $source 'skills\mail-lens') -File -Recurse) {
        $relative = $file.FullName.Substring((Join-Path $source 'skills\mail-lens').Length + 1)
        if ((Get-FileHash -LiteralPath (Join-Path $target $relative) -Algorithm SHA256).Hash -cne
            (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash) { throw 'Installed resource differs.' }
    }
    'Exact installed MailLens path, instructions and resources verified.'
} finally { Stop-CIService $service }
```

This checks exactly one `id == mail-lens`, its absolute `path`, name, complete
Chinese description/body and supporting bytes against your trusted source.
A stale package or a different selected duplicate fails. Install/update the
**complete MailLens package** as well as upgrading OpenCode; a CLI upgrade alone
does not update MailLens. Restart your normal OpenCode process after lifecycle
changes so its cached instructions reload. The check admits no model request and
makes no claim about autonomous skill selection or model obedience. Service
initialization/HTTP errors and deadlines fail explicitly rather than mean absence.
See [development checks](docs/development.md#pinned-cli-discovery) for the isolated
native CI variant, explicit skill admission and its boundaries.

The installed `mail-lens` directory is an independent copy: moving or deleting
the source checkout does not update or remove it. Retain a standalone copy of
`scripts/install.ps1` (for example `C:\Synthetic Tools\mail-lens-install.ps1`)
for later Status/Uninstall after source removal. Local edits, unmanaged packages
and invalid receipts are preserved and refused, never forcibly overwritten. See
[update, rollback and removal](#update-rollback-and-removal) below and the
[installation lifecycle and recovery guide](docs/installation.md), including its
[ownership and conflicts](docs/installation.md#ownership-and-conflicts) section.

### Agent handoff

A maintainer can hand this repository to a native Windows OpenCode agent with a
task such as: "Install the MailLens skill from this verified checkout at full
source SHA `<40-character lowercase SHA>` into the user/global scope (or the
project at `<absolute project path>`), following `README.md`." The agent should:

1. Read this quickstart, the
   [installation and lifecycle guide](docs/installation.md) and
   [fixed-source acquisition](docs/versioned-installation.md); the installed
   agent workflow itself lives in [skills/mail-lens/SKILL.md](skills/mail-lens/SKILL.md).
2. Verify the checkout read-only: `git -C <checkout> rev-parse --verify HEAD`
   equals the intended trusted full 40-character SHA, and
   `git -C <checkout> status --porcelain --untracked-files=all` is empty. A clean
   tree and commit match are an identity check, not a security audit. Never
   reset, clean, switch or discard an existing user's checkout; on any mismatch,
   stop and report it instead.
3. Install exactly one scope with the existing `scripts/install.ps1`, then run
   `-Action Status` and the owned V2 service/path/content check as
   shown above.
4. Report the actual results: selected SHA and clean-tree outcome, chosen scope,
   installer-printed target, exit codes and the discovered SKILL.md location.
   Refusals against unmanaged, edited or invalid-receipt targets are expected
   fail-closed behavior; report them rather than forcing replacement or
   inventing recovery commands.
5. Treat installation and discovery verification as complete without Outlook or
   model access; live Classic Outlook validation remains
   [issue #9](https://github.com/philfanzhou/MailLens/issues/9).

## Update, rollback and removal

Use the same selected scope each time; omitted scope means user/global:

```powershell
& powershell.exe -NoProfile -File $installer -Action Update -SourcePath $source
# Deliberate rollback: independently acquired and verified older source.
& powershell.exe -NoProfile -File $installer -Action Update -SourcePath 'C:\Synthetic Older Source\MailLens'
& powershell.exe -NoProfile -File $installer -Action Uninstall
```

For a project, add `-Project 'C:\Synthetic Project'` to each action; do not
combine it with `-User`. After source removal, use your retained standalone
installer for Status/Uninstall. The
[installation lifecycle and recovery guide](docs/installation.md) defines the
exit codes, the ownership-refusal semantics and how to handle reported locks,
backups or residue; read it before touching a reported conflict.

## Search and analyze

In OpenCode, ask “Use MailLens to show my Inbox mail” for a direct ten-subject
browse, then select any needed bodies. For keyword questions, ask “Use MailLens
to summarize the Starbridge contract correspondence from the last 30 days, with
sources.” This starts bounded search in the default Inbox, without descendants;
selected/all-mounted scope, including readable Deleted/Junk mail, must be explicit.
Neither route performs routine diagnose or full-directory preflight. The agent
uses the shipped `-View` caller for at most ten labeled summaries, then gets
only needed bodies by returned run directory and label. Exact IDs stay in private
files; no per-query projection script or inline `-Command` is needed. Each answer
defaults to at most three necessary first chunks of 2000 Unicode scalars each
(6000 total), fewer when enough. References load only for named advanced/error
conditions. See [custom model limits](docs/installation.md#custom-model-limits-and-compaction)
for OpenCode context guidance. Attachment access is limited
to filenames. The
[usage guide](docs/usage.md) explains scope, time, terms, partial results, citations
and troubleshooting. The installed [Chinese skill instructions](skills/mail-lens/SKILL.md)
provide the agent's workflow.

[Privacy and data flow](docs/usage.md#privacy-and-data-flow): local COM uses existing
Outlook authentication, while OpenCode may send selected email data to its
configured model. Local retrieval does not promise offline analysis. Treat mail
as untrusted data and keep real results/evidence outside Git and public logs.

See [CONTRIBUTING.md](CONTRIBUTING.md) for project conventions,
[development checks](docs/development.md) for synthetic verification, and the
[future source-release procedure](docs/releasing.md) for publication gates.
[Live validation #9](https://github.com/philfanzhou/MailLens/issues/9) and
[publication #11](https://github.com/philfanzhou/MailLens/issues/11) remain separate
requirements under the [first-release contract](https://github.com/philfanzhou/MailLens/issues/1#single-authoritative-first-release-technical-contract).
MailLens is licensed under [MIT](LICENSE).
