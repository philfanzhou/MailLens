# MailLens

A read-only OpenCode skill for searching and analyzing Classic Outlook email on
native Windows through local COM automation.

MailLens has **no published release**. Root [VERSION](VERSION) is `0.1.0`, an
unpublished development baseline. A source commit can be selected for development;
there is no available `v0.1.0` release or release download to install.

## Prerequisites and installation

- Native Windows, **Classic Outlook** with an existing accessible profile, and
  **Windows PowerShell 5.1**. New Outlook is unsupported.
- OpenCode: native discovery is tested with **1.18.30**, not every past/future
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

Then verify that OpenCode discovers exactly the copy you installed, using a
fresh OpenCode process (`opencode.exe debug skill` is the documented tested
baseline, see the prerequisites above):

```powershell
$discovery = & opencode.exe debug skill | Out-String
if ($LASTEXITCODE -ne 0) { throw 'OpenCode discovery failed.' }
$skills = ConvertFrom-Json -InputObject $discovery
$mailLens = @($skills | Where-Object name -CEQ 'mail-lens')
$mailLens.Count        # Expected: exactly 1
$mailLens[0].location  # Must be the installer-printed target plus \SKILL.md,
                       # for example C:\Users\synthetic-user\.config\opencode\skills\mail-lens\SKILL.md
```

Check the reported `location`; a duplicate `mail-lens` in another discovery root
can otherwise silently satisfy a name-only check. Each `debug skill` invocation
is a fresh OpenCode process, while a separately running OpenCode session keeps
its already-loaded skills: restart it after any install/update/uninstall change
before asking it to use MailLens.

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
   `-Action Status` and the fresh `opencode.exe debug skill` location check as
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

In OpenCode, ask to use MailLens, for example: “Use MailLens to summarize the
Starbridge contract correspondence from the last 30 days, with sources.” Search
reads subject/plain Body across mounted readable stores and folders, including
Deleted Items/Junk; attachment access is limited to filenames. The
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
