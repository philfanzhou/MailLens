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

First [acquire and verify a chosen full source commit in a fresh checkout](docs/versioned-installation.md).
Retain that complete source and an independent copy of `scripts/install.ps1`.
Then use Windows PowerShell with your verified **absolute** paths:

```powershell
$source = 'C:\Synthetic Source\MailLens'
$installer = Join-Path $source 'scripts\install.ps1'
# Default installation: selected user/global OpenCode skill directory.
& powershell.exe -NoProfile -File $installer -SourcePath $source
# Explicit equivalent:
& powershell.exe -NoProfile -File $installer -User -SourcePath $source
# Alternative: one existing project, selected explicitly.
& powershell.exe -NoProfile -File $installer -Project 'C:\Synthetic Project' -SourcePath $source
```

These are alternative scope examples, not a request to install three copies.
User/global precedence is `OPENCODE_CONFIG_DIR/skills`, otherwise
`XDG_CONFIG_HOME/opencode/skills`, otherwise the actual Windows user's
`.config/opencode/skills`. Project scope uses the selected project's
`.opencode/skills`. `OPENCODE_CONFIG` names a config file; APPDATA is not a fallback.
The installer prints its selected target. Respect existing execution/security
policy; do not add bypass flags or change policy to make installation work.

The installed `mail-lens` directory is an independent copy. Moving the source
checkout does not update it. Restart OpenCode after changes and check which copy
it loads; duplicate names in other roots can affect discovery.

## Status, update, rollback and removal

Use the same selected scope each time; omitted scope means user/global:

```powershell
& powershell.exe -NoProfile -File $installer -Action Status -SourcePath $source
& powershell.exe -NoProfile -File $installer -Action Update -SourcePath $source
# Deliberate rollback: independently acquired and verified older source.
& powershell.exe -NoProfile -File $installer -Action Update -SourcePath 'C:\Synthetic Older Source\MailLens'
& powershell.exe -NoProfile -File $installer -Action Uninstall
```

For a project, add `-Project 'C:\Synthetic Project'` to each action; do not combine
it with `-User`. After source removal, use your retained standalone installer for
Status/Uninstall. Local edits, unmanaged packages and invalid receipts are
preserved and refused, never forcibly overwritten. Read the
[installation lifecycle and recovery guide](docs/installation.md) before handling
reported locks/backups/residue.

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
