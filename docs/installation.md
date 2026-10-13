# Windows installation and lifecycle

MailLens targets OpenCode on native Windows, Classic Outlook and Windows
PowerShell 5.1. The source `VERSION` is the target package version; it still reads
`0.1.0-rc.1`. The [published rc.1](https://github.com/philfanzhou/MailLens/releases/tag/v0.1.0-rc.1)
uses historical schema 1; the current schema 2 source is untagged development
content. Select the exact trusted commit/package, rather than treating VERSION
alone as source identity. This page
covers the installer from [issue #6](https://github.com/philfanzhou/MailLens/issues/6).
The [first-release contract](https://github.com/philfanzhou/MailLens/issues/1#single-authoritative-first-release-technical-contract)
defines the skill and query interfaces.

Obtain and retain a trusted fixed source checkout/archive containing `VERSION`,
`scripts/install.ps1` and the complete `skills/mail-lens/` tree. See
[fixed-source acquisition](versioned-installation.md) and the
[future authorized release procedure](releasing.md).
The installer needs only built-in PowerShell/.NET. Git is optional source metadata;
Python, development Pester modules, Office interop packages, model calls and
separate credentials are unnecessary. Respect your existing execution policy;
the installer does not alter policy or OpenCode configuration.

## Select a destination

Run these commands in Windows PowerShell. Replace the sample absolute paths with
your fixed checkout and, for project scope, an **existing** project directory.
Paths are separate argument values; quote paths containing spaces.

```powershell
$installer = 'C:\Synthetic Source\MailLens\scripts\install.ps1'
$source = 'C:\Synthetic Source\MailLens'
& powershell.exe -NoProfile -File $installer -SourcePath $source
# Explicit equivalent of the default user/global scope:
& powershell.exe -NoProfile -File $installer -User -SourcePath $source
# One selected project:
& powershell.exe -NoProfile -File $installer -Project 'C:\Synthetic Project' -SourcePath $source
```

`-Action` defaults to `Install`; the only values are `Install`, `Status`, `Update`
and `Uninstall`. `-User` and `-Project` are mutually exclusive. SourcePath defaults
to the checkout containing the script, and an explicit source must be absolute.
The installer prints the selected scope and exact target.

| Scope | Destination |
| --- | --- |
| User, with nonempty `OPENCODE_CONFIG_DIR` | `<OPENCODE_CONFIG_DIR>\skills\mail-lens` |
| User, otherwise with nonempty `XDG_CONFIG_HOME` | `<XDG_CONFIG_HOME>\opencode\skills\mail-lens` |
| User, otherwise | `<actual Windows user home>\.config\opencode\skills\mail-lens` |
| Project | `<project>\.opencode\skills\mail-lens` |

The selected environment override must be absolute and valid. `OPENCODE_CONFIG`
selects a configuration **file**, not this skill destination. APPDATA is not a
fallback. See the [official skill paths](https://opencode.ai/docs/skills/),
[custom configuration directory](https://opencode.ai/docs/config/#custom-directory)
and [pinned OpenCode 1.18.30 resolver](https://github.com/anomalyco/opencode/blob/5cd8e68fdd72b27818d26d168b9c7a06b359567e/packages/core/src/global.ts).
Only the selected destination is managed; other scopes and renamed copies remain
caller-owned. Restart OpenCode after a lifecycle change and inspect its skill
loading. Native pinned CLI discovery is verified separately by
[issue #8](https://github.com/philfanzhou/MailLens/issues/8); the installer does not
promise an atomic snapshot for an already running OpenCode process.

## Status, update, deliberate rollback and removal

```powershell
& powershell.exe -NoProfile -File $installer -Action Status -SourcePath $source
& powershell.exe -NoProfile -File $installer -Action Update -SourcePath $source
# Deliberate rollback uses a separately retained older fixed checkout/archive:
& powershell.exe -NoProfile -File $installer -Action Update -SourcePath 'C:\Synthetic Older Source\MailLens'
& powershell.exe -NoProfile -File $installer -Action Uninstall
# The same actions support a project:
& powershell.exe -NoProfile -File $installer -Action Status -Project 'C:\Synthetic Project'
& powershell.exe -NoProfile -File $installer -Action Uninstall -Project 'C:\Synthetic Project'
$LASTEXITCODE
```

Status creates no directories, lock, receipt or other file. Identical managed
install/update preserves all package and receipt bytes/timestamps, including its
original source metadata. A changed package/version requires explicit Update;
Update cannot adopt an absent or unmanaged directory. An older source is allowed
for deliberate rollback; it changes the installed copy without changing tags.
For schema 1 -> 2, update the complete package even when VERSION is unchanged;
changed file hashes still require Update. Restart OpenCode and create fresh
schema 2 requests/results. A rollback to published rc.1 restores schema 1 and
its older defaults; do not continue using schema 2 callers with that package.
No mailbox/configuration/receipt migration is required; receipts stay schema 1.

The entire package is copied byte-for-byte, including scripts, references and
empty directories. Installed relative references and launcher/module paths remain
valid after moving/removing the source. Keep `install.ps1` independently for status
and uninstall: that standalone script does not require the source package or
VERSION for these operations. An unavailable source is reported by installed
status. Uninstall removes only the unchanged owned selected package; an absent
repeat succeeds. It leaves configuration, unrelated skills and parent directories.

Exit 0 means completed, unchanged or truthfully absent. Exit 1 means validation,
conflict, lock or recovery failure. Read warnings as well as the exit code: cleanup
can fail **after** a successful commit, leaving explicitly reported residue.

## Ownership and conflicts

The installed `.mail-lens-install.json` receipt uses schema 1 and binds MailLens,
`mail-lens`, VERSION, source metadata and complete relative SHA-256 file hashes
plus directory inventory. Git metadata records commit/dirty when available;
archives or unavailable Git use explicit unknown values. It contains no mail data
or credentials. Receipt input uses strict UTF-8 with an optional leading BOM;
output uses UTF-8 without a BOM. The receipt itself is separately validated and
excluded from its file hashes. The receipt is accidental-modification evidence,
not a signed authenticity guarantee; trust the source you select.

Same-name files, unmanaged directories, local edits, corrupt receipts, missing or
extra files/directories, unsafe manifest paths and links/junctions/reparse points
in the relevant trees/ancestors are rejected. Preserve or move your edits manually
before continuing. There is no force switch, automatic adoption or migration from
an older receipt format. Reported selected-name conflicts are preserved; the
installer does not search/delete every other discovery root or renamed copy.
Source and destination package paths must not overlap.

## Failure and recovery

The [authoritative lifecycle table in issue #6](https://github.com/philfanzhou/MailLens/issues/6)
defines the operation states. A mutator obtains `.mail-lens.lock` with an exclusive
file handle and revalidates source/target before staging and swaps. A collision
fails without stealing the live lock. Different selected targets use independent
locks/transaction directories; read-only status has no snapshot guarantee.

Staged `new-package` and an update/uninstall `old-package` backup live in an owned
`.mail-lens-transaction-<unique-id>` directory directly under the selected config
root, outside `skills/*/SKILL.md` discovery. The original complete package remains
discoverable until its move; publication uses a complete directory move rather
than a partially copied SKILL.md. Ordinary failure or cooperative Ctrl+C before
commit restores the prior complete package/receipt. Cancellation is checked
between phases; forcibly terminating the host or losing power has no automatic
recovery guarantee.

If restoration fails, the error prints the exact transaction and intact prior
backup path. Preserve them. Later mutation refuses unfinished transaction residue.
If commit succeeded but cleanup failed, output says completed and the warning
states `committed=True` with the exact retained residue. Inspect the current target
before deciding what to recover; do not blindly repeat a destructive commit.

For manual recovery, stop the affected installer/OpenCode processes, preserve
both copies, inspect the current target and backup, and choose the intended complete
copy. Move only the reported owned complete package back to the selected target
when needed, retaining edited/unexplained content elsewhere. Remove owned residue
and a leftover lock file only after confirming process termination and ownership.
Never delete a live lock, unknown backup, other skill or entire configuration root.
Hostile concurrent writers, power-loss recovery, signed authenticity, exhaustive
other-scope discovery and forensic erasure of history are outside the guarantee.

Installer operations do not call Outlook, a model, attachment content, registry
or security-policy APIs. Query operation and live read-only state validation remain
separate responsibilities of the [skill](../skills/mail-lens/SKILL.md) and
[issue #9](https://github.com/philfanzhou/MailLens/issues/9).
