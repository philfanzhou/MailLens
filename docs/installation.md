# Windows installation and lifecycle

MailLens targets OpenCode on native Windows, Classic Outlook and Windows
PowerShell 5.1. The source [VERSION](../VERSION) targets `0.1.0-rc.4` under
[issue #50](https://github.com/philfanzhou/MailLens/issues/50). The immutable
[published rc.1](https://github.com/philfanzhou/MailLens/releases/tag/v0.1.0-rc.1)
uses historical schema 1; the current source implements schema 2.
Published [rc.2](https://github.com/philfanzhou/MailLens/releases/tag/v0.1.0-rc.2)
also uses schema 2; published rc.3 adds OpenCode 2.0.28 verification. The rc.4
candidate adds the compact caller, 2000-scalar get default and slim quickstart;
see [compatibility changes](releases/unreleased.md) and
[prepared rc.4 notes](releases/v0.1.0-rc.4.md). Query schema 2 and receipt schema 1
remain unchanged; VERSION alone does not establish publication.
Verify any release's exact tag/commit through
[fixed-source acquisition](versioned-installation.md), rather than treating VERSION
alone as publication or source identity. This page
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
fallback. See the [official skill paths](https://opencode.ai/v2/docs/skills/),
[V2 configuration](https://opencode.ai/v2/docs/config/)
and [pinned OpenCode 2.0.28 resolver](https://github.com/anomalyco/opencode/blob/026b106bc5d41b679ffc413264229c340871e4d3/packages/util/src/global.ts).
Only the selected destination is managed; other scopes and renamed copies remain
caller-owned. Restart OpenCode after a lifecycle change and inspect its skill
loading. Use the [PowerShell 5.1 V2 verification commands](../README.md#windows-quickstart)
to compare the exact installed path, complete instructions and resources. They
start an owned authenticated loopback service and wait for plugin activation
before reading the V2 API; target absence alone is not readiness. Update the
complete MailLens package as well as the CLI. Native V2 discovery evidence is
owned by [issue #37](https://github.com/philfanzhou/MailLens/issues/37); the installer does not
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
When moving to rc.4, update the complete owned package, restart OpenCode, and use
fresh requests/results and private selection runs. Query schema 2 and receipt
schema 1 remain unchanged; the new managed caller/private format 1 indexes need
the complete new package. Omitted get length changes from rc.3's 20000 to 2000
Unicode scalars; explicit lengths remain compatible. Whole-package rollback to
rc.3 restores its older default/instructions and lacks the new managed modes;
do not reuse rc.4 private indexes with older packages. Verify the exact discovered
copy using the OpenCode 2.0.28 steps in the README. Moving from rc.1
also changes query schema 1 to schema 2 and requires fresh schema 2 requests/results.
A rollback to published rc.1 restores schema 1 and
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

## Custom model limits and compaction

After installation/restart, the [self-contained skill](../skills/mail-lens/SKILL.md)
uses the shipped compact caller and labels for common queries. Its initial body
budget is at most three necessary 2000-scalar chunks per answer (6000 total),
fewer when enough. The default runtime get length is 2000; complete-package
update/rollback restores the selected package's instructions and defaults, with
no mailbox/configuration/receipt migration. Live Outlook/model assessment stays
with [#9](https://github.com/philfanzhou/MailLens/issues/9).

For a custom OpenCode provider model, the owner should declare accurate
`limit.context` and `limit.output` for the existing endpoint/model. OpenCode V2
uses plural `providers`; its [model limits](https://opencode.ai/v2/docs/models/#aliases)
and [configuration schema](https://opencode.ai/v2/docs/config/#schema) describe
these values. This illustrative **limit-only fragment** belongs under the user's
existing provider/model entries; IDs and numbers are synthetic, not a complete
provider setup or a recommendation for a particular model:

```json
{
  "providers": {
    "existing-provider-id": {
      "models": {
        "existing-model-id": {
          "limit": { "context": 131072, "output": 8192 }
        }
      }
    }
  }
}
```

Use limits documented by the actual endpoint, including a separate input limit
when applicable; do not raise them to hide an error. Cross-check the pinned
2.0.28 [model Limit schema](https://github.com/anomalyco/opencode/blob/026b106bc5d41b679ffc413264229c340871e4d3/packages/schema/src/config/provider.ts).
MailLens never reads or edits this configuration, selects a model, adds
credentials or changes permissions; this fragment authorizes no automated edit.

Accurate limits let OpenCode make informed [compaction decisions](https://opencode.ai/v2/docs/compaction/).
In the pinned 2.0.28 [compaction ceiling](https://github.com/anomalyco/opencode/blob/026b106bc5d41b679ffc413264229c340871e4d3/packages/core/src/session/compaction.ts#L900),
the usable input ceiling uses the input limit (or context fallback) minus an
explicit/default buffer. Setting output alone does not define that threshold.
Compaction is lossy and cannot guarantee overflow avoidance, especially when
fixed instructions/tool definitions already consume much of the window. Declared
limits, smaller skill load and bounded excerpts do not certify model obedience,
latency or a successful live context run. Verify future host versions against
their own documentation; the supported discovery baseline remains 2.0.28.
