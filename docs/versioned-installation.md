# Acquire a fixed source snapshot

MailLens is unpublished: [VERSION](../VERSION) contains the development target
`0.1.0`, not an available tag/Release. Select an explicitly trusted complete
**40-character commit SHA** for development. Do not assume the latest `main` is a
reviewed release. End-user installation uses [the existing installer](installation.md);
Git below is acquisition/metadata tooling only.

## Fresh checkout at a chosen commit

Use Git from Windows PowerShell. Fill `$sourceCommit` from your trusted source
record and choose a new absolute destination whose parent already exists. The
empty placeholder deliberately stops the example before any clone:

```powershell
$sourceCommit = '' # Required: trusted full 40-character lowercase commit SHA.
$source = 'C:\Synthetic Source\MailLens'
if ($sourceCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'Choose a full trusted source commit.' }
if (Test-Path -LiteralPath $source) { throw 'Choose a fresh checkout destination.' }
& git clone --no-checkout --no-tags https://github.com/philfanzhou/MailLens.git $source
if ($LASTEXITCODE -ne 0) { throw 'Clone failed; inspect the new destination before retry.' }
& git -C $source fetch --no-tags origin $sourceCommit
if ($LASTEXITCODE -ne 0) { throw 'Exact commit fetch failed.' }
& git -C $source switch --detach $sourceCommit
if ($LASTEXITCODE -ne 0) { throw 'Detached checkout failed.' }
$actual = & git -C $source rev-parse --verify HEAD
if ($LASTEXITCODE -ne 0 -or $actual -cne $sourceCommit) { throw 'Source commit mismatch.' }
$changes = @(& git -C $source status --porcelain --untracked-files=all)
if ($LASTEXITCODE -ne 0 -or $changes.Count -ne 0) { throw 'Source must be clean.' }
```

This detaches only the newly created checkout. Do not reset, clean or switch a
user's existing edited clone. Stop on permission/network/checkout failures;
inspect the new directory before deciding whether a retry needs another fresh
path. Retain `VERSION`, `scripts/install.ps1`, and all of `skills/mail-lens/`.
Inspect the trusted source before installation. A clean tree/commit match is an
identity check, not a security audit or a cryptographic signature.

The [README installation examples](../README.md#prerequisites-and-installation)
now apply using that verified `$source`. Keep a copy of `scripts/install.ps1`
independently, in your own absolute tools path, for example
`C:\Synthetic Tools\mail-lens-install.ps1`; do not overwrite an existing tool.
After source removal, point `$installer` at that retained file for Status/Uninstall.

## Future released versions

This procedure applies **only after a matching release exists**. Publication
under [issue #11](https://github.com/philfanzhou/MailLens/issues/11) also requires
live evidence and explicit maintainer authorization. No example here creates a
tag or a Release.

Choose a tag from the actual [repository Releases](https://github.com/philfanzhou/MailLens/releases)
and obtain its full commit SHA from its reviewed English notes. Read the release
and the remote tag before acquiring source:

```powershell
$tag = '' # Required: actual published tag, not a guessed development version.
$sourceCommit = '' # Required: full SHA recorded in that release's reviewed notes.
if ($tag -notmatch '^v[0-9]+\.[0-9]+\.[0-9]+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$' -or
    $sourceCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'Choose a published tag and its full SHA.' }
& gh release view $tag --repo philfanzhou/MailLens --json tagName,isDraft,name,body,url
if ($LASTEXITCODE -ne 0) { throw 'Release could not be verified.' }
& git ls-remote https://github.com/philfanzhou/MailLens.git "refs/tags/$tag" "refs/tags/$tag^{}"
if ($LASTEXITCODE -ne 0) { throw 'Remote tag could not be verified.' }
```

Check that the release is non-draft, `tagName` and title `name` equal the chosen
tag, and its notes record `$sourceCommit`. Require **both** the tag-object line
and the peeled `refs/tags/<tag>^{}` line; the peeled value must equal the full
recorded commit. A lightweight tag, missing line or identity mismatch stops the
installation. In a fresh checkout acquired as above, also inspect the remote tag
object without overwriting an existing local tag:

```powershell
& git -C $source fetch --no-tags origin "refs/tags/$tag"
if ($LASTEXITCODE -ne 0) { throw 'Tag fetch failed.' }
$tagType = & git -C $source cat-file -t FETCH_HEAD
if ($LASTEXITCODE -ne 0 -or $tagType -cne 'tag') { throw 'Expected an annotated tag object.' }
$tagCommit = & git -C $source rev-parse 'FETCH_HEAD^{commit}'
if ($LASTEXITCODE -ne 0 -or $tagCommit -cne $sourceCommit) { throw 'Tag commit mismatch.' }
```

Recheck the clean HEAD before installing. Annotation alone is not a
cryptographic signature. An identity disagreement is not permission to replace a
tag, use a different commit silently, or fall back to `main`.

## Archives, installed copies and rollback

A future GitHub source archive is another source-only distribution of the
verified tag/commit. Retain the complete extracted source tree and verify its
origin against the same release record. An extracted archive lacks `.git`;
the installation receipt honestly records unknown Git provenance when metadata
is unavailable. It must not be presented as a verified Git checkout merely
because its directory was named after a version. There is no standalone binary,
extra release asset, package registry or remote one-click installer.

The installer copies the complete package into the selected discovery root;
source moves/deletion do not remove the installed copy. Status/Uninstall can use
the retained standalone script. Updates are explicit and require the complete
selected source. Keep separately verified snapshots for deliberate rollback and
use `-Action Update -SourcePath <older-source>` as documented in the
[lifecycle guide](installation.md#status-update-deliberate-rollback-and-removal).
Never move a historical tag to roll back an installed copy.

Check the exact printed scope/target and loaded skill location. Selected removal
does not erase duplicate names/renamed copies in other roots. Preserve edited or
unmanaged packages and configuration; resolve ownership conflicts and recovery
residue using the [existing recovery instructions](installation.md#failure-and-recovery).
