# Future manually authorized source releases

MailLens currently has no published release. Root [VERSION](../VERSION) contains
`0.1.0-rc.1`, an **unpublished prerelease candidate** prepared for the Windows
testing channel under
[issue #25](https://github.com/philfanzhou/MailLens/issues/25). This document,
the [blank notes template](releases/TEMPLATE.md) and the
[prepared RC notes input](releases/v0.1.0-rc.1.md) are publication inputs, not
evidence that any version/tag/Release exists. [Issue #11](https://github.com/philfanzhou/MailLens/issues/11)
owns stable publication; it follows [live validation #9](https://github.com/philfanzhou/MailLens/issues/9)
and [documentation #10](https://github.com/philfanzhou/MailLens/issues/10).
A prerelease RC channel may be published before #9 only under #25's recorded
maintainer direction; the channel gates below keep the two channels separate.

Use the [first-release contract](https://github.com/philfanzhou/MailLens/issues/1#single-authoritative-first-release-technical-contract),
[project conventions](../CONTRIBUTING.md#version-and-protected-branches) and the
[fixed reference's immutable-tag/exact-commit conventions](https://github.com/philfanzhou/ForgeSteward/blob/0693af9bf3bc2067d856fc0032e860c6754a064a/docs/releasing.md).
MailLens distributes the complete Git source snapshot and GitHub automatic
source archives, with no registry, marketplace, standalone binary or extra asset.
No commit or CI success automatically authorizes a tag or Release.

## Approve and validate one final commit

Before publication, persist the following using sanitized English evidence in
#11, or in the task that owns the channel (the RC under
[issue #25](https://github.com/philfanzhou/MailLens/issues/25) while it is open):

1. Explicit maintainer decision for the final version and exact full reviewed
   main commit SHA, plus **separate explicit authorization to publish**. Root
   VERSION alone is not that authorization. Change VERSION only when directed,
   through a feature PR; no direct/force push to main or bypass is allowed.
2. Reviewed, merged implementation/documentation PRs and compatibility/migration
   assessment, complete package/installer and installation/update/rollback/removal
   validation. Required main checks `windows-tests` and `opencode-discovery` must
   actually pass on that **exact main SHA**, under effective strict protection.
3. Final English notes completed from the blank template (or the prepared RC
   notes input for a prerelease channel) with evidence links, full SHA,
   platform/version scope, limitations, compatibility/migration and
   [fixed installation](versioned-installation.md) / [lifecycle](installation.md)
   links. Use links pinned to the selected commit in public notes.
4. The channel gate for that exact version:
   - **Prerelease RC channel** (for example `v0.1.0-rc.1` under #25's recorded
     maintainer direction): completed #9 live validation is deliberately **not**
     a prerequisite. Pending #9 is an acknowledged RC limitation recorded in the
     notes, and the RC exists to supply the fixed candidate for #9. Synthetic
     hosted checks and discovery still do not certify live behavior, and the
     notes must present the #9 evidence row as pending.
   - **Stable channel** (for example `v0.1.0`): requires completed #9 with
     actual native Windows/OpenCode/Classic Outlook evidence and human signoff
     for that exact candidate, including read-only before/after state and model
     assessment limits. Hosted COM doubles/discovery do not supply live signoff.
     Keep live content and identifiers private, outside Git/public logs.

Any final change, even documentation/version metadata, changes the commit.
Assess/revalidate affected paths and obtain the channel gate's exact-candidate
evidence/signoff on the resulting final SHA before publishing; do not silently
reuse evidence for a different commit. Promoting an RC line to stable (for
example `0.1.0-rc.1` to `0.1.0`) therefore requires separate maintainer
direction, a reviewed VERSION-change PR and fresh assessment/revalidation plus
#9 signoff on the new SHA; RC evidence does not carry over. Record review head,
merge SHA and exact-main CI separately. A source commit cannot contain its own
final SHA: prepare notes externally after selecting the final commit, rather
than making another commit to insert its SHA. If published source needs a fix,
prepare a reviewed fix and a newly authorized version (for example
`v0.1.0-rc.2`); never move, overwrite or recreate a published tag.

## Read-only preflight

Use a fresh maintainer checkout, Git and authenticated `gh` with existing
permissions. Replace the empty placeholders with the recorded approved values.
These examples intentionally stop while inputs are blank; paths are separate
argument values. No sample value grants publication authorization.

```powershell
$repo = 'philfanzhou/MailLens'
$releaseCommit = '' # Required: approved full lowercase SHA from #11.
$releaseVersion = '' # Required: maintainer-approved VERSION at that SHA.
$releaseCheckout = 'C:\Synthetic Release Source\MailLens'
$notesFile = 'C:\Synthetic Release Notes\notes.md' # Completed external English notes.
if ($releaseCommit -cnotmatch '^[0-9a-f]{40}$' -or
    $releaseVersion -notmatch '^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$') { throw 'Resolve approved publication inputs first.' }
$tag = 'v' + $releaseVersion
& git -C $releaseCheckout fetch --no-tags origin main
if ($LASTEXITCODE -ne 0) { throw 'Main fetch failed.' }
& git -C $releaseCheckout show "${releaseCommit}:VERSION"
if ($LASTEXITCODE -ne 0) { throw 'Candidate VERSION unavailable.' }
& gh api "repos/$repo/commits/$releaseCommit/check-runs" --paginate
if ($LASTEXITCODE -ne 0) { throw 'Candidate checks unavailable.' }
& gh api "repos/$repo/rules/branches/main" --paginate
if ($LASTEXITCODE -ne 0) { throw 'Effective protection unavailable.' }
& git -C $releaseCheckout ls-remote origin "refs/tags/$tag" "refs/tags/$tag^{}"
if ($LASTEXITCODE -ne 0) { throw 'Tag preflight failed.' }
& gh api "repos/$repo/releases" --paginate
if ($LASTEXITCODE -ne 0) { throw 'Release preflight failed.' }
```

Compare VERSION exactly with the approved value and tag (`v` plus VERSION).
Verify the selected SHA was actually reviewed/merged on main (for example,
`git merge-base --is-ancestor <full-SHA> origin/main`, exit 0), not just reachable
on some feature branch. Inspect both successful checks' source and tested commit;
use actual [development validation](development.md), not old green runs. Read
back main/tag rules and preserve all protection and the empty bypass lists.
The active [v* tag ruleset](https://github.com/philfanzhou/MailLens/rules/24914289)
protects updates/deletion of existing tags. This does not establish that GitHub
release-level immutability is enabled.

Inspect all remote release pages for the selected tag and distinguish confirmed
absence from permissions/network/API failure. Inspect local tags too. Both a new
local/remote tag and a new release must be absent before creation; an existing
identity goes through the recovery checks below. Do not guess absence from a
failed command. Ensure the checkout is clean and at the recorded commit, the
external notes are complete, and all approvals/gates above still hold.

## Authorized publication only

Run this section only under #11's explicit publication authorization and passed
preflight. Never use latest moving main as an implicit target:

```powershell
& git -C $releaseCheckout tag -a $tag $releaseCommit -m "MailLens $tag; commit $releaseCommit"
if ($LASTEXITCODE -ne 0) { throw 'Tag creation failed; inspect existing local/remote state.' }
$localType = & git -C $releaseCheckout cat-file -t "refs/tags/$tag"
if ($LASTEXITCODE -ne 0 -or $localType -cne 'tag') { throw 'Expected annotated tag.' }
$localCommit = & git -C $releaseCheckout rev-parse "refs/tags/$tag^{commit}"
if ($LASTEXITCODE -ne 0 -or $localCommit -cne $releaseCommit) { throw 'Local tag mismatch.' }
& git -C $releaseCheckout push origin "refs/tags/${tag}:refs/tags/${tag}"
if ($LASTEXITCODE -ne 0) { throw 'Tag push failed/unknown; read remote state before retry.' }
& git -C $releaseCheckout ls-remote origin "refs/tags/$tag" "refs/tags/$tag^{}"
if ($LASTEXITCODE -ne 0) { throw 'Remote tag readback failed.' }
& git -C $releaseCheckout fetch --no-tags origin "refs/tags/$tag"
if ($LASTEXITCODE -ne 0) { throw 'Remote tag fetch failed.' }
$remoteType = & git -C $releaseCheckout cat-file -t FETCH_HEAD
if ($LASTEXITCODE -ne 0 -or $remoteType -cne 'tag') { throw 'Remote object is not annotated.' }
$remoteCommit = & git -C $releaseCheckout rev-parse 'FETCH_HEAD^{commit}'
if ($LASTEXITCODE -ne 0 -or $remoteCommit -cne $releaseCommit) { throw 'Remote commit mismatch.' }
```

Manually compare the remote tag-object SHA to the local tag object and its peeled
SHA to `$releaseCommit` before proceeding. An annotated tag is not automatically
cryptographically signed. Push this single explicit ref, never `--tags`, force,
or a tag replacement/deletion. A readback disagreement stops publication.

Then create the matching non-draft source Release with the **tag alone as title**:

```powershell
& gh release create $tag --repo $repo --verify-tag --title $tag --notes-file $notesFile
if ($LASTEXITCODE -ne 0) { throw 'Release creation failed/unknown; read back before retry.' }
& gh release view $tag --repo $repo --json tagName,name,isDraft,isPrerelease,body,url
if ($LASTEXITCODE -ne 0) { throw 'Release readback failed.' }
```

A prerelease RC channel (for example `v0.1.0-rc.1` under #25) must add
`--prerelease` to creation and verify that channel on readback.
`--verify-tag` requires the pre-existing remote
tag instead of permitting implicit tag creation from latest main. Check tagName
and title equal `$tag`, `isDraft:false`, the intended prerelease status, notes'
full SHA/evidence/links and the actual returned URL. Re-read the remote annotated
tag and peeled commit after release creation. Only this confirmed pair completes
publication; persist the actual URL and result in #11. Do not treat local notes,
a pushed tag alone or a draft as delivery.

## Failures, unknown outcomes and corrections

| Observed state | Next step under existing authorization |
| --- | --- |
| Required checks/version/publish approval missing (stable channel additionally: #9 live signoff) | Keep the owning issue open with exact missing evidence; do not tag/publish. |
| Permission/network/API failure or unknown write outcome | Query the remote tag/Release before retry. If unreadable, stop with the blocker; never assume success or absence. |
| Local tag exists, remote absent | Verify annotation/object and exact SHA against the approved record; push only that verified ref. Do not recreate a tag to hide a mismatch. |
| Remote tag exists, Release absent | Verify annotation/peeled SHA and approvals. Retain that tag; retry only the matching Release creation with `--verify-tag`. |
| Matching non-draft Release already exists after an uncertain response | Verify title/channel/notes/URL and tag object/peeled SHA; record the confirmed result rather than create a duplicate. |
| Existing local/remote tag, Release or recorded commit disagrees | Stop and report the identity conflict. Never move, overwrite, delete or recreate a historical tag. |
| Tag pushed successfully, Release creation fails | Preserve the pushed tag. Read back, then retry the same Release against the same SHA when the external failure is resolved. |
| Released source needs a fix | Prepare a reviewed fix and a newly authorized version/release. Explain known limits and installed rollback; keep historical source identity. |

A notes erratum must be labeled and must not change the tagged source identity.
Installed rollback uses an older verified source via explicit Update, as in the
[lifecycle guide](installation.md#status-update-deliberate-rollback-and-removal),
without moving tags or undoing unrelated user data/configuration. Only clean up
owned workspaces/branches after actual merge/publication and preservation of
needed evidence; never touch user installations or private live evidence.

Command references: [git tag](https://git-scm.com/docs/git-tag),
[gh release create](https://cli.github.com/manual/gh_release_create) and
[GitHub source releases](https://docs.github.com/en/repositories/releasing-projects-on-github/about-releases).
