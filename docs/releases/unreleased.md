# Pending compatibility changes (unreleased)

These are development changes, not a published version. Root VERSION and existing
tags/releases are unchanged; publication needs separate maintainer direction under
[the release procedure](../releasing.md).

## Omitted get body length

After [issue #43](https://github.com/philfanzhou/MailLens/issues/43), a schema-2
`get` request that omits `bodyLength` requests 2000 Unicode scalar values instead
of 20000. This reduces the default excerpt from long reply chains. Explicit
lengths from 1 through 100000, including 20000, continue to work. Offset defaults
to zero; scalar units, valid surrogate pairs, truncation flags, exact identity,
consistency checks, response format and status/exit codes are unchanged.

Update or deliberately roll back the complete owned installed package using the
[installation lifecycle](../installation.md), then restart OpenCode. Do not mix
individual scripts and instructions from different snapshots. No mailbox, query,
receipt or configuration migration is required. Rollback restores the selected
older package's own default; explicit lengths remain the way to request a fixed
excerpt size across these packages. See [body usage](../usage.md#bodies-attachment-names-and-sources).

Synthetic checks do not certify native Classic Outlook, live read-only behavior,
latency or model assessment. Those remain separate [issue #9](https://github.com/philfanzhou/MailLens/issues/9)
acceptance work; a stable release additionally needs [issue #11](https://github.com/philfanzhou/MailLens/issues/11).

## Shipped compact caller and smaller instruction load

Current development source includes the reviewed #44 `-View` caller: one fixed
`-RequestPath` command emits strict UTF-8 compact output and its external run
path. Follow-up get uses that selection run and a case-sensitive label; exact
identities remain in an immutable private format 1 index. Complete/partial
selections authorize labels; failed/cancelled runs do not. Raw schema-2 two-path
mode remains available. New callers/indexes need the complete updated package;
old packages cannot consume them. Existing query/receipt schemas remain 2/1.

The #45 Chinese entry is self-contained for browse, keyword/list-only search and
first label get. Advanced/error/multi-chunk references load on demand; regenerated
projection scripts and model-side private-ID variables are no longer needed.
Each answer defaults to <=3 necessary first chunks, <=2000 Unicode scalars each,
<=6000 total, fewer when enough. Additional reads need a deliberate bounded purpose.
See [usage](../usage.md) and [custom OpenCode model limits](../installation.md#custom-model-limits-and-compaction).
These changes reduce avoidable context load without certifying live performance,
model obedience or overflow prevention. No model configuration is inspected or edited.
