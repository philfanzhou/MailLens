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
