# Automation design references

The runtime borrows small, proven workflow patterns without adding these projects as dependencies:

- [paperless-ngx](https://github.com/paperless-ngx/paperless-ngx): preserve originals, consume incoming mail through explicit rules, and keep processing records that support later audit/reprocessing.
- [Prefect](https://github.com/PrefectHQ/prefect): expose workflow state instead of treating every non-crash as success. Agent Mail runs use `running`, `completed`, `needs_manual_review`, and `failed`.
- [Temporal](https://github.com/temporalio/temporal): retry failure-prone external activities rather than replaying an entire workflow. Only transient `agently-cli` calls are retried; document transformations are not.
- [Docling](https://github.com/docling-project/docling): select parsers by input format and retain source provenance. The manifest stores each accumulated file's size and SHA-256.

These are design references only. The ship-agent runtime remains a small local Python workflow so it can run under Codex, Hermes, or OpenClaw without a workflow server, database, or message broker.
