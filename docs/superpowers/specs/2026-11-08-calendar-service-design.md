# Calendar integration service

The user selected extraction and refactoring of the existing Google and Outlook calendar integrations, followed by a GitHub push. The service makes their GraphQL implementation experience visible.

## Approach
Extract provider request mappings and recurrence conversion, replacing platform reservation/database lookups with explicit event inputs and tenant configuration. A fresh portfolio rewrite was offered; the user chose extraction. Use native fetch rather than carrying the original SDK/database dependency tree.

## Scope and constraints
- NestJS / Apollo code-first GraphQL queries and mutations.
- Google OAuth refresh-token authentication and Outlook for Microsoft 365 application-permission authentication, isolated by tenant.
- List/get/create/update/delete events; all six original recurrence patterns; explicit calendar allowlists.
- Provider watch creation/deletion and authenticated webhook delivery to a tenant-configured receiver. Return expirations for caller-managed renewal.
- Secrets configured outside git; no original repository history, tenant data, assets, or credentials.
- Preserve provider errors as sanitized errors; no silent mutation failures. Support pagination without following arbitrary caller-supplied URLs.
- No original-platform deployment changes, room import, on-prem Exchange, automatic cross-provider mirroring, or booking policy engine.
- Document app-only Outlook support (business accounts), caller-managed webhook renewal, receiver deduplication, and live-account testing limits.

## Verification
Run deterministic tests with external HTTP replaced at the provider boundary, including real GraphQL HTTP requests, auth failures, cross-tenant/calendar isolation, partial updates, recurrence, pagination, webhook validation, and provider failures. Build and scan the staged files for secrets before publishing.
