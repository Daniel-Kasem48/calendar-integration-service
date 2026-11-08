# Calendar Service Implementation Plan

> Execute directly in the new standalone repository using the executing-plans and test-driven-development skills. The user explicitly requested extraction, refactoring and push.

**Goal:** Publish a runnable standalone Google/Outlook calendar service demonstrating GraphQL implementation.
**Architecture:** NestJS/Apollo resolver delegates to tenant-scoped calendar adapters; native HTTP OAuth and provider calls replace the original database coupling.
**Tech Stack:** TypeScript, Node 22+, NestJS, Apollo, node:test.
**Spec:** ../specs/2026-11-08-calendar-service-design.md

## Global Constraints
Secrets remain outside git. Only explicitly allowed calendars are accessible. No edits to the original platform checkout.

## Review Focus
Cross-tenant access, unsafe pagination URLs, partial date updates, invalid recurrence, webhook spoofing or lost notifications. Each is tested at the service or HTTP boundary.

### Task 1: Provider extraction
Files: src/config.ts, src/types.ts, src/recurrence.ts, src/providers/http.ts, src/providers/google.ts, src/providers/outlook.ts, src/calendar.service.ts; test/service.test.cjs.
- [x] Write and run failing behavioral tests.
- [x] Extract recurrence/request mappings; replace entity lookups with explicit arguments.
- [x] Add tenant credentials/calendar allowlists and sanitized HTTP errors.
- [x] Run tests for mapping, recurrence, auth, paging and error behavior.

### Task 2: GraphQL and callbacks
Files: src/graphql.ts, src/webhooks.ts, src/app.ts, src/main.ts; test/http.test.cjs.
- [x] Add failing GraphQL authentication, mutation and webhook tests.
- [x] Implement typed code-first schema, auth and callback delivery.
- [x] Run full tests/build.

### Task 3: Publish
Files: README.md, .env.example, tenants.example.json, Dockerfile, .github/workflows/ci.yml, docs/EXTRACTION.md.
- [x] Document provenance, exact supported scope, credentials, queries, mutations, webhooks and limits.
- [x] Review all code; fix material findings and rerun checks.
- [ ] Scan staged content and create a new GitHub repository under Daniel-Kasem48.
- [ ] Push and verify the remote commit and visibility.
