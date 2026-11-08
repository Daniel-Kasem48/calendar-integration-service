# Calendar Integration Service

A standalone **NestJS / Apollo GraphQL** service for **Google Calendar** and **Microsoft 365 / Outlook Calendar**, extracted and refactored from my work on Spantrek, a multi-tenant workplace booking platform.

This project demonstrates building GraphQL APIs, rather than only consuming them: code-first schemas, typed inputs and enums, guarded resolvers, tenant-scoped authorization, queries, mutations, pagination, and structured errors.

## What it does

- List, retrieve, create, update, and delete calendar events through one GraphQL contract.
- Convert six recurrence patterns into Google RRULEs and Microsoft pattern/range objects.
- Refresh Google OAuth access tokens and acquire Microsoft application access tokens per tenant.
- Restrict each tenant to an explicit set of calendars.
- Create/stop Google push channels and Microsoft Graph subscriptions.
- Validate provider notifications and forward signed change messages to the tenant's application.

It integrates an application's calendar operations with either provider. It does **not** automatically mirror events between Google and Outlook. See [extraction notes](docs/EXTRACTION.md) for the original components, refactoring decisions, and deliberately excluded Spantrek features.

## Run locally

Requires Node.js 22.13+ (or Node 24) and npm.

```bash
npm ci
cp .env.example .env
cp tenants.example.json tenants.json
# Edit tenants.json with your own credentials and permitted calendars.
# Generate independent API/webhook secrets with: openssl rand -hex 32
chmod 600 tenants.json
npm run build
npm start
```

`GET http://127.0.0.1:3000/health` returns `{"status":"ok"}`. GraphQL is at `POST /graphql`; authenticate with `Authorization: Bearer <tenant-api-key>`. Tenant identity comes from that token, never from a caller-supplied tenant ID. Introspection is enabled for local development and disabled when `NODE_ENV=production`.

Remove a provider's configuration block if you only need the other provider. Configuration loads at startup; restart after changing it. Never commit `tenants.json` or `.env`. Example values are placeholders, and the actual credential files are ignored by Git and Docker.

### Google credentials

Enable Google Calendar API in your Google Cloud project, configure OAuth consent, and obtain a refresh token for your user with calendar read/write access (`https://www.googleapis.com/auth/calendar.events`). Obtain consent with offline access using your OAuth client; this service accepts the resulting refresh token and does not host the consent UI. The authenticated user must have access to each allowed calendar. OAuth apps in testing may have short-lived refresh tokens; use the appropriate publishing status for ongoing use.

[Google OAuth web-server flow](https://developers.google.com/identity/protocols/oauth2/web-server) · [Calendar authorization](https://developers.google.com/workspace/calendar/api/auth)

### Microsoft credentials

Register an application in Microsoft Entra ID. Add the Microsoft Graph **application** permission `Calendars.ReadWrite`, grant administrator consent, and configure `tenantId`, `clientId`, and `clientSecret`. Allowed calendars are Microsoft 365 user/room mailbox addresses. The service uses the client-credentials flow with `https://graph.microsoft.com/.default`.

This extraction supports Microsoft 365 business/organization accounts through application permissions. Personal Outlook.com accounts and delegated user OAuth are outside this version. Restrict the application's mailbox access in Exchange Online as well as configuring this service's allowlist.

[Microsoft client-credentials flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow) · [Create calendar event](https://learn.microsoft.com/en-us/graph/api/user-post-events?view=graph-rest-1.0)

## GraphQL examples

Use a GraphQL client with the Bearer header. These requests run against the configured real calendar and mutations change its events.

```graphql
query Events($provider: CalendarProvider!, $calendar: String!, $page: String) {
  events(provider: $provider, calendarId: $calendar, pageToken: $page) {
    items {
      id
      title
      start
      end
      timeZone
      attendees
      status
      allDay
    }
    nextPageToken
    nextSyncToken
  }
}
```

```json
{ "provider": "GOOGLE", "calendar": "room@example.com" }
```

Use `MICROSOFT` for Outlook. Follow `nextPageToken` until it is null. For Google incremental refresh, save `nextSyncToken` from the final page and supply it as `syncToken` on the next refresh; preserve that same sync token while paging through the refresh. On `SYNC_TOKEN_EXPIRED`, discard the sync token and do a full refresh. Deleted Google events can be tombstones with only an ID and `status: "cancelled"`; dates are nullable. Microsoft supports pagination here, not delta synchronization.

```graphql
mutation {
  createEvent(
    provider: GOOGLE
    calendarId: "room@example.com"
    input: {
      title: "Team planning"
      description: "Weekly planning session"
      start: "2026-12-01T09:00:00+01:00"
      end: "2026-12-01T10:00:00+01:00"
      timeZone: "Europe/Berlin"
      attendees: ["guest@example.com"]
      recurrence: {
        recurrenceType: WEEK
        interval: 1
        daysOfWeek: [2]
        rangeType: OCCURRENCE_COUNT
        occurrenceCount: 8
      }
    }
  ) {
    id
    title
    start
    end
  }
}
```

```graphql
mutation {
  updateEvent(
    provider: GOOGLE
    calendarId: "room@example.com"
    eventId: "YOUR_EVENT_ID"
    input: { end: "2026-12-01T10:30:00+01:00" }
  ) {
    id
    title
    end
  }
}
```

Omitted patch fields stay unchanged. Use `attendees: []` to clear attendees, `description: ""` to clear the description, or `recurrence: null` to remove recurrence. Other explicit nulls are rejected. Extend/end an event using an `end` update; the new end must be after its start. Microsoft returns normalized UTC timestamps while preserving the original timezone for recurrence and updates. If an existing Outlook event uses a Windows timezone name, supply its equivalent IANA `timeZone` when updating its dates. All-day events can be read; all-day date creation/updates are outside this API.

```graphql
query {
  event(
    provider: MICROSOFT
    calendarId: "room@example.com"
    eventId: "YOUR_EVENT_ID"
  ) {
    id
    title
    start
    end
  }
}

mutation {
  deleteEvent(
    provider: MICROSOFT
    calendarId: "room@example.com"
    eventId: "YOUR_EVENT_ID"
  )
}
```

`deleteEvent` deletes the provider event. For an organizer's meeting, provider behavior can notify attendees of cancellation; deleting an attendee's local copy does not cancel the organizer's meeting. Google create/update/delete requests explicitly use `sendUpdates=all`. Mutations are not automatically retried: retrying a create after an ambiguous network failure may create a duplicate. Persist returned event IDs in the calling application.

### Recurrence

| Field                    | Convention                                                                              |
| ------------------------ | --------------------------------------------------------------------------------------- |
| `recurrenceType`         | `DAY`, `WEEK`, `AbsoluteMonthly`, `RelativeMonthly`, `AbsoluteYearly`, `RelativeYearly` |
| `rangeType`              | `NEVER`, `UNTIL_DATE`, `OCCURRENCE_COUNT`                                               |
| `interval`               | 1–999                                                                                   |
| `daysOfWeek` / `weekDay` | Sunday = 0 through Saturday = 6                                                         |
| `weekIndex`              | First = 0 through fourth = 3                                                            |
| `month`                  | January = 0 through December = 11                                                       |
| `dayOfMonth`             | 1–31                                                                                    |
| `untilDate`              | Inclusive `YYYY-MM-DD` in the event's timezone                                          |
| `occurrenceCount`        | 1–999                                                                                   |

These index conventions are retained from Spantrek. Supply fields appropriate to the pattern; for example, `RelativeYearly` needs `month`, `weekDay`, and `weekIndex`.

## Webhooks

Set `PUBLIC_URL` to this service's public HTTPS origin and a tenant `notificationUrl` to your application's HTTPS receiver.

```graphql
mutation {
  watchCalendar(provider: GOOGLE, calendarId: "room@example.com") {
    id
    resourceId
    expiresAt
  }
}
```

The service registers a provider callback under `/webhooks/{google|microsoft}/{tenant}/{calendar}`. It handles Microsoft's validation challenge, verifies Google channel tokens / Microsoft `clientState`, and POSTs tenant-scoped change metadata to `notificationUrl`. Google notifications carry change signals; fetch the changes through `events`. Microsoft notifications also require fetching authoritative event data when reconciling.

The receiver must verify `x-calendar-signature` against `sha256=` plus the hex HMAC-SHA256 of the **raw request body** using that tenant's `webhookSecret`, then durably accept the notification and return a 2xx response within five seconds. Compare signatures in constant time. Failed forwarding returns HTTP 503 to the provider; successful delivery returns 204. Provider retry policies are finite: deduplicate messages and perform periodic reconciliation to recover missed changes. This service has no durable notification queue.

Persist each watch's `id`, optional Google `resourceId`, and actual `expiresAt`. The returned `id` is an opaque signed handle bound to its tenant, provider and calendar; pass it unchanged to stop/renew operations. Callback messages contain the raw provider channel/subscription ID. Rotating `webhookSecret` invalidates saved handles and existing callback tokens, so stop existing watches before rotating it.

**Renew before expiration.** For Google, create a new watch and then stop the old one. For Microsoft, renew the existing subscription with the mutation below; creating a duplicate resource subscription fails with HTTP 409. Google asks for a one-day channel and Microsoft for a three-day subscription; the provider's returned expiration is authoritative. The caller manages renewal; no in-process scheduler is included.

```graphql
mutation {
  renewWatchingCalendar(
    provider: MICROSOFT
    calendarId: "room@example.com"
    watchId: "YOUR_SIGNED_WATCH_HANDLE"
  ) { id expiresAt }
}
```

```graphql
mutation {
  stopWatchingCalendar(
    provider: GOOGLE
    calendarId: "room@example.com"
    watchId: "YOUR_SIGNED_WATCH_HANDLE"
    resourceId: "YOUR_RESOURCE_ID"
  )
}
```

Microsoft only needs `watchId`. Treat subscription identifiers and provider verification tokens as sensitive.

[Google push notifications](https://developers.google.com/workspace/calendar/api/guides/push) · [Microsoft webhook delivery](https://learn.microsoft.com/en-us/graph/change-notifications-delivery-webhooks)

## Errors and deployment

Errors expose stable GraphQL codes including `UNAUTHENTICATED`, `FORBIDDEN`, `BAD_USER_INPUT`, `REAUTH_REQUIRED`, `PROVIDER_FORBIDDEN`, `NOT_FOUND`, `SYNC_TOKEN_EXPIRED`, `RATE_LIMITED`, `UPSTREAM_UNAVAILABLE`, and `UPSTREAM_ERROR`. Provider response bodies and credentials are not returned. `RATE_LIMITED` can include the upstream `retryAfter` header. Upstream calls have a 15-second timeout.

Run behind TLS with ingress rate limits. Keep API keys private; an API key authorizes calendar reads and writes for that tenant. Tenant credentials are kept in a mounted private file and access tokens in process memory. Outbound webhook URLs are administrator-controlled configuration, never GraphQL input. Configure an HTTPS receiver you trust and restrict outbound network access as appropriate for the deployment.

```bash
docker build -t calendar-integration-service .
docker run --rm -p 127.0.0.1:3000:3000 \
  --mount type=bind,source="$PWD/tenants.json",target=/run/secrets/tenants.json,readonly \
  calendar-integration-service
```

Ensure the mounted secret is readable by the container's non-root user. For watches, also pass `-e PUBLIC_URL=https://calendar.your-domain.example`. The Docker image does not embed tenant credentials.

## Verification

```bash
npm test
```

Tests exercise tenant/calendar access control, provider payloads, timezone/recurrence conversion, partial updates, pagination, sanitized failures, real GraphQL HTTP requests, and callback validation/delivery. External provider HTTP is replaced with deterministic fixtures; tests never access real calendars. CI runs on Node 22 and 24. Live provider accounts and the Docker image require separate deployment validation.
