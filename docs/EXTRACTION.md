# Extraction notes

This standalone service was extracted and refactored from the Google Calendar and Office365 integration modules in Spantrek. It is a narrower, independently runnable service rather than a copy of the original platform.

| Original component                            | Standalone equivalent                                        |
| --------------------------------------------- | ------------------------------------------------------------ |
| `GoogleCalendarClient`                        | `src/providers/google.ts`                                    |
| `Office365CalendarClient`                     | `src/providers/microsoft.ts`                                 |
| `GoogleRecurrencePatternAdapter`              | `googleRecurrence` in `src/recurrence.ts`                    |
| `Office365RecurrencePatternAdapter`           | `microsoftRecurrence` in `src/recurrence.ts`                 |
| `GoogleAuthenticationService`                 | Google refresh-token flow in `src/providers/http.ts`         |
| `ApplicationPermissionsAuthenticationService` | Microsoft client-credentials flow in `src/providers/http.ts` |
| `CalendarProviderFactory`                     | Provider selection in `src/calendar.service.ts`              |
| Guarded NestJS calendar resolver              | Tenant-authenticated resolver in `src/graphql.ts`            |

Provider field mappings, the six recurrence patterns, Google channels, Microsoft subscriptions, and the two authentication mechanisms come from the existing integration. Database entities, repositories, request-scoped module lookups, office resource imports, the booking-policy engine, other integrations, assets, and source Git history were excluded.

The extraction fixes partial-update assumptions (missing attendee arrays and changes to only one date), removes shared mutable credential state, validates tenant calendar access, and propagates sanitized provider errors. Google recurrence end dates are emitted as UTC datetimes; Microsoft's recurrence start date is populated, and its event timestamps are normalized to UTC for output while preserving local timezones in write payloads. Native fetch replaces the SDK wrappers. There is no claim of unchanged behavior across the entire Spantrek application.

Only Microsoft application permissions are retained. The original delegated permissions and on-premises Exchange implementations are outside this extraction. Webhooks notify a tenant's receiver; the receiver owns reservation reconciliation, persistence, deduplication, and recovery polling. This service does not synchronize Google and Outlook directly with each other.

Tests replace only external HTTP; they run the actual service, payload conversion, GraphQL HTTP requests, and callback handlers. No live Google or Microsoft tenant was used for validation.

Review fixes include Microsoft weekly week-boundary alignment, supported event subscription resource paths, in-place subscription renewal, signed watch handles to bind cancellation to its tenant/calendar, and updates to valid untitled provider events.
