export type Provider = "GOOGLE" | "MICROSOFT";
export interface Recurrence {
  recurrenceType:
    | "DAY"
    | "WEEK"
    | "AbsoluteMonthly"
    | "RelativeMonthly"
    | "AbsoluteYearly"
    | "RelativeYearly";
  rangeType: "NEVER" | "UNTIL_DATE" | "OCCURRENCE_COUNT";
  interval: number;
  daysOfWeek?: number[];
  dayOfMonth?: number;
  weekDay?: number;
  weekIndex?: number;
  month?: number;
  untilDate?: string;
  occurrenceCount?: number;
}
export interface EventInput {
  title: string;
  description?: string;
  start: string;
  end: string;
  timeZone: string;
  attendees?: string[];
  recurrence?: Recurrence | null;
}
export interface CalendarEvent {
  id: string;
  title: string;
  description?: string;
  start?: string;
  end?: string;
  timeZone?: string;
  attendees: string[];
  status: string;
  allDay: boolean;
}
export interface EventPage {
  items: CalendarEvent[];
  nextPageToken?: string;
  nextSyncToken?: string;
}
export interface ListOptions {
  pageToken?: string;
  syncToken?: string;
}
export interface Watch {
  id: string;
  resourceId?: string;
  expiresAt: string;
}
export type Fetch = typeof fetch;
