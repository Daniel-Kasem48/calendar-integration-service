// Provider payloads extracted from Spantrek's GoogleCalendarClient.
import { createHmac, randomUUID } from "node:crypto";
import { Tenant } from "../config";
import { googleRecurrence } from "../recurrence";
import {
  CalendarEvent,
  EventInput,
  EventPage,
  ListOptions,
  Watch,
} from "../types";
import { ProviderHttp } from "./http";

export class GoogleCalendar {
  constructor(private readonly http: ProviderHttp) {}
  private url(calendar: string, suffix = ""): string {
    return `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(calendar)}/events${suffix}`;
  }
  private event(e: any): CalendarEvent {
    return {
      id: e.id,
      title: e.summary || "",
      description: e.description,
      start: e.start?.dateTime || e.start?.date,
      end: e.end?.dateTime || e.end?.date,
      timeZone: e.start?.timeZone || "UTC",
      attendees: (e.attendees || []).map((a: any) => a.email),
      status: e.status || "confirmed",
      allDay: Boolean(e.start?.date),
    };
  }
  private body(input: Partial<EventInput>): Record<string, unknown> {
    const body: Record<string, unknown> = {};
    if (input.title !== undefined) body.summary = input.title;
    if (input.description !== undefined) body.description = input.description;
    if (input.start !== undefined)
      body.start = { dateTime: input.start, timeZone: input.timeZone || "UTC" };
    if (input.end !== undefined)
      body.end = { dateTime: input.end, timeZone: input.timeZone || "UTC" };
    if (input.attendees !== undefined)
      body.attendees = input.attendees.map((email) => ({ email }));
    if (input.recurrence !== undefined)
      body.recurrence =
        input.recurrence === null
          ? []
          : [googleRecurrence(input.recurrence, input.start!, input.timeZone!)];
    return body;
  }
  async list(
    t: Tenant,
    calendar: string,
    options: ListOptions,
  ): Promise<EventPage> {
    const url = new URL(this.url(calendar));
    url.searchParams.set("maxResults", "250");
    url.searchParams.set("showDeleted", "true");
    if (options.pageToken) url.searchParams.set("pageToken", options.pageToken);
    if (options.syncToken) url.searchParams.set("syncToken", options.syncToken);
    const data = await this.http.request(t, "GOOGLE", url.href);
    return {
      items: (data.items || []).map((e: any) => this.event(e)),
      nextPageToken: data.nextPageToken,
      nextSyncToken: data.nextSyncToken,
    };
  }
  async get(t: Tenant, calendar: string, id: string): Promise<CalendarEvent> {
    return this.event(
      await this.http.request(
        t,
        "GOOGLE",
        this.url(calendar, "/" + encodeURIComponent(id)),
      ),
    );
  }
  async create(
    t: Tenant,
    calendar: string,
    input: EventInput,
  ): Promise<CalendarEvent> {
    return this.event(
      await this.http.request(
        t,
        "GOOGLE",
        this.url(calendar) + "?sendUpdates=all",
        "POST",
        { visibility: "private", ...this.body(input) },
      ),
    );
  }
  async update(
    t: Tenant,
    calendar: string,
    id: string,
    input: Partial<EventInput>,
  ): Promise<CalendarEvent> {
    return this.event(
      await this.http.request(
        t,
        "GOOGLE",
        this.url(calendar, "/" + encodeURIComponent(id)) + "?sendUpdates=all",
        "PATCH",
        this.body(input),
      ),
    );
  }
  async remove(t: Tenant, calendar: string, id: string): Promise<void> {
    await this.http.request(
      t,
      "GOOGLE",
      this.url(calendar, "/" + encodeURIComponent(id)) + "?sendUpdates=all",
      "DELETE",
    );
  }
  async watch(
    t: Tenant,
    calendar: string,
    callbackUrl: string,
  ): Promise<Watch> {
    const id = randomUUID();
    const token = webhookToken(t, "GOOGLE", calendar);
    const result = await this.http.request(
      t,
      "GOOGLE",
      this.url(calendar, "/watch"),
      "POST",
      {
        id,
        type: "web_hook",
        address: callbackUrl,
        token,
        params: { ttl: "86400" },
      },
    );
    return {
      id: result.id,
      resourceId: result.resourceId,
      expiresAt: new Date(Number(result.expiration)).toISOString(),
    };
  }
  async unwatch(t: Tenant, id: string, resourceId: string): Promise<void> {
    await this.http.request(
      t,
      "GOOGLE",
      "https://www.googleapis.com/calendar/v3/channels/stop",
      "POST",
      { id, resourceId },
    );
  }
}
export function webhookToken(
  t: Tenant,
  provider: string,
  calendar: string,
): string {
  return createHmac("sha256", t.webhookSecret)
    .update(JSON.stringify([t.id, provider, calendar]))
    .digest("base64url");
}
