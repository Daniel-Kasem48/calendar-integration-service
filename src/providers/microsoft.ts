// Provider payloads extracted from Spantrek's Office365CalendarClient.
import { createHmac } from "node:crypto";
import moment from "moment-timezone";
import { Tenant, fail, secretEquals } from "../config";
import { microsoftRecurrence } from "../recurrence";
import {
  CalendarEvent,
  EventInput,
  EventPage,
  ListOptions,
  Watch,
} from "../types";
import { webhookToken } from "./google";
import { ProviderHttp } from "./http";

export class MicrosoftCalendar {
  constructor(private readonly http: ProviderHttp) {}
  private url(calendar: string): string {
    return `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(calendar)}/calendar/events`;
  }
  private continuation(t: Tenant, calendar: string, value: string): string {
    const url = new URL(value);
    const expected = new URL(this.url(calendar));
    if (
      url.origin !== expected.origin ||
      decodeURIComponent(url.pathname) !== decodeURIComponent(expected.pathname) ||
      url.username ||
      url.password ||
      url.hash
    )
      fail("Unsafe Microsoft continuation URL");
    return url.href;
  }
  private cursor(t: Tenant, calendar: string, url: string): string {
    const data = Buffer.from(this.continuation(t, calendar, url)).toString(
      "base64url",
    );
    return `${data}.${createHmac("sha256", t.apiKey).update(data).digest("base64url")}`;
  }
  private decode(t: Tenant, calendar: string, cursor: string): string {
    const [data, signature, extra] = cursor.split(".");
    if (
      !data ||
      !signature ||
      extra ||
      !secretEquals(
        signature,
        createHmac("sha256", t.apiKey).update(data).digest("base64url"),
      )
    )
      fail("Invalid pagination cursor");
    return this.continuation(
      t,
      calendar,
      Buffer.from(data, "base64url").toString("utf8"),
    );
  }
  private event(e: any): CalendarEvent {
    const date = (d: any) =>
      d?.dateTime
        ? moment.tz(d.dateTime, d.timeZone || "UTC").toISOString()
        : undefined;
    return {
      id: e.id,
      title: e.subject || "",
      description: e.body?.content,
      start: date(e.start),
      end: date(e.end),
      timeZone: e.originalStartTimeZone || e.start?.timeZone || "UTC",
      attendees: (e.attendees || []).map((a: any) => a.emailAddress.address),
      status: e.isCancelled ? "cancelled" : "confirmed",
      allDay: Boolean(e.isAllDay),
    };
  }
  private body(input: Partial<EventInput>): Record<string, unknown> {
    const body: Record<string, unknown> = {};
    if (input.title !== undefined) body.subject = input.title;
    if (input.description !== undefined)
      body.body = { contentType: "text", content: input.description };
    for (const key of ["start", "end"] as const)
      if (input[key] !== undefined)
        body[key] = {
          dateTime: moment(input[key]!).tz(input.timeZone || "UTC").format("YYYY-MM-DDTHH:mm:ss.SSS"),
          timeZone: input.timeZone || "UTC",
        };
    if (input.attendees !== undefined)
      body.attendees = input.attendees.map((address) => ({
        emailAddress: { address },
        type: "required",
      }));
    if (input.recurrence !== undefined)
      body.recurrence =
        input.recurrence === null
          ? null
          : microsoftRecurrence(
              input.recurrence,
              input.start!,
              input.timeZone!,
            );
    return body;
  }
  async list(
    t: Tenant,
    calendar: string,
    options: ListOptions,
  ): Promise<EventPage> {
    if (options.syncToken)
      fail(
        "Microsoft syncToken is not supported; use pageToken or a full refresh",
      );
    const url = options.pageToken
      ? this.decode(t, calendar, options.pageToken)
      : this.url(calendar) + "?$top=250";
    const data = await this.http.request(t, "MICROSOFT", url);
    return {
      items: (data.value || []).map((e: any) => this.event(e)),
      nextPageToken: data["@odata.nextLink"]
        ? this.cursor(t, calendar, data["@odata.nextLink"])
        : undefined,
    };
  }
  async get(t: Tenant, calendar: string, id: string): Promise<CalendarEvent> {
    return this.event(
      await this.http.request(
        t,
        "MICROSOFT",
        this.url(calendar) + "/" + encodeURIComponent(id),
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
        "MICROSOFT",
        this.url(calendar),
        "POST",
        this.body(input),
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
        "MICROSOFT",
        this.url(calendar) + "/" + encodeURIComponent(id),
        "PATCH",
        this.body(input),
      ),
    );
  }
  async remove(t: Tenant, calendar: string, id: string): Promise<void> {
    await this.http.request(
      t,
      "MICROSOFT",
      this.url(calendar) + "/" + encodeURIComponent(id),
      "DELETE",
    );
  }
  async watch(
    t: Tenant,
    calendar: string,
    callbackUrl: string,
  ): Promise<Watch> {
    const expiresAt = new Date(Date.now() + 3 * 86_400_000).toISOString();
    const data = await this.http.request(
      t,
      "MICROSOFT",
      "https://graph.microsoft.com/v1.0/subscriptions",
      "POST",
      {
        changeType: "created,updated,deleted",
        notificationUrl: callbackUrl,
        resource: `/users/${calendar}/events`,
        expirationDateTime: expiresAt,
        clientState: webhookToken(t, "MICROSOFT", calendar),
      },
    );
    return { id: data.id, expiresAt: data.expirationDateTime };
  }
  async renew(t: Tenant, id: string): Promise<Watch> {
    const data = await this.http.request(t, "MICROSOFT", "https://graph.microsoft.com/v1.0/subscriptions/" + encodeURIComponent(id), "PATCH", { expirationDateTime: new Date(Date.now() + 3 * 86_400_000).toISOString() });
    return { id: data.id, expiresAt: data.expirationDateTime };
  }
  async unwatch(t: Tenant, id: string): Promise<void> {
    await this.http.request(
      t,
      "MICROSOFT",
      "https://graph.microsoft.com/v1.0/subscriptions/" +
        encodeURIComponent(id),
      "DELETE",
    );
  }
}
