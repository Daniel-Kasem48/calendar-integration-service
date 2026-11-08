import moment from "moment-timezone";
import { createHmac } from "node:crypto";
import { TenantRegistry, Tenant, fail, secretEquals } from "./config";
import { EventInput, Fetch, ListOptions, Provider, Watch } from "./types";
import { GoogleCalendar } from "./providers/google";
import { OutlookCalendar } from "./providers/outlook";
import { ProviderHttp } from "./providers/http";
import { googleRecurrence } from "./recurrence";

export class CalendarService {
  private readonly google: GoogleCalendar;
  private readonly outlook: OutlookCalendar;
  constructor(
    private readonly registry: TenantRegistry,
    http: Fetch = fetch,
  ) {
    const transport = new ProviderHttp(http);
    this.google = new GoogleCalendar(transport);
    this.outlook = new OutlookCalendar(transport);
  }
  private client(provider: Provider) {
    return provider === "GOOGLE" ? this.google : this.outlook;
  }
  private validate(input: Partial<EventInput>): void {
    if (
      input.title !== undefined &&
      (typeof input.title !== "string" ||
        !input.title.trim() ||
        input.title.length > 1000)
    )
      fail("A nonempty title of at most 1000 characters is required");
    if (
      input.description !== undefined &&
      (typeof input.description !== "string" ||
        input.description.length > 20_000)
    )
      fail("Description is too long");
    for (const date of [input.start, input.end])
      if (
        date !== undefined &&
        (typeof date !== "string" ||
          !/T.*(?:Z|[+-]\d{2}:\d{2})$/.test(date) ||
          !moment.parseZone(date, moment.ISO_8601, true).isValid())
      )
        fail("Dates must be valid ISO timestamps with an explicit offset");
    if (
      input.start &&
      input.end &&
      new Date(input.end) <= new Date(input.start)
    )
      fail("Event end must be after start");
    if (
      input.timeZone !== undefined &&
      (typeof input.timeZone !== "string" || !moment.tz.zone(input.timeZone))
    )
      fail("A valid IANA timeZone is required");
    if (
      input.attendees !== undefined &&
      (!Array.isArray(input.attendees) ||
        input.attendees.length > 200 ||
        input.attendees.some(
          (e) =>
            typeof e !== "string" ||
            e.length > 254 ||
            !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e),
        ))
    )
      fail("Attendees must contain valid email addresses (at most 200)");
    if (input.recurrence)
      googleRecurrence(input.recurrence, input.start!, input.timeZone!);
  }
  async list(
    tenant: string,
    provider: Provider,
    calendar: string,
    options: ListOptions = {},
  ) {
    const t = this.registry.calendar(tenant, provider, calendar);
    return this.client(provider).list(t, calendar, options);
  }
  async get(tenant: string, provider: Provider, calendar: string, id: string) {
    return this.client(provider).get(
      this.registry.calendar(tenant, provider, calendar),
      calendar,
      id,
    );
  }
  async create(
    tenant: string,
    provider: Provider,
    calendar: string,
    input: EventInput,
  ) {
    const t = this.registry.calendar(tenant, provider, calendar);
    for (const field of ["title", "start", "end", "timeZone"] as const)
      if (input[field] == null) fail(`Missing event ${field}`);
    this.validate(input);
    return this.client(provider).create(t, calendar, input);
  }
  async update(
    tenant: string,
    provider: Provider,
    calendar: string,
    id: string,
    input: Partial<EventInput>,
  ) {
    const t = this.registry.calendar(tenant, provider, calendar);
    const client = this.client(provider);
    if (!Object.keys(input).length)
      fail("At least one event field must be supplied");
    this.validate({ ...input, recurrence: undefined });
    let patch = { ...input };
    if (
      input.start !== undefined ||
      input.end !== undefined ||
      input.recurrence ||
      input.timeZone !== undefined
    ) {
      const old = await client.get(t, calendar, id);
      if (old.allDay) fail("Updating all-day event dates is not supported");
      const merged = {
        ...old,
        ...input,
        timeZone: input.timeZone || old.timeZone || "UTC",
      };
      this.validate({ start: merged.start, end: merged.end, timeZone: merged.timeZone, recurrence: input.recurrence });
      if (input.start !== undefined || input.end !== undefined)
        patch.timeZone = merged.timeZone;
      if (
        input.recurrence ||
        (input.timeZone !== undefined &&
          input.start === undefined &&
          input.end === undefined)
      )
        patch = {
          ...patch,
          start: merged.start,
          end: merged.end,
          timeZone: merged.timeZone,
        };
    }
    return client.update(t, calendar, id, patch);
  }
  async remove(
    tenant: string,
    provider: Provider,
    calendar: string,
    id: string,
  ): Promise<boolean> {
    await this.client(provider).remove(
      this.registry.calendar(tenant, provider, calendar),
      calendar,
      id,
    );
    return true;
  }
  async watch(
    tenant: string,
    provider: Provider,
    calendar: string,
    publicUrl: string,
  ) {
    const t = this.registry.calendar(tenant, provider, calendar);
    if (!t.notificationUrl)
      fail("Configure the tenant notificationUrl before subscribing");
    const callback = `${publicUrl.replace(/\/$/, "")}/webhooks/${provider.toLowerCase()}/${encodeURIComponent(tenant)}/${encodeURIComponent(calendar)}`;
    const watch = await this.client(provider).watch(t, calendar, callback);
    return { ...watch, id: this.watchHandle(t, provider, calendar, watch) };
  }
  private watchHandle(t: Tenant, provider: Provider, calendar: string, watch: Watch): string {
    const data = Buffer.from(JSON.stringify([t.id, provider, calendar, watch.id, watch.resourceId || null])).toString("base64url");
    return data + "." + createHmac("sha256", t.webhookSecret).update(data).digest("base64url");
  }
  private readWatch(t: Tenant, provider: Provider, calendar: string, handle: string): {id: string; resourceId?: string} {
    const [data, signature, extra] = handle.split(".");
    if (!data || !signature || extra || handle.length > 4096 || !secretEquals(signature, createHmac("sha256", t.webhookSecret).update(data).digest("base64url"))) fail("Invalid watch handle");
    let values;
    try { values = JSON.parse(Buffer.from(data, "base64url").toString("utf8")); }
    catch { fail("Invalid watch handle"); }
    if (!Array.isArray(values) || values.length !== 5 || values[0] !== t.id || values[1] !== provider || values[2] !== calendar || typeof values[3] !== "string") fail("Invalid watch handle for this tenant/calendar");
    return { id: values[3], resourceId: values[4] || undefined };
  }
  async renew(tenant: string, provider: Provider, calendar: string, handle: string): Promise<Watch> {
    const t = this.registry.calendar(tenant, provider, calendar);
    const watch = this.readWatch(t, provider, calendar, handle);
    if (provider !== "OUTLOOK") fail("Google channels cannot be renewed; create a new watch then stop the old watch");
    const renewed = await this.outlook.renew(t, watch.id);
    return { ...renewed, id: this.watchHandle(t, provider, calendar, renewed) };
  }
  async unwatch(
    tenant: string,
    provider: Provider,
    calendar: string,
    id: string,
    resourceId?: string,
  ): Promise<boolean> {
    const t = this.registry.calendar(tenant, provider, calendar);
    const watch = this.readWatch(t, provider, calendar, id);
    if (resourceId !== undefined && resourceId !== watch.resourceId) fail("Invalid watch resourceId");
    if (provider === "GOOGLE") {
      if (!watch.resourceId) fail("Google resourceId is required");
      await this.google.unwatch(t, watch.id, watch.resourceId);
    } else await this.outlook.unwatch(t, watch.id);
    return true;
  }
}
