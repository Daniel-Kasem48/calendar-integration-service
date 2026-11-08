const { test } = require("node:test");
const assert = require("node:assert/strict");
const { CalendarService } = require("../dist/calendar.service");
const { TenantRegistry } = require("../dist/config");
const { googleRecurrence, microsoftRecurrence } = require("../dist/recurrence");

const config = () => [
  {
    id: "acme",
    apiKey: "a".repeat(40),
    webhookSecret: "s".repeat(40),
    notificationUrl: "https://receiver.example/changes",
    google: {
      clientId: "g-client",
      clientSecret: "g-secret",
      refreshToken: "g-refresh",
      calendars: ["room@example.com"],
    },
    microsoft: {
      tenantId: "directory-id",
      clientId: "m-client",
      clientSecret: "m-secret",
      calendars: ["room@example.com"],
    },
  },
];
const input = {
  title: "Planning",
  start: "2026-12-01T09:00:00+03:00",
  end: "2026-12-01T10:00:00+03:00",
  timeZone: "Asia/Damascus",
  attendees: ["guest@example.com"],
};
function setup(responder) {
  const calls = [];
  const http = async (url, options = {}) => {
    calls.push({ url: String(url), ...options });
    if (String(url).includes("/token"))
      return Response.json({ access_token: "access", expires_in: 3600 });
    return responder(String(url), options);
  };
  const registry = new TenantRegistry(config());
  return { service: new CalendarService(registry, http), calls, registry };
}

test("API keys identify a single tenant and reject missing or incorrect credentials", () => {
  const registry = new TenantRegistry(config());
  assert.equal(registry.authenticate("Bearer " + "a".repeat(40)).id, "acme");
  for (const token of [undefined, "Bearer wrong", "Basic " + "a".repeat(40)])
    assert.throws(() => registry.authenticate(token));
});
test("duplicate API keys and weak secrets fail startup", () => {
  const c = config();
  assert.throws(() => new TenantRegistry([c[0], { ...c[0], id: "other" }]));
  assert.throws(() => new TenantRegistry([{ ...c[0], apiKey: "short" }]));
});
test("calendar allowlist prevents calls to other mailboxes", async () => {
  const { service, calls } = setup(() => {
    throw Error("must not call upstream");
  });
  await assert.rejects(
    service.list("acme", "GOOGLE", "other@example.com"),
    /calendar/i,
  );
  assert.equal(calls.length, 0);
});
test("Google event creation maps dates, attendees and private visibility without database entities", async () => {
  const { service, calls } = setup((url, options) => {
    const body = JSON.parse(options.body);
    assert.equal(body.summary, "Planning");
    assert.equal(body.start.dateTime, "2026-12-01T09:00:00+03:00");
    assert.equal(body.start.timeZone, "Asia/Damascus");
    assert.deepEqual(body.attendees, [{ email: "guest@example.com" }]);
    assert.equal(body.visibility, "private");
    return Response.json({ id: "event-1", ...body });
  });
  const event = await service.create(
    "acme",
    "GOOGLE",
    "room@example.com",
    input,
  );
  assert.equal(event.id, "event-1");
  assert.equal(event.title, "Planning");
  assert.match(calls[1].url, /calendars\/room%40example.com\/events/);
});
test("Microsoft creation preserves the event timezone and wall clock for recurring meetings", async () => {
  const { service } = setup((url, options) => {
    const body = JSON.parse(options.body);
    assert.deepEqual(body.start, {
      dateTime: "2026-12-01T09:00:00.000",
      timeZone: "Asia/Damascus",
    });
    assert.deepEqual(body.attendees, [
      { emailAddress: { address: "guest@example.com" }, type: "required" },
    ]);
    return Response.json({ id: "ms-1", ...body });
  });
  const event = await service.create(
    "acme",
    "MICROSOFT",
    "room@example.com",
    input,
  );
  assert.equal(event.start, "2026-12-01T06:00:00.000Z");
});
test("a title-only update does not clear attendees or access missing dates", async () => {
  const { service } = setup((url, options) => {
    if (!options.method)
      return Response.json({
        id: "e",
        summary: "Old",
        start: { dateTime: input.start },
        end: { dateTime: input.end },
      });
    assert.deepEqual(JSON.parse(options.body), { summary: "Updated" });
    return Response.json({
      id: "e",
      summary: "Updated",
      start: { dateTime: input.start },
      end: { dateTime: input.end },
    });
  });
  assert.equal(
    (
      await service.update("acme", "GOOGLE", "room@example.com", "e", {
        title: "Updated",
      })
    ).title,
    "Updated",
  );
});
test("an end-only update is applied and validated against the existing start", async () => {
  const { service } = setup((url, options) => {
    if (!options.method)
      return Response.json({
        id: "e",
        summary: "Old",
        start: { dateTime: input.start },
        end: { dateTime: input.end },
      });
    assert.deepEqual(JSON.parse(options.body), {
      end: { dateTime: "2026-12-01T11:00:00+03:00", timeZone: "Asia/Damascus" },
    });
    return Response.json({
      id: "e",
      summary: "Old",
      start: { dateTime: input.start },
      end: { dateTime: "2026-12-01T11:00:00+03:00" },
    });
  });
  await service.update("acme", "GOOGLE", "room@example.com", "e", {
    end: "2026-12-01T11:00:00+03:00",
    timeZone: "Asia/Damascus",
  });
  await assert.rejects(
    service.update("acme", "GOOGLE", "room@example.com", "e", {
      end: "2026-12-01T08:00:00+03:00",
    }),
    /end/i,
  );
});
test("invalid dates, email addresses, unknown timezones and inverted intervals are rejected before creation", async () => {
  const { service } = setup(() => {
    throw Error("must not call upstream");
  });
  for (const patch of [
    { start: "tomorrow" },
    { start: "2026-12-01T09:00:00" },
    { end: input.start },
    { timeZone: "fake/zone" },
    { attendees: ["invalid"] },
    { title: "" },
  ]) {
    await assert.rejects(
      service.create("acme", "GOOGLE", "room@example.com", {
        ...input,
        ...patch,
      }),
    );
  }
});
test("Google pagination preserves opaque cursors and exposes sync tokens only from upstream", async () => {
  const { service } = setup((url) => {
    assert.equal(new URL(url).searchParams.get("pageToken"), "page+/=");
    assert.equal(new URL(url).searchParams.get("syncToken"), "sync-1");
    return Response.json({
      items: [{ id: "gone", status: "cancelled" }],
      nextSyncToken: "sync-2",
    });
  });
  const page = await service.list("acme", "GOOGLE", "room@example.com", {
    pageToken: "page+/=",
    syncToken: "sync-1",
  });
  assert.equal(page.nextSyncToken, "sync-2");
  assert.equal(page.items[0].status, "cancelled");
});
test("Microsoft continuation URLs cannot target another host or mailbox", async () => {
  const { service } = setup(() =>
    Response.json({
      value: [],
      "@odata.nextLink": "https://evil.example/steal",
    }),
  );
  await assert.rejects(
    service.list("acme", "MICROSOFT", "room@example.com"),
    /continuation/i,
  );
  await assert.rejects(
    service.list("acme", "MICROSOFT", "room@example.com", {
      pageToken: "https://evil.example",
    }),
    /cursor/i,
  );
});
test("OAuth caches stay scoped to the tenant", async () => {
  const c = config();
  c.push({
    ...c[0],
    id: "second",
    apiKey: "b".repeat(40),
    google: { ...c[0].google, refreshToken: "second-refresh" },
  });
  const tokens = [];
  const http = async (url, options) => {
    if (String(url).includes("/token")) {
      const refresh = new URLSearchParams(options.body).get("refresh_token");
      tokens.push(refresh);
      return Response.json({ access_token: refresh, expires_in: 3600 });
    }
    return Response.json({ items: [] });
  };
  const s = new CalendarService(new TenantRegistry(c), http);
  await Promise.all(
    ["acme", "second"].map((t) => s.list(t, "GOOGLE", "room@example.com")),
  );
  assert.deepEqual(tokens.sort(), ["g-refresh", "second-refresh"]);
});
test("provider failures are not swallowed and do not reveal response secrets", async () => {
  for (const [status, code] of [
    [410, "SYNC_TOKEN_EXPIRED"],
    [429, "RATE_LIMITED"],
    [401, "REAUTH_REQUIRED"],
    [500, "UPSTREAM_ERROR"],
  ]) {
    const { service } = setup(() =>
      Response.json({ error: "SECRET-DETAIL" }, { status }),
    );
    await assert.rejects(
      service.list("acme", "GOOGLE", "room@example.com"),
      (error) =>
        error.extensions.code === code && !error.message.includes("SECRET"),
    );
  }
});
test("delete failures propagate rather than falsely claiming cancellation", async () => {
  const { service } = setup(() => new Response(null, { status: 403 }));
  await assert.rejects(
    service.remove("acme", "GOOGLE", "room@example.com", "event"),
  );
});

const recurrence = {
  recurrenceType: "WEEK",
  interval: 2,
  rangeType: "OCCURRENCE_COUNT",
  occurrenceCount: 5,
  daysOfWeek: [1, 3],
};
test("weekly recurrence is mapped to RRULE and Microsoft pattern/range", () => {
  assert.equal(
    googleRecurrence(recurrence, input.start, input.timeZone),
    "RRULE:FREQ=WEEKLY;BYDAY=MO,WE;INTERVAL=2;COUNT=5",
  );
  assert.deepEqual(
    microsoftRecurrence(recurrence, input.start, input.timeZone),
    {
      pattern: {
        type: "weekly",
        interval: 2,
        daysOfWeek: ["monday", "wednesday"],
      firstDayOfWeek: "monday",
      },
      range: {
        type: "numbered",
        startDate: "2026-12-01",
        recurrenceTimeZone: "Asia/Damascus",
        numberOfOccurrences: 5,
      },
    },
  );
});
test("all six original recurrence kinds keep their month and weekday conventions", () => {
  const base = { interval: 1, rangeType: "NEVER" };
  const cases = [
    [{ recurrenceType: "DAY" }, "RRULE:FREQ=DAILY", "daily"],
    [
      { recurrenceType: "WEEK", daysOfWeek: [0] },
      "RRULE:FREQ=WEEKLY;BYDAY=SU",
      "weekly",
    ],
    [
      { recurrenceType: "AbsoluteMonthly", dayOfMonth: 3 },
      "RRULE:FREQ=MONTHLY;BYMONTHDAY=3",
      "absoluteMonthly",
    ],
    [
      { recurrenceType: "RelativeMonthly", weekDay: 0, weekIndex: 1 },
      "RRULE:FREQ=MONTHLY;BYDAY=2SU",
      "relativeMonthly",
    ],
    [
      { recurrenceType: "AbsoluteYearly", month: 3, dayOfMonth: 5 },
      "RRULE:FREQ=YEARLY;BYMONTH=4;BYMONTHDAY=5",
      "absoluteYearly",
    ],
    [
      { recurrenceType: "RelativeYearly", month: 3, weekDay: 0, weekIndex: 1 },
      "RRULE:FREQ=YEARLY;BYMONTH=4;BYDAY=2SU",
      "relativeYearly",
    ],
  ];
  for (const [fields, rule, kind] of cases) {
    const r = { ...base, ...fields };
    assert.equal(googleRecurrence(r, input.start, input.timeZone), rule);
    assert.equal(
      microsoftRecurrence(r, input.start, input.timeZone).pattern.type,
      kind,
    );
  }
});
test("invalid recurrence fields are rejected instead of producing malformed provider data", () => {
  for (const r of [
    { ...recurrence, daysOfWeek: [] },
    { ...recurrence, daysOfWeek: [7] },
    { ...recurrence, interval: 0 },
    { ...recurrence, occurrenceCount: 0 },
    { ...recurrence, recurrenceType: "unknown" },
  ]) {
    assert.throws(() => googleRecurrence(r, input.start, input.timeZone));
  }
});
test("Google UNTIL uses a UTC datetime and Microsoft uses a local inclusive end date", () => {
  const r = {
    recurrenceType: "DAY",
    interval: 1,
    rangeType: "UNTIL_DATE",
    untilDate: "2026-12-31",
  };
  assert.equal(
    googleRecurrence(r, input.start, input.timeZone),
    "RRULE:FREQ=DAILY;UNTIL=20261231T205959Z",
  );
  assert.equal(
    microsoftRecurrence(r, input.start, input.timeZone).range.endDate,
    "2026-12-31",
  );
});

test('Google watches include a verification token and actual expiration; stopping uses both channel identifiers', async () => {
  let watchRawId;
  const { service } = setup((url, options) => {
    const body = JSON.parse(options.body);
    if (url.endsWith('/watch')) {
      assert.equal(body.address, 'https://calendar.example/webhooks/google/acme/room%40example.com');
      assert.equal(body.token.length, 43);
      watchRawId = body.id;
      return Response.json({ id: body.id, resourceId: 'resource', expiration: '1796126400000' });
    }
    assert.deepEqual(body, { id: watchRawId, resourceId: 'resource' });
    return new Response(null, { status: 204 });
  });
  const watch = await service.watch('acme', 'GOOGLE', 'room@example.com', 'https://calendar.example');
  assert.equal(watch.resourceId, 'resource');
  assert.equal(watch.expiresAt, '2026-12-01T12:00:00.000Z');
  assert.equal(await service.unwatch('acme', 'GOOGLE', 'room@example.com', watch.id, 'resource'), true);
});
test('Microsoft watches set clientState and return provider expiration', async () => {
  const { service } = setup((url, options) => {
    const body = JSON.parse(options.body);
    assert.equal(body.resource, '/users/room@example.com/events');
    assert.equal(body.changeType, 'created,updated,deleted');
    assert.equal(body.clientState.length, 43);
    assert.equal(body.notificationUrl, 'https://calendar.example/webhooks/microsoft/acme/room%40example.com');
    return Response.json({ id: 'subscription', expirationDateTime: '2026-12-01T12:00:00Z' });
  });
  const result = await service.watch('acme', 'MICROSOFT', 'room@example.com', 'https://calendar.example');
  assert.ok(result.id);
  assert.equal(result.expiresAt, '2026-12-01T12:00:00Z');
});
test('Microsoft next-page tokens round-trip a legitimate encoded calendar URL', async () => {
  let pages = 0;
  const { service } = setup((url) => {
    if (pages++ === 0) return Response.json({ value: [], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users/room%40example.com/calendar/events?$skiptoken=next' });
    assert.equal(new URL(url).searchParams.get('$skiptoken'), 'next');
    return Response.json({ value: [] });
  });
  const first = await service.list('acme', 'MICROSOFT', 'room@example.com');
  const second = await service.list('acme', 'MICROSOFT', 'room@example.com', { pageToken: first.nextPageToken });
  assert.deepEqual(second.items, []);
});
test('Microsoft continuation accepts equivalent unescaped email addresses', async () => {
  const { service } = setup(() => Response.json({ value: [], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users/room@example.com/calendar/events?$skiptoken=next' }));
  assert.equal(typeof (await service.list('acme', 'MICROSOFT', 'room@example.com')).nextPageToken, 'string');
});

test('Microsoft recurrence keeps the Berlin wall clock across daylight-saving transitions', async () => {
  const { service } = setup((url, options) => {
    const body = JSON.parse(options.body);
    assert.deepEqual(body.start, { dateTime: '2027-03-22T09:00:00.000', timeZone: 'Europe/Berlin' });
    assert.equal(body.recurrence.pattern.firstDayOfWeek, 'monday');
    return Response.json({ id: 'series', ...body });
  });
  await service.create('acme', 'MICROSOFT', 'room@example.com', { ...input, start: '2027-03-22T09:00:00+01:00', end: '2027-03-22T10:00:00+01:00', timeZone: 'Europe/Berlin', recurrence });
});
test('watch cancellation rejects a handle from another tenant or calendar without calling upstream', async () => {
  const c = config();
  c[0].google.calendars.push('second@example.com');
  c.push({ ...c[0], id: 'second', apiKey: 'b'.repeat(40) });
  let calls = 0;
  const s = new CalendarService(new TenantRegistry(c), async (url, options) => {
    calls++;
    if (String(url).includes('/token')) return Response.json({ access_token: 'token' });
    const body = JSON.parse(options.body);
    return Response.json({ id: body.id, resourceId: 'resource', expiration: '1796126400000' });
  });
  const watch = await s.watch('acme', 'GOOGLE', 'room@example.com', 'https://calendar.example');
  const before = calls;
  await assert.rejects(s.unwatch('second', 'GOOGLE', 'room@example.com', watch.id, 'resource'), /watch/i);
  await assert.rejects(s.unwatch('acme', 'GOOGLE', 'second@example.com', watch.id, 'resource'), /watch/i);
  assert.equal(calls, before);
});
test('Microsoft renewal patches the existing subscription rather than creating a duplicate', async () => {
  const { service } = setup((url, options) => {
    if (options.method === 'POST') return Response.json({ id: 'subscription', expirationDateTime: '2026-12-01T12:00:00Z' });
    assert.equal(options.method, 'PATCH');
    assert.equal(url, 'https://graph.microsoft.com/v1.0/subscriptions/subscription');
    assert.ok(JSON.parse(options.body).expirationDateTime);
    return Response.json({ id: 'subscription', expirationDateTime: '2026-12-03T12:00:00Z' });
  });
  const watch = await service.watch('acme', 'MICROSOFT', 'room@example.com', 'https://calendar.example');
  const renewed = await service.renew('acme', 'MICROSOFT', 'room@example.com', watch.id);
  assert.equal(renewed.id, watch.id);
  assert.equal(renewed.expiresAt, '2026-12-03T12:00:00Z');
});

test('end-only updates work for untitled provider events', async () => {
  const { service } = setup((url, options) => {
    if (!options.method) return Response.json({ id: 'untitled', summary: '', start: { dateTime: input.start, timeZone: input.timeZone }, end: { dateTime: input.end } });
    return Response.json({ id: 'untitled', summary: '', start: { dateTime: input.start }, end: { dateTime: '2026-12-01T11:00:00+03:00' } });
  });
  assert.equal((await service.update('acme', 'GOOGLE', 'room@example.com', 'untitled', { end: '2026-12-01T11:00:00+03:00' })).title, '');
});
test('Microsoft partial updates preserve a series original IANA timezone', async () => {
  const { service } = setup((url, options) => {
    if (!options.method) return Response.json({ id: 'e', subject: 'Series', originalStartTimeZone: 'Europe/Berlin', start: { dateTime: '2027-03-22T08:00:00', timeZone: 'UTC' }, end: { dateTime: '2027-03-22T09:00:00', timeZone: 'UTC' } });
    const body = JSON.parse(options.body);
    assert.deepEqual(body.end, { dateTime: '2027-03-22T10:30:00.000', timeZone: 'Europe/Berlin' });
    return Response.json({ id: 'e', subject: 'Series', start: { dateTime: '2027-03-22T09:00:00', timeZone: 'Europe/Berlin' }, end: body.end });
  });
  await service.update('acme', 'MICROSOFT', 'room@example.com', 'e', { end: '2027-03-22T10:30:00+01:00' });
});
