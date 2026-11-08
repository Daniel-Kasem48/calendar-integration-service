const { test } = require("node:test");
const assert = require("node:assert/strict");
const { createApp } = require("../dist/app");
const { TenantRegistry } = require("../dist/config");
const { webhookToken } = require("../dist/providers/google");
const key = "a".repeat(40);
const tenant = {
  id: "acme",
  apiKey: key,
  webhookSecret: "s".repeat(40),
  notificationUrl: "https://receiver.example/changes",
  google: {
    clientId: "id",
    clientSecret: "secret",
    refreshToken: "refresh",
    calendars: ["room@example.com"],
  },
  outlook: {
    tenantId: "tenant",
    clientId: "id",
    clientSecret: "secret",
    calendars: ["room@example.com"],
  },
};
async function fixture(t, responder = () => Response.json({ items: [] })) {
  const app = await createApp(
    new TenantRegistry([tenant]),
    responder,
    "https://calendar.example",
  );
  await app.listen(0, "127.0.0.1");
  t.after(() => app.close());
  const base = await app.getUrl();
  return {
    base,
    gql: async (query, variables = {}, auth = key) => {
      const response = await fetch(base + "/graphql", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          ...(auth ? { authorization: "Bearer " + auth } : {}),
        },
        body: JSON.stringify({ query, variables }),
      });
      return { response, body: await response.json() };
    },
  };
}
test("real GraphQL HTTP requests require a valid tenant token", async (t) => {
  const { gql } = await fixture(t);
  const result = await gql(
    '{ events(provider: GOOGLE, calendarId: "room@example.com") { items { id } } }',
    {},
    "wrong",
  );
  assert.equal(result.body.errors[0].extensions.code, "UNAUTHENTICATED");
});
test("GraphQL queries return normalized provider results", async (t) => {
  const { gql } = await fixture(t, (url) =>
    String(url).includes("/token")
      ? Response.json({ access_token: "token" })
      : Response.json({
          items: [
            {
              id: "e1",
              summary: "Planning",
              start: { dateTime: "2026-12-01T09:00:00Z" },
              end: { dateTime: "2026-12-01T10:00:00Z" },
            },
          ],
        }),
  );
  const result = await gql(
    '{ events(provider: GOOGLE, calendarId: "room@example.com") { items { id title } } }',
  );
  assert.deepEqual(result.body.data.events.items, [
    { id: "e1", title: "Planning" },
  ]);
});
test("GraphQL mutation creates an event through the real service and provider mapping", async (t) => {
  const { gql } = await fixture(t, (url, options) => {
    if (String(url).includes("/token"))
      return Response.json({ access_token: "token" });
    const body = JSON.parse(options.body);
    return Response.json({ id: "created", ...body });
  });
  const result = await gql(
    'mutation($input: CreateEventInput!) { createEvent(provider: GOOGLE, calendarId: "room@example.com", input: $input) { id title } }',
    {
      input: {
        title: "Planning",
        start: "2026-12-01T09:00:00Z",
        end: "2026-12-01T10:00:00Z",
        timeZone: "UTC",
      },
    },
  );
  assert.deepEqual(result.body.data?.createEvent, {
    id: "created",
    title: "Planning",
  });
});
test("GraphQL null patch fields are rejected before invoking the provider", async (t) => {
  const { gql } = await fixture(t, () => {
    throw Error("must not call");
  });
  const result = await gql(
    'mutation { updateEvent(provider: GOOGLE, calendarId: "room@example.com", eventId: "e", input: {title:null}) { id } }',
  );
  assert.equal(result.body.errors[0].extensions.code, "BAD_USER_INPUT");
});
test("Outlook webhook validation echoes the decoded token as plain text", async (t) => {
  const { base } = await fixture(t);
  const response = await fetch(
    base +
      "/webhooks/outlook/acme/room%40example.com?validationToken=hello%20world",
    { method: "POST" },
  );
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type"), /text\/plain/);
  assert.equal(await response.text(), "hello world");
});
test("spoofed webhook notifications are rejected before forwarding", async (t) => {
  const { base } = await fixture(t, () => {
    throw Error("must not forward");
  });
  const response = await fetch(
    base + "/webhooks/google/acme/room%40example.com",
    { method: "POST", headers: { "x-goog-channel-token": "invalid" } },
  );
  assert.equal(response.status, 401);
});
test("valid Google webhook forwards tenant-scoped metadata with a signed body", async (t) => {
  let delivery;
  const { base } = await fixture(t, async (url, options) => {
    delivery = { url: String(url), ...options };
    return new Response(null, { status: 204 });
  });
  const response = await fetch(
    base + "/webhooks/google/acme/room%40example.com",
    {
      method: "POST",
      headers: {
        "x-goog-channel-token": webhookToken(
          tenant,
          "GOOGLE",
          "room@example.com",
        ),
        "x-goog-channel-id": "channel-1",
        "x-goog-resource-state": "exists",
        "x-goog-message-number": "4",
      },
    },
  );
  assert.equal(response.status, 204);
  assert.equal(delivery.url, "https://receiver.example/changes");
  assert.equal(JSON.parse(delivery.body).tenantId, "acme");
  assert.equal(JSON.parse(delivery.body).calendarId, "room@example.com");
  assert.match(
    delivery.headers["x-calendar-signature"],
    /^sha256=[0-9a-f]{64}$/,
  );
});
test("notification delivery failure returns 503 so the provider can retry", async (t) => {
  const { base } = await fixture(
    t,
    async () => new Response(null, { status: 500 }),
  );
  const response = await fetch(
    base + "/webhooks/outlook/acme/room%40example.com",
    {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        value: [
          {
            clientState: webhookToken(tenant, "OUTLOOK", "room@example.com"),
            subscriptionId: "sub",
            resource: "users/room@example.com/events/1",
            changeType: "updated",
          },
        ],
      }),
    },
  );
  assert.equal(response.status, 503);
});
test("health endpoint reveals no tenant credentials", async (t) => {
  const { base } = await fixture(t);
  const response = await fetch(base + "/health");
  assert.deepEqual(await response.json(), { status: "ok" });
});

test('OUTLOOK is the GraphQL provider name for Outlook calendar queries', async t => {
  const { gql } = await fixture(t, url => String(url).includes('/token')
    ? Response.json({ access_token: 'token' })
    : Response.json({ value: [{ id: 'outlook-event', subject: 'Planning' }] }));
  const result = await gql('{ events(provider: OUTLOOK, calendarId: "room@example.com") { items { id title } } }');
  assert.deepEqual(result.body.data?.events.items, [{ id: 'outlook-event', title: 'Planning' }]);
});
