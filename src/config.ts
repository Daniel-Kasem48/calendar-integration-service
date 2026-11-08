import { createHash, timingSafeEqual } from "node:crypto";
import { readFileSync } from "node:fs";
import { GraphQLError } from "graphql";
import { Provider } from "./types";

export interface GoogleConfig {
  clientId: string;
  clientSecret: string;
  refreshToken: string;
  calendars: string[];
}
export interface MicrosoftConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  calendars: string[];
}
export interface Tenant {
  id: string;
  apiKey: string;
  webhookSecret: string;
  notificationUrl?: string;
  google?: GoogleConfig;
  microsoft?: MicrosoftConfig;
}
export function fail(message: string, code = "BAD_USER_INPUT"): never {
  throw new GraphQLError(message, { extensions: { code } });
}
export function secretEquals(a: string, b: string): boolean {
  return timingSafeEqual(
    createHash("sha256").update(a).digest(),
    createHash("sha256").update(b).digest(),
  );
}
export function httpsUrl(value: string): string {
  const url = new URL(value);
  if (url.protocol !== "https:" || url.username || url.password || url.hash)
    throw Error("Expected an HTTPS URL without credentials or fragment");
  return url.href;
}
export class TenantRegistry {
  constructor(private readonly tenants: Tenant[]) {
    if (!Array.isArray(tenants) || tenants.length === 0)
      throw Error("At least one tenant is required");
    const ids = new Set<string>();
    const keys = new Set<string>();
    for (const t of tenants) {
      if (
        !t ||
        typeof t.id !== "string" ||
        !/^[a-zA-Z0-9_-]{1,64}$/.test(t.id) ||
        ids.has(t.id)
      )
        throw Error("Tenant IDs must be unique URL-safe strings");
      if (
        typeof t.apiKey !== "string" ||
        t.apiKey.length < 32 ||
        keys.has(t.apiKey)
      )
        throw Error(
          "Tenant API keys must be unique and at least 32 characters",
        );
      if (typeof t.webhookSecret !== "string" || t.webhookSecret.length < 32)
        throw Error("Webhook secrets must be at least 32 characters");
      if (!t.google && !t.microsoft)
        throw Error("A calendar provider must be configured");
      if (t.notificationUrl) httpsUrl(t.notificationUrl);
      for (const provider of [t.google, t.microsoft]) {
        if (!provider) continue;
        if (
          !Array.isArray(provider.calendars) ||
          !provider.calendars.length ||
          provider.calendars.some((c) => typeof c !== "string" || !c.trim())
        )
          throw Error("Explicit calendar allowlists are required");
        const fields =
          "refreshToken" in provider
            ? ["clientId", "clientSecret", "refreshToken"]
            : ["tenantId", "clientId", "clientSecret"];
        for (const field of fields)
          if (
            typeof (provider as unknown as Record<string, unknown>)[field] !==
              "string" ||
            !(provider as unknown as Record<string, string>)[field].trim()
          )
            throw Error(`Missing provider credential: ${field}`);
      }
      ids.add(t.id);
      keys.add(t.apiKey);
    }
  }
  static fromEnvironment(): TenantRegistry {
    return new TenantRegistry(
      JSON.parse(
        readFileSync(process.env.TENANTS_FILE || "tenants.json", "utf8"),
      ),
    );
  }
  authenticate(header?: string): Tenant {
    const key = header?.startsWith("Bearer ") ? header.slice(7) : "";
    const tenant = this.tenants.find((t) => secretEquals(t.apiKey, key));
    return (
      tenant ||
      fail("A valid tenant Bearer token is required", "UNAUTHENTICATED")
    );
  }
  get(id: string): Tenant {
    return (
      this.tenants.find((t) => t.id === id) ||
      fail("Unknown tenant", "UNAUTHENTICATED")
    );
  }
  calendar(id: string, provider: Provider, calendarId: string): Tenant {
    const tenant = this.get(id);
    if (provider !== "GOOGLE" && provider !== "MICROSOFT")
      fail("Unknown calendar provider");
    const settings = provider === "GOOGLE" ? tenant.google : tenant.microsoft;
    if (!settings?.calendars.includes(calendarId))
      fail("Calendar is not allowed for this tenant", "FORBIDDEN");
    return tenant;
  }
}
