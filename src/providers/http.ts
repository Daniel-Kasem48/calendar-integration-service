import { GraphQLError } from "graphql";
import { Tenant } from "../config";
import { Fetch, Provider } from "../types";

export class ProviderHttp {
  private readonly tokens = new Map<
    string,
    { value: string; expires: number }
  >();
  constructor(private readonly http: Fetch = fetch) {}
  private async send(url: string, init: RequestInit): Promise<any> {
    let response: Response;
    try {
      response = await this.http(url, {
        ...init,
        redirect: "error",
        signal: AbortSignal.timeout(15_000),
      });
    } catch {
      throw new GraphQLError("Calendar provider could not be reached", {
        extensions: { code: "UPSTREAM_UNAVAILABLE" },
      });
    }
    if (!response.ok) {
      const code =
        (
          {
            401: "REAUTH_REQUIRED",
            403: "PROVIDER_FORBIDDEN",
            404: "NOT_FOUND",
            410: "SYNC_TOKEN_EXPIRED",
            429: "RATE_LIMITED",
          } as Record<number, string>
        )[response.status] || "UPSTREAM_ERROR";
      throw new GraphQLError(
        `Calendar provider request failed (HTTP ${response.status})`,
        {
          extensions: {
            code,
            status: response.status,
            retryAfter: response.headers.get("retry-after"),
          },
        },
      );
    }
    if (
      response.status === 204 ||
      response.headers.get("content-length") === "0"
    )
      return undefined;
    const text = await response.text();
    if (!text) return undefined;
    try {
      return JSON.parse(text);
    } catch {
      throw new GraphQLError("Calendar provider returned invalid JSON", {
        extensions: { code: "UPSTREAM_ERROR" },
      });
    }
  }
  private async token(tenant: Tenant, provider: Provider): Promise<string> {
    const key = `${tenant.id}:${provider}`;
    const cached = this.tokens.get(key);
    if (cached && cached.expires > Date.now() + 60_000) return cached.value;
    let url: string;
    let body: URLSearchParams;
    if (provider === "GOOGLE") {
      const c = tenant.google!;
      url = "https://oauth2.googleapis.com/token";
      body = new URLSearchParams({
        grant_type: "refresh_token",
        client_id: c.clientId,
        client_secret: c.clientSecret,
        refresh_token: c.refreshToken,
      });
    } else {
      const c = tenant.outlook!;
      url = `https://login.microsoftonline.com/${encodeURIComponent(c.tenantId)}/oauth2/v2.0/token`;
      body = new URLSearchParams({
        grant_type: "client_credentials",
        client_id: c.clientId,
        client_secret: c.clientSecret,
        scope: "https://graph.microsoft.com/.default",
      });
    }
    const token = await this.send(url, {
      method: "POST",
      body,
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
    });
    if (typeof token?.access_token !== "string" || !token.access_token)
      throw new GraphQLError("Provider did not return an access token", {
        extensions: { code: "REAUTH_REQUIRED" },
      });
    this.tokens.set(key, {
      value: token.access_token,
      expires: Date.now() + (Number(token.expires_in) || 3600) * 1000,
    });
    return token.access_token;
  }
  async request(
    tenant: Tenant,
    provider: Provider,
    url: string,
    method?: string,
    body?: unknown,
  ): Promise<any> {
    const token = await this.token(tenant, provider);
    try {
      return await this.send(url, {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          ...(provider === "OUTLOOK"
            ? { Prefer: 'outlook.timezone="UTC"' }
            : {}),
        },
        body: body === undefined ? undefined : JSON.stringify(body),
      });
    } catch (error) {
      if (
        error instanceof GraphQLError &&
        error.extensions.code === "REAUTH_REQUIRED"
      )
        this.tokens.delete(`${tenant.id}:${provider}`);
      throw error;
    }
  }
}
