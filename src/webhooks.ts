import {
  Body,
  Controller,
  Get,
  Headers,
  Inject,
  Param,
  Post,
  Query,
  Res,
  ServiceUnavailableException,
  UnauthorizedException,
  BadRequestException,
} from "@nestjs/common";
import { createHmac } from "node:crypto";
import { Response } from "express";
import { TenantRegistry, secretEquals } from "./config";
import { webhookToken } from "./providers/google";
import { Fetch, Provider } from "./types";

@Controller()
export class WebhookController {
  constructor(
    private readonly registry: TenantRegistry,
    @Inject("FETCH") private readonly http: Fetch,
  ) {}
  @Get("health") health() {
    return { status: "ok" };
  }
  @Post("webhooks/:provider/:tenant/:calendar")
  async receive(
    @Param("provider") name: string,
    @Param("tenant") tenantId: string,
    @Param("calendar") calendarId: string,
    @Query("validationToken") validationToken: string | undefined,
    @Headers() headers: Record<string, string>,
    @Body() body: any,
    @Res() response: Response,
  ) {
    if (!["google", "microsoft"].includes(name))
      throw new BadRequestException("Unknown provider");
    const provider = name.toUpperCase() as Provider;
    let tenant;
    try {
      tenant = this.registry.calendar(tenantId, provider, calendarId);
    } catch {
      throw new UnauthorizedException();
    }
    if (name === "microsoft" && typeof validationToken === "string")
      return response.status(200).type("text/plain").send(validationToken);
    const expected = webhookToken(tenant, provider, calendarId);
    let notification: Record<string, unknown>;
    if (name === "google") {
      if (!secretEquals(headers["x-goog-channel-token"] || "", expected))
        throw new UnauthorizedException();
      if (!headers["x-goog-channel-id"] || !headers["x-goog-resource-state"])
        throw new BadRequestException("Missing channel metadata");
      if (headers["x-goog-resource-state"] === "sync")
        return response.status(204).send();
      notification = {
        channelId: headers["x-goog-channel-id"],
        resourceId: headers["x-goog-resource-id"],
        state: headers["x-goog-resource-state"],
        messageNumber: headers["x-goog-message-number"],
      };
    } else {
      if (
        !Array.isArray(body?.value) ||
        body.value.length === 0 ||
        body.value.length > 100
      )
        throw new BadRequestException("Invalid notifications");
      if (
        body.value.some(
          (item: any) =>
            typeof item?.clientState !== "string" ||
            !secretEquals(item.clientState, expected),
        )
      )
        throw new UnauthorizedException();
      notification = {
        value: body.value.map((item: any) => ({
          subscriptionId: item.subscriptionId,
          changeType: item.changeType,
          resource: item.resource,
          resourceData: item.resourceData,
        })),
      };
    }
    if (!tenant.notificationUrl)
      throw new ServiceUnavailableException(
        "Notification receiver is not configured",
      );
    const payload = JSON.stringify({
      tenantId,
      provider,
      calendarId,
      ...notification,
    });
    const signature = createHmac("sha256", tenant.webhookSecret)
      .update(payload)
      .digest("hex");
    try {
      const delivery = await this.http(tenant.notificationUrl, {
        method: "POST",
        redirect: "error",
        signal: AbortSignal.timeout(5000),
        headers: {
          "content-type": "application/json",
          "x-calendar-signature": `sha256=${signature}`,
        },
        body: payload,
      });
      if (!delivery.ok) throw Error("Receiver failed");
    } catch {
      throw new ServiceUnavailableException(
        "Notification receiver is unavailable",
      );
    }
    return response.status(204).send();
  }
}
