import "reflect-metadata";
import { Module } from "@nestjs/common";
import { NestFactory } from "@nestjs/core";
import { GraphQLModule } from "@nestjs/graphql";
import { ApolloDriver, ApolloDriverConfig } from "@nestjs/apollo";
import { GraphQLError, ValidationRule } from "graphql";
import { CalendarService } from "./calendar.service";
import { TenantRegistry, httpsUrl } from "./config";
import { ApiKeyGuard, CalendarResolver } from "./graphql";
import { WebhookController } from "./webhooks";
import { Fetch } from "./types";

// shortcut: cap operation size; add a shared rate limiter before exposing a high-traffic public endpoint.
const operationLimit: ValidationRule = (context) => {
  let fields = 0;
  return {
    Field() {
      if (++fields === 101)
        context.reportError(new GraphQLError("Operation exceeds 100 fields"));
    },
  };
};
export async function createApp(
  registry = TenantRegistry.fromEnvironment(),
  http: Fetch = fetch,
  publicUrl = process.env.PUBLIC_URL || "",
) {
  if (publicUrl) httpsUrl(publicUrl);
  @Module({
    imports: [
      GraphQLModule.forRoot<ApolloDriverConfig>({
        driver: ApolloDriver,
        autoSchemaFile: true,
        sortSchema: true,
        playground: false,
        introspection: process.env.NODE_ENV !== "production",
        validationRules: [operationLimit],
        includeStacktraceInErrorResponses: false,
        context: ({ req }: { req: unknown }) => ({ req }),
      }),
    ],
    controllers: [WebhookController],
    providers: [
      CalendarResolver,
      ApiKeyGuard,
      { provide: TenantRegistry, useValue: registry },
      {
        provide: CalendarService,
        useValue: new CalendarService(registry, http),
      },
      { provide: "FETCH", useValue: http },
      { provide: "PUBLIC_URL", useValue: publicUrl },
    ],
  })
  class AppModule {}
  return NestFactory.create(AppModule, { logger: ["error", "warn"] });
}
