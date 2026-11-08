import {
  CanActivate,
  ExecutionContext,
  Inject,
  Injectable,
  UseGuards,
} from "@nestjs/common";
import {
  Args,
  Context,
  Field,
  GqlExecutionContext,
  InputType,
  Int,
  Mutation,
  ObjectType,
  Query,
  Resolver,
  registerEnumType,
} from "@nestjs/graphql";
import { TenantRegistry, fail } from "./config";
import { CalendarService } from "./calendar.service";
import { EventInput, Provider, Recurrence } from "./types";

enum CalendarProvider {
  GOOGLE = "GOOGLE",
  MICROSOFT = "MICROSOFT",
}
enum RecurrenceType {
  DAY = "DAY",
  WEEK = "WEEK",
  AbsoluteMonthly = "AbsoluteMonthly",
  RelativeMonthly = "RelativeMonthly",
  AbsoluteYearly = "AbsoluteYearly",
  RelativeYearly = "RelativeYearly",
}
enum RangeType {
  NEVER = "NEVER",
  UNTIL_DATE = "UNTIL_DATE",
  OCCURRENCE_COUNT = "OCCURRENCE_COUNT",
}
registerEnumType(CalendarProvider, { name: "CalendarProvider" });
registerEnumType(RecurrenceType, { name: "RecurrenceType" });
registerEnumType(RangeType, { name: "RangeType" });

@InputType()
class RecurrenceInput implements Recurrence {
  @Field(() => RecurrenceType) recurrenceType!: RecurrenceType;
  @Field(() => RangeType) rangeType!: RangeType;
  @Field(() => Int) interval!: number;
  @Field(() => [Int], { nullable: true }) daysOfWeek?: number[];
  @Field(() => Int, { nullable: true }) dayOfMonth?: number;
  @Field(() => Int, { nullable: true }) weekDay?: number;
  @Field(() => Int, { nullable: true }) weekIndex?: number;
  @Field(() => Int, { nullable: true }) month?: number;
  @Field({ nullable: true }) untilDate?: string;
  @Field(() => Int, { nullable: true }) occurrenceCount?: number;
}
@InputType()
class CreateEventInput implements EventInput {
  @Field() title!: string;
  @Field({ nullable: true }) description?: string;
  @Field() start!: string;
  @Field() end!: string;
  @Field() timeZone!: string;
  @Field(() => [String], { nullable: true }) attendees?: string[];
  @Field(() => RecurrenceInput, { nullable: true })
  recurrence?: RecurrenceInput;
}
@InputType()
class UpdateEventInput implements Partial<EventInput> {
  @Field({ nullable: true }) title?: string;
  @Field({ nullable: true }) description?: string;
  @Field({ nullable: true }) start?: string;
  @Field({ nullable: true }) end?: string;
  @Field({ nullable: true }) timeZone?: string;
  @Field(() => [String], { nullable: true }) attendees?: string[];
  @Field(() => RecurrenceInput, { nullable: true })
  recurrence?: RecurrenceInput;
}
@ObjectType("CalendarEvent")
class EventObject {
  @Field() id!: string;
  @Field() title!: string;
  @Field({ nullable: true }) description?: string;
  @Field({ nullable: true }) start?: string;
  @Field({ nullable: true }) end?: string;
  @Field({ nullable: true }) timeZone?: string;
  @Field(() => [String]) attendees!: string[];
  @Field() status!: string;
  @Field() allDay!: boolean;
}
@ObjectType("EventPage")
class PageObject {
  @Field(() => [EventObject]) items!: EventObject[];
  @Field({ nullable: true }) nextPageToken?: string;
  @Field({ nullable: true }) nextSyncToken?: string;
}
@ObjectType("CalendarWatch")
class WatchObject {
  @Field() id!: string;
  @Field({ nullable: true }) resourceId?: string;
  @Field() expiresAt!: string;
}
interface RequestContext {
  req: { headers: { authorization?: string }; tenantId?: string };
}
@Injectable()
export class ApiKeyGuard implements CanActivate {
  constructor(private readonly registry: TenantRegistry) {}
  canActivate(context: ExecutionContext): boolean {
    const request =
      GqlExecutionContext.create(context).getContext<RequestContext>().req;
    request.tenantId = this.registry.authenticate(
      request.headers.authorization,
    ).id;
    return true;
  }
}
@Resolver()
@UseGuards(ApiKeyGuard)
export class CalendarResolver {
  constructor(
    private readonly service: CalendarService,
    @Inject("PUBLIC_URL") private readonly publicUrl: string,
  ) {}
  @Query(() => PageObject)
  events(
    @Context() context: RequestContext,
    @Args("provider", { type: () => CalendarProvider }) provider: Provider,
    @Args("calendarId") calendarId: string,
    @Args("pageToken", { nullable: true }) pageToken?: string,
    @Args("syncToken", { nullable: true }) syncToken?: string,
  ) {
    return this.service.list(context.req.tenantId!, provider, calendarId, {
      pageToken,
      syncToken,
    });
  }
  @Query(() => EventObject)
  event(
    @Context() context: RequestContext,
    @Args("provider", { type: () => CalendarProvider }) provider: Provider,
    @Args("calendarId") calendarId: string,
    @Args("eventId") eventId: string,
  ) {
    return this.service.get(
      context.req.tenantId!,
      provider,
      calendarId,
      eventId,
    );
  }
  @Mutation(() => EventObject)
  createEvent(
    @Context() context: RequestContext,
    @Args("provider", { type: () => CalendarProvider }) provider: Provider,
    @Args("calendarId") calendarId: string,
    @Args("input") input: CreateEventInput,
  ) {
    return this.service.create(
      context.req.tenantId!,
      provider,
      calendarId,
      input,
    );
  }
  @Mutation(() => EventObject)
  updateEvent(
    @Context() context: RequestContext,
    @Args("provider", { type: () => CalendarProvider }) provider: Provider,
    @Args("calendarId") calendarId: string,
    @Args("eventId") eventId: string,
    @Args("input") input: UpdateEventInput,
  ) {
    return this.service.update(
      context.req.tenantId!,
      provider,
      calendarId,
      eventId,
      input,
    );
  }
  @Mutation(() => Boolean)
  deleteEvent(
    @Context() context: RequestContext,
    @Args("provider", { type: () => CalendarProvider }) provider: Provider,
    @Args("calendarId") calendarId: string,
    @Args("eventId") eventId: string,
  ) {
    return this.service.remove(
      context.req.tenantId!,
      provider,
      calendarId,
      eventId,
    );
  }
  @Mutation(() => WatchObject)
  watchCalendar(
    @Context() context: RequestContext,
    @Args("provider", { type: () => CalendarProvider }) provider: Provider,
    @Args("calendarId") calendarId: string,
  ) {
    if (!this.publicUrl)
      fail("PUBLIC_URL must be configured for provider webhooks");
    return this.service.watch(
      context.req.tenantId!,
      provider,
      calendarId,
      this.publicUrl,
    );
  }
  @Mutation(() => WatchObject)
  renewWatchingCalendar(@Context() context: RequestContext, @Args("provider", {type: () => CalendarProvider}) provider: Provider, @Args("calendarId") calendarId: string, @Args("watchId") watchId: string) {
    return this.service.renew(context.req.tenantId!, provider, calendarId, watchId);
  }
  @Mutation(() => Boolean)
  stopWatchingCalendar(
    @Context() context: RequestContext,
    @Args("provider", { type: () => CalendarProvider }) provider: Provider,
    @Args("calendarId") calendarId: string,
    @Args("watchId") id: string,
    @Args("resourceId", { nullable: true }) resourceId?: string,
  ) {
    return this.service.unwatch(
      context.req.tenantId!,
      provider,
      calendarId,
      id,
      resourceId,
    );
  }
}
