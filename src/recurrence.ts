// Extracted from Spantrek's GoogleRecurrencePatternAdapter and Office365RecurrencePatternAdapter.
import moment from "moment-timezone";
import { fail } from "./config";
import { Recurrence } from "./types";

const DAYS = [
  "sunday",
  "monday",
  "tuesday",
  "wednesday",
  "thursday",
  "friday",
  "saturday",
];
const GOOGLE_DAYS = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"];
const WEEKS = ["first", "second", "third", "fourth"];
function integer(value: unknown, min: number, max: number): void {
  if (
    !Number.isInteger(value) ||
    (value as number) < min ||
    (value as number) > max
  )
    fail(`Recurrence value must be an integer between ${min} and ${max}`);
}
function validate(r: Recurrence, start: string, zone: string): void {
  integer(r.interval, 1, 999);
  if (
    ![
      "DAY",
      "WEEK",
      "AbsoluteMonthly",
      "RelativeMonthly",
      "AbsoluteYearly",
      "RelativeYearly",
    ].includes(r.recurrenceType)
  )
    fail("Unsupported recurrence type");
  if (r.recurrenceType === "WEEK") {
    if (!r.daysOfWeek?.length || r.daysOfWeek.length > 7)
      fail("Weekly recurrence requires daysOfWeek");
    r.daysOfWeek.forEach((d) => integer(d, 0, 6));
  }
  if (r.recurrenceType.startsWith("Absolute")) integer(r.dayOfMonth, 1, 31);
  if (r.recurrenceType.startsWith("Relative")) {
    integer(r.weekDay, 0, 6);
    integer(r.weekIndex, 0, 3);
  }
  if (r.recurrenceType.endsWith("Yearly")) integer(r.month, 0, 11);
  if (r.rangeType === "OCCURRENCE_COUNT") integer(r.occurrenceCount, 1, 999);
  else if (r.rangeType === "UNTIL_DATE") {
    if (
      !r.untilDate ||
      !moment(r.untilDate, "YYYY-MM-DD", true).isValid() ||
      r.untilDate < moment(start).tz(zone).format("YYYY-MM-DD")
    )
      fail("untilDate must be a valid local date on or after the start");
  } else if (r.rangeType !== "NEVER") fail("Unsupported recurrence range");
}
export function googleRecurrence(
  r: Recurrence,
  start: string,
  zone: string,
): string {
  validate(r, start, zone);
  let rule: string;
  switch (r.recurrenceType) {
    case "DAY":
      rule = "RRULE:FREQ=DAILY";
      break;
    case "WEEK":
      rule = `RRULE:FREQ=WEEKLY;BYDAY=${r.daysOfWeek!.map((d) => GOOGLE_DAYS[d]).join(",")}`;
      break;
    case "AbsoluteMonthly":
      rule = `RRULE:FREQ=MONTHLY;BYMONTHDAY=${r.dayOfMonth}`;
      break;
    case "RelativeMonthly":
      rule = `RRULE:FREQ=MONTHLY;BYDAY=${r.weekIndex! + 1}${GOOGLE_DAYS[r.weekDay!]}`;
      break;
    case "AbsoluteYearly":
      rule = `RRULE:FREQ=YEARLY;BYMONTH=${r.month! + 1};BYMONTHDAY=${r.dayOfMonth}`;
      break;
    case "RelativeYearly":
      rule = `RRULE:FREQ=YEARLY;BYMONTH=${r.month! + 1};BYDAY=${r.weekIndex! + 1}${GOOGLE_DAYS[r.weekDay!]}`;
      break;
  }
  if (r.interval > 1) rule += `;INTERVAL=${r.interval}`;
  if (r.rangeType === "OCCURRENCE_COUNT") rule += `;COUNT=${r.occurrenceCount}`;
  if (r.rangeType === "UNTIL_DATE")
    rule += `;UNTIL=${moment.tz(r.untilDate!, zone).endOf("day").utc().format("YYYYMMDD[T]HHmmss[Z]")}`;
  return rule;
}
export function microsoftRecurrence(
  r: Recurrence,
  start: string,
  zone: string,
) {
  validate(r, start, zone);
  const pattern: Record<string, unknown> = {
    type: "daily",
    interval: r.interval,
  };
  switch (r.recurrenceType) {
    case "DAY":
      break;
    case "WEEK":
      Object.assign(pattern, {
        type: "weekly",
        daysOfWeek: r.daysOfWeek!.map((d) => DAYS[d]),
        firstDayOfWeek: "monday",
      });
      break;
    case "AbsoluteMonthly":
      Object.assign(pattern, {
        type: "absoluteMonthly",
        dayOfMonth: r.dayOfMonth,
      });
      break;
    case "RelativeMonthly":
      Object.assign(pattern, {
        type: "relativeMonthly",
        daysOfWeek: [DAYS[r.weekDay!]],
        index: WEEKS[r.weekIndex!],
      });
      break;
    case "AbsoluteYearly":
      Object.assign(pattern, {
        type: "absoluteYearly",
        dayOfMonth: r.dayOfMonth,
        month: r.month! + 1,
      });
      break;
    case "RelativeYearly":
      Object.assign(pattern, {
        type: "relativeYearly",
        daysOfWeek: [DAYS[r.weekDay!]],
        index: WEEKS[r.weekIndex!],
        month: r.month! + 1,
      });
      break;
  }
  const range: Record<string, unknown> = {
    type: "noEnd",
    startDate: moment(start).tz(zone).format("YYYY-MM-DD"),
    recurrenceTimeZone: zone,
  };
  if (r.rangeType === "OCCURRENCE_COUNT")
    Object.assign(range, {
      type: "numbered",
      numberOfOccurrences: r.occurrenceCount,
    });
  if (r.rangeType === "UNTIL_DATE")
    Object.assign(range, { type: "endDate", endDate: r.untilDate });
  return { pattern, range };
}
