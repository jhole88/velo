import type {
  CalendarProvider,
  CalendarProviderType,
  CalendarInfo,
  CalendarEventData,
  CalendarSyncResult,
  CreateEventInput,
  UpdateEventInput,
} from "./types";
import { GraphError, getGraphClient, type GraphClient } from "@/services/graph/graphClient";

/** Ask Graph for UTC times so no Windows time-zone names need translating. */
const UTC_PREFERENCE = 'outlook.timezone="UTC"';

/** The sync window is fixed when a delta link is created, so rebuild it periodically. */
const WINDOW_REFRESH_MS = 30 * 24 * 60 * 60 * 1000;

interface GraphCalendar {
  id: string;
  name: string;
  hexColor?: string | null;
  isDefaultCalendar?: boolean | null;
}

interface GraphDateTime {
  dateTime: string;
  timeZone?: string;
}

interface GraphEvent {
  id: string;
  iCalUId?: string | null;
  "@odata.etag"?: string;
  subject?: string | null;
  bodyPreview?: string | null;
  location?: { displayName?: string | null } | null;
  start?: GraphDateTime | null;
  end?: GraphDateTime | null;
  isAllDay?: boolean | null;
  isCancelled?: boolean | null;
  showAs?: string | null;
  organizer?: { emailAddress?: { address?: string | null; name?: string | null } } | null;
  attendees?: {
    emailAddress?: { address?: string | null; name?: string | null };
    status?: { response?: string | null };
  }[] | null;
  webLink?: string | null;
  "@removed"?: { reason?: string };
}

interface StoredSyncToken {
  link: string;
  createdAt: number;
}

/** Graph returns UTC times without a zone suffix: "2026-11-13T15:00:00.0000000". */
function utcToEpochSeconds(value: GraphDateTime | null | undefined): number {
  if (!value?.dateTime) return 0;
  const iso = /[zZ]|[+-]\d{2}:\d{2}$/.test(value.dateTime) ? value.dateTime : `${value.dateTime}Z`;
  return Math.floor(new Date(iso).getTime() / 1000);
}

const RESPONSE_STATUS: Record<string, string> = {
  accepted: "accepted",
  organizer: "accepted",
  declined: "declined",
  tentativelyAccepted: "tentative",
  notResponded: "needsAction",
  none: "needsAction",
};

export function mapGraphEvent(event: GraphEvent): CalendarEventData {
  const isAllDay = !!event.isAllDay;
  let startTime = utcToEpochSeconds(event.start);
  let endTime = utcToEpochSeconds(event.end);

  if (isAllDay && event.start?.dateTime && event.end?.dateTime) {
    // All-day events are stored as local midnight → 23:59:59 (same as the Google provider).
    // Graph's end date is exclusive, so step back one second.
    const startDate = event.start.dateTime.slice(0, 10);
    const endDate = event.end.dateTime.slice(0, 10);
    startTime = Math.floor(new Date(`${startDate}T00:00:00`).getTime() / 1000);
    endTime = Math.floor(new Date(`${endDate}T00:00:00`).getTime() / 1000) - 1;
  }

  const attendees = (event.attendees ?? [])
    .filter((a) => a.emailAddress?.address)
    .map((a) => ({
      email: a.emailAddress!.address!,
      displayName: a.emailAddress?.name ?? undefined,
      responseStatus: RESPONSE_STATUS[a.status?.response ?? "none"] ?? "needsAction",
    }));

  return {
    remoteEventId: event.id,
    uid: event.iCalUId ?? null,
    etag: event["@odata.etag"] ?? null,
    summary: event.subject ?? null,
    description: event.bodyPreview ?? null,
    location: event.location?.displayName || null,
    startTime,
    endTime,
    isAllDay,
    status: event.isCancelled ? "cancelled" : event.showAs === "tentative" ? "tentative" : "confirmed",
    organizerEmail: event.organizer?.emailAddress?.address ?? null,
    attendeesJson: attendees.length > 0 ? JSON.stringify(attendees) : null,
    htmlLink: event.webLink ?? null,
    icalData: null,
  };
}

function eventBody(event: CreateEventInput | UpdateEventInput): Record<string, unknown> {
  const body: Record<string, unknown> = {};
  if (event.summary !== undefined) body.subject = event.summary;
  if (event.description !== undefined) body.body = { contentType: "text", content: event.description };
  if (event.location !== undefined) body.location = { displayName: event.location };

  if (event.startTime && event.endTime) {
    if (event.isAllDay) {
      // All-day events must start and end at midnight; the end date is exclusive
      const start = event.startTime.split("T")[0]!;
      const end = new Date(`${event.endTime.split("T")[0]}T00:00:00Z`);
      end.setUTCDate(end.getUTCDate() + 1);
      body.isAllDay = true;
      body.start = { dateTime: `${start}T00:00:00`, timeZone: "UTC" };
      body.end = { dateTime: `${end.toISOString().slice(0, 10)}T00:00:00`, timeZone: "UTC" };
    } else {
      body.isAllDay = false;
      body.start = { dateTime: new Date(event.startTime).toISOString().replace("Z", ""), timeZone: "UTC" };
      body.end = { dateTime: new Date(event.endTime).toISOString().replace("Z", ""), timeZone: "UTC" };
    }
  }

  return body;
}

/** Outlook / Microsoft 365 calendar through Microsoft Graph. */
export class GraphCalendarProvider implements CalendarProvider {
  readonly type: CalendarProviderType = "microsoft_graph";

  constructor(readonly accountId: string) {}

  private get client(): GraphClient {
    return getGraphClient(this.accountId);
  }

  private calendarPath(calendarRemoteId: string): string {
    return `/me/calendars/${encodeURIComponent(calendarRemoteId)}`;
  }

  async listCalendars(): Promise<CalendarInfo[]> {
    const calendars = await this.client.getAll<GraphCalendar>(
      "/me/calendars?$select=id,name,hexColor,isDefaultCalendar&$top=100",
    );
    return calendars.map((cal) => ({
      remoteId: cal.id,
      displayName: cal.name,
      color: cal.hexColor || null,
      isPrimary: !!cal.isDefaultCalendar,
    }));
  }

  async fetchEvents(calendarRemoteId: string, timeMin: string, timeMax: string): Promise<CalendarEventData[]> {
    const params = new URLSearchParams({
      startDateTime: new Date(timeMin).toISOString(),
      endDateTime: new Date(timeMax).toISOString(),
      $top: "100",
    });
    const events = await this.client.getAll<GraphEvent>(
      `${this.calendarPath(calendarRemoteId)}/calendarView?${params}`,
      { prefer: [UTC_PREFERENCE] },
    );
    return events.map(mapGraphEvent);
  }

  async createEvent(calendarRemoteId: string, event: CreateEventInput): Promise<CalendarEventData> {
    const body = eventBody(event);
    if (event.attendees) {
      body.attendees = event.attendees.map((a) => ({ emailAddress: { address: a.email }, type: "required" }));
    }
    const created = await this.client.request<GraphEvent>(`${this.calendarPath(calendarRemoteId)}/events`, {
      method: "POST",
      body,
      prefer: [UTC_PREFERENCE],
    });
    return mapGraphEvent(created);
  }

  async updateEvent(calendarRemoteId: string, remoteEventId: string, event: UpdateEventInput): Promise<CalendarEventData> {
    const updated = await this.client.request<GraphEvent>(
      `${this.calendarPath(calendarRemoteId)}/events/${encodeURIComponent(remoteEventId)}`,
      { method: "PATCH", body: eventBody(event), prefer: [UTC_PREFERENCE] },
    );
    return mapGraphEvent(updated);
  }

  async deleteEvent(calendarRemoteId: string, remoteEventId: string): Promise<void> {
    await this.client.request(
      `${this.calendarPath(calendarRemoteId)}/events/${encodeURIComponent(remoteEventId)}`,
      { method: "DELETE" },
    );
  }

  /**
   * Sync a window of 90 days back to one year ahead. The first call (or an
   * expired/aged token) returns the whole window; later calls return changes.
   */
  async syncEvents(calendarRemoteId: string, syncToken?: string): Promise<CalendarSyncResult> {
    const created: CalendarEventData[] = [];
    const deletedRemoteIds: string[] = [];

    let startUrl = this.initialDeltaUrl(calendarRemoteId);
    let createdAt = Date.now();
    const stored = parseSyncToken(syncToken);
    if (stored && Date.now() - stored.createdAt < WINDOW_REFRESH_MS) {
      startUrl = stored.link;
      createdAt = stored.createdAt;
    }

    const collect = (items: GraphEvent[]): void => {
      for (const item of items) {
        if (item["@removed"]) {
          deletedRemoteIds.push(item.id);
        } else {
          // Everything goes in "created" — the DB upsert handles updates
          created.push(mapGraphEvent(item));
        }
      }
    };

    let deltaLink: string | null;
    try {
      deltaLink = await this.client.forEachPage<GraphEvent>(startUrl, collect, { prefer: [UTC_PREFERENCE] });
    } catch (err) {
      const expired =
        err instanceof GraphError && (err.status === 410 || !!err.code?.toLowerCase().includes("syncstate"));
      const initialUrl = this.initialDeltaUrl(calendarRemoteId);
      if (!expired || startUrl === initialUrl) throw err;

      // Delta link expired — start the window over
      created.length = 0;
      deletedRemoteIds.length = 0;
      createdAt = Date.now();
      deltaLink = await this.client.forEachPage<GraphEvent>(initialUrl, collect, { prefer: [UTC_PREFERENCE] });
    }

    const newSyncToken = deltaLink
      ? JSON.stringify({ link: deltaLink, createdAt } satisfies StoredSyncToken)
      : null;
    return { created, updated: [], deletedRemoteIds, newSyncToken, newCtag: null };
  }

  async testConnection(): Promise<{ success: boolean; message: string }> {
    try {
      await this.listCalendars();
      return { success: true, message: "Connected to Outlook calendar" };
    } catch (err) {
      return { success: false, message: err instanceof Error ? err.message : "Connection failed" };
    }
  }

  private initialDeltaUrl(calendarRemoteId: string): string {
    const start = new Date();
    start.setDate(start.getDate() - 90);
    const end = new Date();
    end.setFullYear(end.getFullYear() + 1);
    const params = new URLSearchParams({
      startDateTime: start.toISOString(),
      endDateTime: end.toISOString(),
    });
    return `${this.calendarPath(calendarRemoteId)}/calendarView/delta?${params}`;
  }
}

function parseSyncToken(token?: string): StoredSyncToken | null {
  if (!token) return null;
  try {
    const parsed = JSON.parse(token) as Partial<StoredSyncToken>;
    if (typeof parsed.link === "string" && typeof parsed.createdAt === "number") {
      return { link: parsed.link, createdAt: parsed.createdAt };
    }
  } catch {
    // Not one of ours — start fresh
  }
  return null;
}
