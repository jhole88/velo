import { describe, it, expect, beforeEach, vi } from "vitest";

vi.mock("@/services/graph/graphClient", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@/services/graph/graphClient")>();
  return { ...actual, getGraphClient: vi.fn() };
});

import { GraphError, getGraphClient, type GraphClient } from "@/services/graph/graphClient";
import { GraphCalendarProvider, mapGraphEvent } from "./graphCalendarProvider";

function createMockClient() {
  return { request: vi.fn(), getAll: vi.fn(), forEachPage: vi.fn() };
}

describe("mapGraphEvent", () => {
  it("maps a timed event from UTC", () => {
    const event = mapGraphEvent({
      id: "ev-1",
      iCalUId: "uid-1",
      "@odata.etag": "W/\"1\"",
      subject: "Dept leads meeting",
      bodyPreview: "Agenda",
      location: { displayName: "Boardroom" },
      start: { dateTime: "2026-11-13T15:00:00.0000000", timeZone: "UTC" },
      end: { dateTime: "2026-11-13T16:00:00.0000000", timeZone: "UTC" },
      organizer: { emailAddress: { address: "boss@example.com" } },
      attendees: [{ emailAddress: { address: "a@example.com", name: "A" }, status: { response: "tentativelyAccepted" } }],
      webLink: "https://outlook.office.com/calendar/item/1",
    });

    expect(event).toEqual({
      remoteEventId: "ev-1",
      uid: "uid-1",
      etag: "W/\"1\"",
      summary: "Dept leads meeting",
      description: "Agenda",
      location: "Boardroom",
      startTime: Date.UTC(2026, 10, 13, 15) / 1000,
      endTime: Date.UTC(2026, 10, 13, 16) / 1000,
      isAllDay: false,
      status: "confirmed",
      organizerEmail: "boss@example.com",
      attendeesJson: JSON.stringify([{ email: "a@example.com", displayName: "A", responseStatus: "tentative" }]),
      htmlLink: "https://outlook.office.com/calendar/item/1",
      icalData: null,
    });
  });

  it("maps an all-day event to local midnight through the end of its last day", () => {
    const event = mapGraphEvent({
      id: "ev-2",
      isAllDay: true,
      start: { dateTime: "2026-12-24T00:00:00.0000000", timeZone: "UTC" },
      end: { dateTime: "2026-12-25T00:00:00.0000000", timeZone: "UTC" },
    });

    expect(event.isAllDay).toBe(true);
    expect(event.startTime).toBe(Math.floor(new Date("2026-12-24T00:00:00").getTime() / 1000));
    expect(event.endTime).toBe(Math.floor(new Date("2026-12-24T23:59:59").getTime() / 1000));
  });

  it("marks cancelled events", () => {
    expect(mapGraphEvent({ id: "ev-3", isCancelled: true }).status).toBe("cancelled");
  });
});

describe("GraphCalendarProvider", () => {
  let client: ReturnType<typeof createMockClient>;
  let provider: GraphCalendarProvider;

  beforeEach(() => {
    vi.clearAllMocks();
    client = createMockClient();
    vi.mocked(getGraphClient).mockReturnValue(client as unknown as GraphClient);
    provider = new GraphCalendarProvider("acc-1");
  });

  it("lists calendars", async () => {
    client.getAll.mockResolvedValue([
      { id: "cal-1", name: "Calendar", hexColor: "#0078d4", isDefaultCalendar: true },
      { id: "cal-2", name: "Site visits", hexColor: "" },
    ]);

    await expect(provider.listCalendars()).resolves.toEqual([
      { remoteId: "cal-1", displayName: "Calendar", color: "#0078d4", isPrimary: true },
      { remoteId: "cal-2", displayName: "Site visits", color: null, isPrimary: false },
    ]);
  });

  it("creates a timed event in UTC with attendees", async () => {
    client.request.mockResolvedValue({ id: "ev-new", subject: "Kickoff" });

    await provider.createEvent("cal-1", {
      summary: "Kickoff",
      location: "Site trailer",
      startTime: "2026-11-20T17:00:00Z",
      endTime: "2026-11-20T18:00:00Z",
      attendees: [{ email: "pm@example.com" }],
    });

    const [url, options] = client.request.mock.calls[0]!;
    expect(url).toBe("/me/calendars/cal-1/events");
    expect(options.method).toBe("POST");
    expect(options.body).toEqual({
      subject: "Kickoff",
      location: { displayName: "Site trailer" },
      isAllDay: false,
      start: { dateTime: "2026-11-20T17:00:00.000", timeZone: "UTC" },
      end: { dateTime: "2026-11-20T18:00:00.000", timeZone: "UTC" },
      attendees: [{ emailAddress: { address: "pm@example.com" }, type: "required" }],
    });
  });

  it("creates an all-day event ending at the next midnight", async () => {
    client.request.mockResolvedValue({ id: "ev-new" });

    await provider.createEvent("cal-1", {
      summary: "Shutdown",
      startTime: "2026-12-24T00:00:00",
      endTime: "2026-12-24T23:59:59",
      isAllDay: true,
    });

    expect(client.request.mock.calls[0]![1].body).toMatchObject({
      isAllDay: true,
      start: { dateTime: "2026-12-24T00:00:00", timeZone: "UTC" },
      end: { dateTime: "2026-12-25T00:00:00", timeZone: "UTC" },
    });
  });

  it("deletes an event", async () => {
    await provider.deleteEvent("cal-1", "ev-1");
    expect(client.request).toHaveBeenCalledWith("/me/calendars/cal-1/events/ev-1", { method: "DELETE" });
  });

  describe("syncEvents", () => {
    it("starts a windowed delta on first sync and returns a token", async () => {
      client.forEachPage.mockImplementation(async (_url: string, onPage: (items: unknown[]) => void) => {
        onPage([{ id: "ev-1", subject: "A" }, { id: "ev-gone", "@removed": { reason: "deleted" } }]);
        return "https://graph.microsoft.com/v1.0/cal-delta";
      });

      const result = await provider.syncEvents("cal-1");

      expect(client.forEachPage.mock.calls[0]![0]).toMatch(/^\/me\/calendars\/cal-1\/calendarView\/delta\?startDateTime=/);
      expect(result.created.map((e) => e.remoteEventId)).toEqual(["ev-1"]);
      expect(result.deletedRemoteIds).toEqual(["ev-gone"]);
      expect(JSON.parse(result.newSyncToken!)).toMatchObject({ link: "https://graph.microsoft.com/v1.0/cal-delta" });
    });

    it("continues from a recent token", async () => {
      client.forEachPage.mockResolvedValue("https://graph.microsoft.com/v1.0/cal-delta-2");
      const token = JSON.stringify({ link: "https://graph.microsoft.com/v1.0/cal-delta", createdAt: Date.now() });

      await provider.syncEvents("cal-1", token);

      expect(client.forEachPage.mock.calls[0]![0]).toBe("https://graph.microsoft.com/v1.0/cal-delta");
    });

    it("rebuilds the window when the token is over 30 days old", async () => {
      client.forEachPage.mockResolvedValue("https://graph.microsoft.com/v1.0/cal-delta-2");
      const old = Date.now() - 31 * 24 * 60 * 60 * 1000;
      const token = JSON.stringify({ link: "https://graph.microsoft.com/v1.0/cal-delta", createdAt: old });

      const result = await provider.syncEvents("cal-1", token);

      expect(client.forEachPage.mock.calls[0]![0]).toContain("/calendarView/delta?startDateTime=");
      expect(JSON.parse(result.newSyncToken!).createdAt).toBeGreaterThan(old);
    });

    it("starts over when the delta link has expired", async () => {
      client.forEachPage
        .mockRejectedValueOnce(new GraphError(410, "SyncStateNotFound", "expired"))
        .mockImplementationOnce(async (_url: string, onPage: (items: unknown[]) => void) => {
          onPage([{ id: "ev-1" }]);
          return "https://graph.microsoft.com/v1.0/cal-delta-fresh";
        });
      const token = JSON.stringify({ link: "https://graph.microsoft.com/v1.0/cal-delta", createdAt: Date.now() });

      const result = await provider.syncEvents("cal-1", token);

      expect(client.forEachPage).toHaveBeenCalledTimes(2);
      expect(result.created).toHaveLength(1);
      expect(JSON.parse(result.newSyncToken!).link).toBe("https://graph.microsoft.com/v1.0/cal-delta-fresh");
    });
  });
});
