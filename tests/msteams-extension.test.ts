import { describe, expect, it } from "vitest";
import {
  createMsteamsExtension,
  type MsteamsPluginState,
} from "../src/msteams-extension";

function freshState(): MsteamsPluginState {
  return {
    initialized: false,
    startedAt: null,
    teams: new Map(),
    channels: new Map(),
    tenants: new Set(),
    messagesProcessed: 0,
    totalConversations: 0,
  };
}

describe("MsteamsExtension WebMCP methods", () => {
  describe("getStatus", () => {
    it("should return offline status when adapter is null", () => {
      const state = freshState();
      const ext = createMsteamsExtension(
        () => state,
      );

      const status = ext.getStatus();
      expect(status).toEqual({
        online: false,
        connectedTenants: 0,
        latencyMs: -1,
        uptimeMs: null,
      });
    });

    it("should return online status with state data", () => {
      const now = Date.now();
      const state: MsteamsPluginState = {
        initialized: true,
        startedAt: now - 5000,
        teams: new Map([["t1", { id: "t1", name: "Engineering" }]]),
        channels: new Map(),
        tenants: new Set(["tenant-a", "tenant-b"]),
        messagesProcessed: 10,
        totalConversations: 1,
      };

      const ext = createMsteamsExtension(
        () => state,
      );

      const status = ext.getStatus();
      expect(status.online).toBe(true);
      expect(status.connectedTenants).toBe(2);
      expect(status.latencyMs).toBe(-1);
      expect(status.uptimeMs).toBeGreaterThanOrEqual(5000);
      expect(status.uptimeMs).toBeLessThan(10000);
    });
  });

  describe("listTeams", () => {
    it("should return empty array when no teams tracked", () => {
      const state = freshState();
      const ext = createMsteamsExtension(
        () => state,
      );

      expect(ext.listTeams()).toEqual([]);
    });

    it("should return team names from state", () => {
      const state: MsteamsPluginState = {
        ...freshState(),
        teams: new Map([
          ["t1", { id: "t1", name: "Engineering" }],
          ["t2", { id: "t2", name: "Marketing" }],
        ]),
      };

      const ext = createMsteamsExtension(
        () => state,
      );

      const teams = ext.listTeams();
      expect(teams).toHaveLength(2);
      expect(teams).toContainEqual({ id: "t1", name: "Engineering" });
      expect(teams).toContainEqual({ id: "t2", name: "Marketing" });
    });
  });

  describe("listChannels", () => {
    it("should return empty array when no channels tracked", () => {
      const state = freshState();
      const ext = createMsteamsExtension(
        () => state,
      );

      expect(ext.listChannels()).toEqual([]);
    });

    it("should return all channels when no teamId filter", () => {
      const state: MsteamsPluginState = {
        ...freshState(),
        channels: new Map([
          [
            "t1",
            new Map([
              ["ch1", { id: "ch1", name: "general", type: "standard" }],
              ["ch2", { id: "ch2", name: "announcements", type: "standard" }],
            ]),
          ],
          [
            "t2",
            new Map([["ch3", { id: "ch3", name: "private-ch", type: "private" }]]),
          ],
        ]),
      };

      const ext = createMsteamsExtension(
        () => state,
      );

      const channels = ext.listChannels();
      expect(channels).toHaveLength(3);
      expect(channels).toContainEqual({ id: "ch1", name: "general", type: "standard" });
      expect(channels).toContainEqual({ id: "ch3", name: "private-ch", type: "private" });
    });

    it("should return channels for specific team when teamId provided", () => {
      const state: MsteamsPluginState = {
        ...freshState(),
        channels: new Map([
          [
            "t1",
            new Map([
              ["ch1", { id: "ch1", name: "general", type: "standard" }],
              ["ch2", { id: "ch2", name: "dev", type: "standard" }],
            ]),
          ],
          [
            "t2",
            new Map([["ch3", { id: "ch3", name: "marketing", type: "standard" }]]),
          ],
        ]),
      };

      const ext = createMsteamsExtension(
        () => state,
      );

      const channels = ext.listChannels("t1");
      expect(channels).toHaveLength(2);
      expect(channels).toContainEqual({ id: "ch1", name: "general", type: "standard" });
      expect(channels).toContainEqual({ id: "ch2", name: "dev", type: "standard" });
    });

    it("should return empty array for unknown teamId", () => {
      const state: MsteamsPluginState = {
        ...freshState(),
        channels: new Map([
          ["t1", new Map([["ch1", { id: "ch1", name: "general", type: "standard" }]])],
        ]),
      };

      const ext = createMsteamsExtension(
        () => state,
      );

      expect(ext.listChannels("nonexistent")).toEqual([]);
    });
  });

  describe("getMessageStats", () => {
    it("should return zeros when state is fresh", () => {
      const state = freshState();
      const ext = createMsteamsExtension(
        () => state,
      );

      const stats = ext.getMessageStats();
      expect(stats).toEqual({ messagesProcessed: 0, activeConversations: 0 });
    });

    it("should count only msteams sessions", () => {
      const state: MsteamsPluginState = {
        ...freshState(),
        messagesProcessed: 25,
        totalConversations: 3,
      };

      const ext = createMsteamsExtension(
        () => state,
      );

      const stats = ext.getMessageStats();
      expect(stats).toEqual({ messagesProcessed: 25, activeConversations: 3 });
    });
  });
});
