import { beforeEach, describe, expect, it, vi } from "vitest";
import { type AuthContext, type WebMCPRegistry, registerMsteamsTools } from "../src/webmcp-msteams";

// Mock global fetch
const mockFetch = vi.fn();
vi.stubGlobal("fetch", mockFetch);

function mockJsonResponse(data: unknown, ok = true, status = 200) {
  return {
    ok,
    status,
    json: vi.fn().mockResolvedValue(data),
  };
}

function createRegistry(): WebMCPRegistry {
  const tools = new Map<string, { name: string; handler: Function; [k: string]: unknown }>();
  return {
    register(tool: { name: string; handler: Function }) {
      tools.set(tool.name, tool);
    },
    get(name: string) {
      return tools.get(name) as any;
    },
    list() {
      return [...tools.keys()];
    },
  };
}

function getTool(registry: WebMCPRegistry, name: string) {
  const tool = registry.get(name);
  if (!tool) throw new Error(`Tool "${name}" not registered`);
  return tool;
}

describe("registerMsteamsTools", () => {
  let registry: WebMCPRegistry;
  const API_BASE = "/api";

  beforeEach(() => {
    registry = createRegistry();
    mockFetch.mockReset();
  });

  it("should register all 4 tools", () => {
    registerMsteamsTools(registry, API_BASE);

    const names = registry.list();
    expect(names).toHaveLength(4);
    expect(names).toContain("getMsteamsStatus");
    expect(names).toContain("listTeams");
    expect(names).toContain("listMsteamsChannels");
    expect(names).toContain("getMsteamsMessageStats");
  });

  it("should use default apiBase when not provided", () => {
    registerMsteamsTools(registry);

    expect(registry.list()).toHaveLength(4);
  });

  describe("getMsteamsStatus", () => {
    it("should GET /plugins/msteams/status", async () => {
      const status = { online: true, connectedTenants: 2, latencyMs: -1, uptimeMs: 360000 };
      mockFetch.mockResolvedValue(mockJsonResponse(status));
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "getMsteamsStatus");
      const result = await tool.handler({}, {});

      expect(mockFetch).toHaveBeenCalledWith("/api/plugins/msteams/status", expect.any(Object));
      expect(result).toEqual(status);
    });

    it("should include bearer token when auth.token is present", async () => {
      mockFetch.mockResolvedValue(mockJsonResponse({ online: true }));
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "getMsteamsStatus");
      const auth: AuthContext = { token: "my-token" };
      await tool.handler({}, auth);

      const headers = mockFetch.mock.calls[0][1].headers;
      expect(headers.Authorization).toBe("Bearer my-token");
    });

    it("should not include Authorization header when no token", async () => {
      mockFetch.mockResolvedValue(mockJsonResponse({ online: false }));
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "getMsteamsStatus");
      await tool.handler({}, {});

      const headers = mockFetch.mock.calls[0][1].headers;
      expect(headers.Authorization).toBeUndefined();
    });
  });

  describe("listTeams", () => {
    it("should GET /plugins/msteams/teams", async () => {
      const teams = { teams: [{ id: "t1", name: "Engineering" }] };
      mockFetch.mockResolvedValue(mockJsonResponse(teams));
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "listTeams");
      const result = await tool.handler({}, {});

      expect(mockFetch).toHaveBeenCalledWith("/api/plugins/msteams/teams", expect.any(Object));
      expect(result).toEqual(teams);
    });

    it("should include bearer token in auth header", async () => {
      mockFetch.mockResolvedValue(mockJsonResponse({ teams: [] }));
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "listTeams");
      await tool.handler({}, { token: "tok-teams" });

      const headers = mockFetch.mock.calls[0][1].headers;
      expect(headers.Authorization).toBe("Bearer tok-teams");
    });
  });

  describe("listMsteamsChannels", () => {
    it("should GET /plugins/msteams/channels when no teamId", async () => {
      const channels = { channels: [{ id: "ch1", name: "general", type: "standard" }] };
      mockFetch.mockResolvedValue(mockJsonResponse(channels));
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "listMsteamsChannels");
      const result = await tool.handler({}, {});

      expect(mockFetch).toHaveBeenCalledWith("/api/plugins/msteams/channels", expect.any(Object));
      expect(result).toEqual(channels);
    });

    it("should GET /plugins/msteams/teams/:teamId/channels when teamId provided", async () => {
      const channels = { channels: [{ id: "ch1", name: "general", type: "standard" }] };
      mockFetch.mockResolvedValue(mockJsonResponse(channels));
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "listMsteamsChannels");
      const result = await tool.handler({ teamId: "team-123" }, {});

      expect(mockFetch).toHaveBeenCalledWith(
        "/api/plugins/msteams/teams/team-123/channels",
        expect.any(Object),
      );
      expect(result).toEqual(channels);
    });

    it("should URL-encode teamId with special characters", async () => {
      mockFetch.mockResolvedValue(mockJsonResponse({ channels: [] }));
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "listMsteamsChannels");
      await tool.handler({ teamId: "id with spaces" }, {});

      expect(mockFetch).toHaveBeenCalledWith(
        "/api/plugins/msteams/teams/id%20with%20spaces/channels",
        expect.any(Object),
      );
    });

    it("should include bearer token in auth header", async () => {
      mockFetch.mockResolvedValue(mockJsonResponse({ channels: [] }));
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "listMsteamsChannels");
      await tool.handler({}, { token: "tok-ch" });

      const headers = mockFetch.mock.calls[0][1].headers;
      expect(headers.Authorization).toBe("Bearer tok-ch");
    });
  });

  describe("getMsteamsMessageStats", () => {
    it("should GET /plugins/msteams/stats", async () => {
      const stats = { messagesProcessed: 42, activeConversations: 3 };
      mockFetch.mockResolvedValue(mockJsonResponse(stats));
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "getMsteamsMessageStats");
      const result = await tool.handler({}, {});

      expect(mockFetch).toHaveBeenCalledWith("/api/plugins/msteams/stats", expect.any(Object));
      expect(result).toEqual(stats);
    });

    it("should include bearer token in auth header", async () => {
      mockFetch.mockResolvedValue(mockJsonResponse({ messagesProcessed: 0 }));
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "getMsteamsMessageStats");
      await tool.handler({}, { token: "tok-stats" });

      const headers = mockFetch.mock.calls[0][1].headers;
      expect(headers.Authorization).toBe("Bearer tok-stats");
    });
  });

  describe("error handling", () => {
    it("should throw on non-ok response with error from body", async () => {
      mockFetch.mockResolvedValue(
        mockJsonResponse({ error: "MS Teams plugin not loaded" }, false, 404),
      );
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "getMsteamsStatus");

      await expect(tool.handler({}, {})).rejects.toThrow("MS Teams plugin not loaded");
    });

    it("should throw with status code when body has no error field", async () => {
      mockFetch.mockResolvedValue(mockJsonResponse({}, false, 500));
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "listTeams");

      await expect(tool.handler({}, {})).rejects.toThrow("Request failed (500)");
    });

    it("should handle json parse failure on error response", async () => {
      mockFetch.mockResolvedValue({
        ok: false,
        status: 502,
        json: vi.fn().mockRejectedValue(new Error("invalid json")),
      });
      registerMsteamsTools(registry, API_BASE);

      const tool = getTool(registry, "getMsteamsMessageStats");

      await expect(tool.handler({}, {})).rejects.toThrow("Request failed");
    });
  });
});
