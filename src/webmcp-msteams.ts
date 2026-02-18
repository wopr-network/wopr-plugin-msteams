/**
 * WebMCP MS Teams Tools
 *
 * Registers 4 read-only browser-side WebMCP tools for MS Teams connection
 * status, team listing, channel listing, and message stats.
 *
 * These tools call the WOPR daemon REST API via fetch() and are only
 * meaningful when the MS Teams plugin is loaded on the instance.
 */

// ============================================================================
// Types (mirrors WebMCPRegistry from wopr-plugin-webui)
// ============================================================================

export interface AuthContext {
  token?: string;
  [key: string]: unknown;
}

export interface ParameterSchema {
  type: string;
  description: string;
  required?: boolean;
}

export interface WebMCPTool {
  name: string;
  description: string;
  parameters: Record<string, ParameterSchema>;
  handler: (params: Record<string, unknown>, auth: AuthContext) => Promise<unknown>;
}

export interface WebMCPRegistry {
  register(tool: WebMCPTool): void;
  get(name: string): WebMCPTool | undefined;
  list(): string[];
}

// ============================================================================
// Internal helpers
// ============================================================================

interface RequestOptions {
  method?: string;
  body?: string;
  headers?: Record<string, string>;
}

async function daemonRequest<T>(
  apiBase: string,
  path: string,
  auth: AuthContext,
  options?: RequestOptions,
): Promise<T> {
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    ...options?.headers,
  };
  if (auth.token) {
    headers.Authorization = `Bearer ${auth.token as string}`;
  }
  const res = await fetch(`${apiBase}${path}`, {
    ...options,
    headers,
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ error: "Request failed" }));
    throw new Error((err as { error?: string }).error || `Request failed (${res.status})`);
  }
  return res.json() as Promise<T>;
}

// ============================================================================
// Tool registration
// ============================================================================

/**
 * Register all 4 MS Teams WebMCP tools on the given registry.
 *
 * The tools proxy to the MS Teams plugin extension via daemon API endpoints
 * at `/plugins/msteams/status`, `/plugins/msteams/teams`, etc.
 *
 * @param registry - The WebMCPRegistry instance to register tools on
 * @param apiBase  - Base URL of the WOPR daemon API (e.g. "/api" or "http://localhost:7437")
 */
export function registerMsteamsTools(registry: WebMCPRegistry, apiBase = "/api"): void {
  // 1. getMsteamsStatus
  registry.register({
    name: "getMsteamsStatus",
    description: "Get MS Teams bot connection status: online/offline, connected tenants, and uptime.",
    parameters: {},
    handler: async (_params: Record<string, unknown>, auth: AuthContext) => {
      return daemonRequest(apiBase, "/plugins/msteams/status", auth);
    },
  });

  // 2. listTeams
  registry.register({
    name: "listTeams",
    description: "List connected MS Teams organizations the bot has interacted with.",
    parameters: {},
    handler: async (_params: Record<string, unknown>, auth: AuthContext) => {
      return daemonRequest(apiBase, "/plugins/msteams/teams", auth);
    },
  });

  // 3. listMsteamsChannels
  registry.register({
    name: "listMsteamsChannels",
    description: "List MS Teams channels the bot is active in, optionally filtered by team.",
    parameters: {
      teamId: {
        type: "string",
        description: "The Teams organization ID to filter channels for",
        required: false,
      },
    },
    handler: async (params: Record<string, unknown>, auth: AuthContext) => {
      const teamId = params.teamId as string | undefined;
      if (teamId) {
        return daemonRequest(
          apiBase,
          `/plugins/msteams/teams/${encodeURIComponent(teamId)}/channels`,
          auth,
        );
      }
      return daemonRequest(apiBase, "/plugins/msteams/channels", auth);
    },
  });

  // 4. getMsteamsMessageStats
  registry.register({
    name: "getMsteamsMessageStats",
    description:
      "Get MS Teams message processing statistics: messages processed and active conversations.",
    parameters: {},
    handler: async (_params: Record<string, unknown>, auth: AuthContext) => {
      return daemonRequest(apiBase, "/plugins/msteams/stats", auth);
    },
  });
}
