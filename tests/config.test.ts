/**
 * Tests for MS Teams plugin config schema and credential resolution (WOP-243)
 *
 * Tests:
 * - Config schema has all required fields
 * - Config schema field types and defaults
 * - Credential resolution from config vs env vars
 * - Config priority (config values override env vars)
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createMockContext } from "./mocks/wopr-context.js";

// Track ConfigurationBotFrameworkAuthentication constructor calls
let lastAuthConfig: any = undefined;

// Mock botbuilder
vi.mock("botbuilder", () => {
  return {
    CloudAdapter: class MockCloudAdapter {
      onTurnError: any;
      constructor() {
        this.onTurnError = null;
      }
      process = vi.fn();
      continueConversationAsync = vi.fn();
    },
    ConfigurationBotFrameworkAuthentication: class MockAuth {
      config: any;
      constructor(config: any) {
        this.config = config;
        lastAuthConfig = config;
      }
    },
    TurnContext: class MockTurnContext {
      static getConversationReference(activity: any) {
        return {
          channelId: activity.channelId || "msteams",
          serviceUrl: activity.serviceUrl || "https://smba.trafficmanager.net/amer/",
          conversation: activity.conversation,
          bot: activity.recipient,
        };
      }
    },
    CardFactory: {
      adaptiveCard: vi.fn((card: any) => ({
        contentType: "application/vnd.microsoft.card.adaptive",
        content: card,
      })),
    },
    MessageFactory: {
      attachment: vi.fn((attachment: any) => ({
        type: "message",
        attachments: [attachment],
      })),
    },
  };
});

// Mock axios
vi.mock("axios", () => ({
  default: {
    get: vi.fn(),
    post: vi.fn(),
  },
}));

// Mock winston
vi.mock("winston", () => {
  const mockLogger = {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
  };
  return {
    default: {
      createLogger: vi.fn(() => mockLogger),
      format: {
        combine: vi.fn(),
        timestamp: vi.fn(),
        errors: vi.fn(),
        json: vi.fn(),
        colorize: vi.fn(),
        simple: vi.fn(),
      },
      transports: {
        File: class MockFileTransport {
          constructor() {}
        },
        Console: class MockConsoleTransport {
          constructor() {}
        },
      },
    },
  };
});

describe("config schema", () => {
  let registeredSchema: any;

  beforeEach(async () => {
    vi.clearAllMocks();
    lastAuthConfig = undefined;
    delete process.env.MSTEAMS_APP_ID;
    delete process.env.MSTEAMS_APP_PASSWORD;
    delete process.env.MSTEAMS_TENANT_ID;
  });

  afterEach(() => {
    vi.resetModules();
  });

  it("registers schema with expected fields", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({});
    mockCtx.registerConfigSchema = vi.fn((name: string, schema: any) => {
      registeredSchema = schema;
    });

    await plugin.init(mockCtx as any);

    expect(registeredSchema).toBeDefined();
    expect(registeredSchema.title).toBe("Microsoft Teams Integration");

    const fieldNames = registeredSchema.fields.map((f: any) => f.name);
    expect(fieldNames).toContain("appId");
    expect(fieldNames).toContain("appPassword");
    expect(fieldNames).toContain("tenantId");
    expect(fieldNames).toContain("webhookPort");
    expect(fieldNames).toContain("webhookPath");
    expect(fieldNames).toContain("dmPolicy");
    expect(fieldNames).toContain("requireMention");
    expect(fieldNames).toContain("replyStyle");
  });

  it("appId, appPassword, and tenantId are required fields", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({});
    mockCtx.registerConfigSchema = vi.fn((name: string, schema: any) => {
      registeredSchema = schema;
    });

    await plugin.init(mockCtx as any);

    const requiredFields = registeredSchema.fields.filter(
      (f: any) => f.required
    );
    const requiredNames = requiredFields.map((f: any) => f.name);
    expect(requiredNames).toContain("appId");
    expect(requiredNames).toContain("appPassword");
    expect(requiredNames).toContain("tenantId");
  });

  it("has correct default values for optional fields", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({});
    mockCtx.registerConfigSchema = vi.fn((name: string, schema: any) => {
      registeredSchema = schema;
    });

    await plugin.init(mockCtx as any);

    const fieldMap = new Map(
      registeredSchema.fields.map((f: any) => [f.name, f])
    );

    expect((fieldMap.get("webhookPort") as any).default).toBe(3978);
    expect((fieldMap.get("webhookPath") as any).default).toBe("/api/messages");
    expect((fieldMap.get("dmPolicy") as any).default).toBe("pairing");
    expect((fieldMap.get("requireMention") as any).default).toBe(true);
    expect((fieldMap.get("replyStyle") as any).default).toBe("thread");
  });

  it("config values take priority over env vars", async () => {
    process.env.MSTEAMS_APP_ID = "env-app-id";
    process.env.MSTEAMS_APP_PASSWORD = "env-password";
    process.env.MSTEAMS_TENANT_ID = "env-tenant-id";

    const mod = await import("../src/index.js");
    const plugin = mod.default;

    // Config has explicit values - these should be used over env
    const mockCtx = createMockContext({
      appId: "config-app-id",
      appPassword: "config-password",
      tenantId: "config-tenant-id",
    });

    await plugin.init(mockCtx as any);

    // ConfigurationBotFrameworkAuthentication should have been constructed
    // with config-supplied values, not env vars
    expect(lastAuthConfig).toBeDefined();
    expect(lastAuthConfig.MicrosoftAppId).toBe("config-app-id");
    expect(lastAuthConfig.MicrosoftAppPassword).toBe("config-password");
    expect(lastAuthConfig.MicrosoftAppTenantId).toBe("config-tenant-id");
  });

  it("falls back to env vars when config values missing", async () => {
    process.env.MSTEAMS_APP_ID = "env-app-id";
    process.env.MSTEAMS_APP_PASSWORD = "env-password";
    process.env.MSTEAMS_TENANT_ID = "env-tenant-id";

    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({});

    await plugin.init(mockCtx as any);

    // ConfigurationBotFrameworkAuthentication should have been constructed
    // with env var values since config values are missing
    expect(lastAuthConfig).toBeDefined();
    expect(lastAuthConfig.MicrosoftAppId).toBe("env-app-id");
    expect(lastAuthConfig.MicrosoftAppPassword).toBe("env-password");
    expect(lastAuthConfig.MicrosoftAppTenantId).toBe("env-tenant-id");
  });
});
