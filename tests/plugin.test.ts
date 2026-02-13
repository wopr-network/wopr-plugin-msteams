/**
 * Tests for MS Teams plugin registration and initialization (WOP-243)
 *
 * Tests:
 * - Plugin exports and shape
 * - Plugin init with valid credentials
 * - Plugin init without credentials (warns gracefully)
 * - Plugin shutdown cleans up state
 * - Config schema registration
 * - Identity refresh on init
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createMockContext } from "./mocks/wopr-context.js";

// Track CloudAdapter instances and calls
let mockAdapterInstance: any;
const mockProcess = vi.fn();

// Mock botbuilder before importing the plugin
vi.mock("botbuilder", () => {
  return {
    CloudAdapter: class MockCloudAdapter {
      onTurnError: any;
      constructor() {
        mockAdapterInstance = this;
        this.onTurnError = null;
      }
      process = mockProcess;
      continueConversationAsync = vi.fn();
    },
    ConfigurationBotFrameworkAuthentication: class MockAuth {
      config: any;
      constructor(config: any) {
        this.config = config;
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

describe("MS Teams plugin", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockAdapterInstance = null;
    // Clear env vars
    delete process.env.MSTEAMS_APP_ID;
    delete process.env.MSTEAMS_APP_PASSWORD;
    delete process.env.MSTEAMS_TENANT_ID;
  });

  afterEach(async () => {
    // Reset module to clear module-level state between tests
    vi.resetModules();
  });

  it("exports a valid WOPRPlugin object", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    expect(plugin).toBeDefined();
    expect(plugin.name).toBe("msteams");
    expect(plugin.version).toBe("1.0.0");
    expect(plugin.description).toContain("Microsoft Teams");
    expect(typeof plugin.init).toBe("function");
    expect(typeof plugin.shutdown).toBe("function");
  });

  it("exports handleWebhook function", async () => {
    const mod = await import("../src/index.js");
    expect(typeof mod.handleWebhook).toBe("function");
  });

  it("init with valid config registers schema and creates adapter", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({
      appId: "test-app-id",
      appPassword: "test-password",
      tenantId: "test-tenant-id",
    });

    await plugin.init(mockCtx as any);

    expect(mockCtx.registerConfigSchema).toHaveBeenCalledWith(
      "msteams",
      expect.objectContaining({
        title: "Microsoft Teams Integration",
        fields: expect.any(Array),
      })
    );
    expect(mockAdapterInstance).not.toBeNull();
  });

  it("init without credentials warns but does not throw", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({});

    // Should not throw
    await expect(plugin.init(mockCtx as any)).resolves.toBeUndefined();

    // Schema should still be registered
    expect(mockCtx.registerConfigSchema).toHaveBeenCalled();
  });

  it("init with env var credentials creates adapter", async () => {
    process.env.MSTEAMS_APP_ID = "env-app-id";
    process.env.MSTEAMS_APP_PASSWORD = "env-password";
    process.env.MSTEAMS_TENANT_ID = "env-tenant-id";

    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({});

    await plugin.init(mockCtx as any);

    expect(mockAdapterInstance).not.toBeNull();
  });

  it("init refreshes agent identity", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({
      appId: "test-app-id",
      appPassword: "test-password",
      tenantId: "test-tenant-id",
    });

    await plugin.init(mockCtx as any);

    expect(mockCtx.getAgentIdentity).toHaveBeenCalled();
  });

  it("shutdown cleans up adapter and context", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({
      appId: "test-app-id",
      appPassword: "test-password",
      tenantId: "test-tenant-id",
    });

    await plugin.init(mockCtx as any);
    expect(mockAdapterInstance).not.toBeNull();

    await plugin.shutdown();

    // After shutdown, handleWebhook should fail gracefully
    const mockRes = {
      status: vi.fn().mockReturnThis(),
      send: vi.fn(),
    };
    await mod.handleWebhook({}, mockRes);
    expect(mockRes.status).toHaveBeenCalledWith(500);
  });
});
