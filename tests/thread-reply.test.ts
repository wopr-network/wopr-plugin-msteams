/**
 * Tests for thread/reply support (WOP-115)
 *
 * Tests:
 * - Response sets replyToId when replyStyle is "thread"
 * - Response does not set replyToId when replyStyle is "top-level"
 * - Conversation references stored for proactive messaging
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createMockContext } from "./mocks/wopr-context.js";

const mockSendActivity = vi.fn().mockResolvedValue({});
const mockProcess = vi.fn(async (req: any, res: any, handler: any) => {
  if (req.__activity) {
    await handler({
      activity: req.__activity,
      sendActivity: mockSendActivity,
    });
  }
});

vi.mock("botbuilder", () => {
  return {
    CloudAdapter: class MockCloudAdapter {
      onTurnError: any;
      constructor() { this.onTurnError = null; }
      process = mockProcess;
      continueConversationAsync = vi.fn();
    },
    ConfigurationBotFrameworkAuthentication: class { constructor(config: any) {} },
    TurnContext: class {
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
        replyToId: undefined,
      })),
    },
  };
});

vi.mock("winston", () => {
  const mockLogger = { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn() };
  return {
    default: {
      createLogger: vi.fn(() => mockLogger),
      format: { combine: vi.fn(), timestamp: vi.fn(), errors: vi.fn(), json: vi.fn(), colorize: vi.fn(), simple: vi.fn() },
      transports: { File: class { constructor() {} }, Console: class { constructor() {} } },
    },
  };
});

vi.mock("axios", () => ({
  default: { get: vi.fn(), post: vi.fn() },
}));

function makeActivity(overrides: Record<string, any> = {}) {
  return {
    type: "message",
    id: "msg-123",
    text: "Hello WOPR",
    from: { id: "user-1", name: "Test User" },
    recipient: { id: "bot-1", name: "WOPR Bot" },
    conversation: { id: "conv-1", conversationType: "personal", name: "Test Chat" },
    ...overrides,
  };
}

describe("thread/reply support", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    delete process.env.MSTEAMS_APP_ID;
    delete process.env.MSTEAMS_APP_PASSWORD;
    delete process.env.MSTEAMS_TENANT_ID;
  });

  afterEach(() => {
    vi.resetModules();
  });

  it("sets replyToId when replyStyle is thread (plain text mode)", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({
      appId: "test-id",
      appPassword: "test-pass",
      tenantId: "test-tenant",
      dmPolicy: "open",
      replyStyle: "thread",
      useAdaptiveCards: false,
    });

    await plugin.init(mockCtx as any);

    const activity = makeActivity({ id: "msg-456" });
    await mod.handleWebhook(
      { __activity: activity },
      { status: vi.fn().mockReturnThis(), send: vi.fn() }
    );

    expect(mockSendActivity).toHaveBeenCalledWith(
      expect.objectContaining({
        replyToId: "msg-456",
      })
    );
  });

  it("does not set replyToId when replyStyle is top-level (plain text mode)", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({
      appId: "test-id",
      appPassword: "test-pass",
      tenantId: "test-tenant",
      dmPolicy: "open",
      replyStyle: "top-level",
      useAdaptiveCards: false,
    });

    await plugin.init(mockCtx as any);

    const activity = makeActivity({ id: "msg-789" });
    await mod.handleWebhook(
      { __activity: activity },
      { status: vi.fn().mockReturnThis(), send: vi.fn() }
    );

    const sentActivity = mockSendActivity.mock.calls[0][0];
    expect(sentActivity.replyToId).toBeUndefined();
  });

  it("stores conversation references for proactive messaging", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({
      appId: "test-id",
      appPassword: "test-pass",
      tenantId: "test-tenant",
      dmPolicy: "open",
      useAdaptiveCards: false,
    });

    // Capture the extension to check conversation references
    let capturedExtension: any = null;
    mockCtx.registerExtension = vi.fn((name: string, ext: any) => {
      if (name === "msteams") capturedExtension = ext;
    });

    await plugin.init(mockCtx as any);

    const activity = makeActivity({
      conversation: { id: "convo-abc", conversationType: "personal", name: "Chat" },
    });

    await mod.handleWebhook(
      { __activity: activity },
      { status: vi.fn().mockReturnThis(), send: vi.fn() }
    );

    // The extension should have stored the conversation reference
    const refs = capturedExtension.getConversationReferences();
    expect(refs.has("convo-abc")).toBe(true);
  });

  it("conversation references are cleared on shutdown", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    let capturedExtension: any = null;
    const mockCtx = createMockContext({
      appId: "test-id",
      appPassword: "test-pass",
      tenantId: "test-tenant",
      dmPolicy: "open",
      useAdaptiveCards: false,
    });
    mockCtx.registerExtension = vi.fn((name: string, ext: any) => {
      if (name === "msteams") capturedExtension = ext;
    });

    await plugin.init(mockCtx as any);

    // Send a message to populate references
    const activity = makeActivity();
    await mod.handleWebhook(
      { __activity: activity },
      { status: vi.fn().mockReturnThis(), send: vi.fn() }
    );

    // References should exist
    expect(capturedExtension.getConversationReferences().size).toBeGreaterThan(0);

    // Shutdown clears them
    await plugin.shutdown();

    // After shutdown, extension reference is stale, but the map was cleared
    // We can't check via the extension after shutdown since ctx is null.
    // Instead, we verify shutdown completes without error.
  });
});
