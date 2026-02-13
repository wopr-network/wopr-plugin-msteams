/**
 * Tests for Adaptive Card support (WOP-115)
 *
 * Tests:
 * - buildAdaptiveCard creates valid card structure
 * - Card with title, body, and actions
 * - Card with image
 * - Card with OpenUrl and Submit actions
 * - Card without optional fields
 * - Adaptive cards sent via sendResponse when enabled
 * - Plain text sent when adaptive cards disabled
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createMockContext } from "./mocks/wopr-context.js";

const mockCardFactoryAdaptiveCard = vi.fn((card: any) => ({
  contentType: "application/vnd.microsoft.card.adaptive",
  content: card,
}));

const mockMessageFactoryAttachment = vi.fn((attachment: any) => ({
  type: "message",
  attachments: [attachment],
}));

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
      constructor() {
        this.onTurnError = null;
      }
      process = mockProcess;
      continueConversationAsync = vi.fn();
    },
    ConfigurationBotFrameworkAuthentication: class MockAuth {
      constructor(config: any) {}
    },
    TurnContext: class MockTurnContext {
      static getConversationReference(activity: any) {
        return { conversation: activity.conversation, bot: activity.recipient };
      }
    },
    CardFactory: {
      adaptiveCard: mockCardFactoryAdaptiveCard,
    },
    MessageFactory: {
      attachment: mockMessageFactoryAttachment,
    },
  };
});

vi.mock("winston", () => {
  const mockLogger = { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn() };
  return {
    default: {
      createLogger: vi.fn(() => mockLogger),
      format: { combine: vi.fn(), timestamp: vi.fn(), errors: vi.fn(), json: vi.fn(), colorize: vi.fn(), simple: vi.fn() },
      transports: {
        File: class { constructor() {} },
        Console: class { constructor() {} },
      },
    },
  };
});

vi.mock("axios", () => ({
  default: { get: vi.fn(), post: vi.fn() },
}));

describe("adaptive cards", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    delete process.env.MSTEAMS_APP_ID;
    delete process.env.MSTEAMS_APP_PASSWORD;
    delete process.env.MSTEAMS_TENANT_ID;
  });

  afterEach(() => {
    vi.resetModules();
  });

  it("buildAdaptiveCard creates card with body only", async () => {
    const { buildAdaptiveCard } = await import("../src/index.js");
    buildAdaptiveCard({ body: "Hello world" });

    expect(mockCardFactoryAdaptiveCard).toHaveBeenCalledWith(
      expect.objectContaining({
        type: "AdaptiveCard",
        version: "1.4",
        body: expect.arrayContaining([
          expect.objectContaining({
            type: "TextBlock",
            text: "Hello world",
            wrap: true,
          }),
        ]),
      })
    );
  });

  it("buildAdaptiveCard includes title when provided", async () => {
    const { buildAdaptiveCard } = await import("../src/index.js");
    buildAdaptiveCard({ title: "Test Title", body: "Content" });

    const cardArg = mockCardFactoryAdaptiveCard.mock.calls[0][0];
    expect(cardArg.body[0]).toEqual(
      expect.objectContaining({
        type: "TextBlock",
        text: "Test Title",
        size: "Large",
        weight: "Bolder",
      })
    );
    expect(cardArg.body[1]).toEqual(
      expect.objectContaining({
        type: "TextBlock",
        text: "Content",
      })
    );
  });

  it("buildAdaptiveCard includes image when provided", async () => {
    const { buildAdaptiveCard } = await import("../src/index.js");
    buildAdaptiveCard({ body: "With image", imageUrl: "https://example.com/img.png" });

    const cardArg = mockCardFactoryAdaptiveCard.mock.calls[0][0];
    const imageBlock = cardArg.body.find((b: any) => b.type === "Image");
    expect(imageBlock).toEqual(
      expect.objectContaining({
        type: "Image",
        url: "https://example.com/img.png",
        size: "Auto",
      })
    );
  });

  it("buildAdaptiveCard includes OpenUrl actions", async () => {
    const { buildAdaptiveCard } = await import("../src/index.js");
    buildAdaptiveCard({
      body: "With action",
      actions: [
        { type: "Action.OpenUrl", title: "Visit", url: "https://example.com" },
      ],
    });

    const cardArg = mockCardFactoryAdaptiveCard.mock.calls[0][0];
    expect(cardArg.actions).toEqual([
      { type: "Action.OpenUrl", title: "Visit", url: "https://example.com" },
    ]);
  });

  it("buildAdaptiveCard includes Submit actions", async () => {
    const { buildAdaptiveCard } = await import("../src/index.js");
    buildAdaptiveCard({
      body: "With submit",
      actions: [
        { type: "Action.Submit", title: "Confirm", data: { action: "confirm" } },
      ],
    });

    const cardArg = mockCardFactoryAdaptiveCard.mock.calls[0][0];
    expect(cardArg.actions).toEqual([
      { type: "Action.Submit", title: "Confirm", data: { action: "confirm" } },
    ]);
  });

  it("buildAdaptiveCard omits actions when not provided", async () => {
    const { buildAdaptiveCard } = await import("../src/index.js");
    buildAdaptiveCard({ body: "No actions" });

    const cardArg = mockCardFactoryAdaptiveCard.mock.calls[0][0];
    expect(cardArg.actions).toBeUndefined();
  });

  it("sends adaptive card response when useAdaptiveCards is true", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({
      appId: "test-id",
      appPassword: "test-pass",
      tenantId: "test-tenant",
      useAdaptiveCards: true,
      dmPolicy: "open",
    });

    await plugin.init(mockCtx as any);

    const activity = {
      type: "message",
      text: "Hi",
      from: { id: "user-1", name: "Alice" },
      recipient: { id: "bot-1", name: "Bot" },
      conversation: { id: "conv-1", conversationType: "personal", name: "Chat" },
    };

    await mod.handleWebhook(
      { __activity: activity },
      { status: vi.fn().mockReturnThis(), send: vi.fn() }
    );

    // Should have called CardFactory.adaptiveCard for the response
    expect(mockCardFactoryAdaptiveCard).toHaveBeenCalled();
    expect(mockMessageFactoryAttachment).toHaveBeenCalled();
  });

  it("sends plain text when useAdaptiveCards is false", async () => {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({
      appId: "test-id",
      appPassword: "test-pass",
      tenantId: "test-tenant",
      useAdaptiveCards: false,
      dmPolicy: "open",
    });

    await plugin.init(mockCtx as any);

    const activity = {
      type: "message",
      text: "Hi",
      from: { id: "user-1", name: "Alice" },
      recipient: { id: "bot-1", name: "Bot" },
      conversation: { id: "conv-1", conversationType: "personal", name: "Chat" },
    };

    await mod.handleWebhook(
      { __activity: activity },
      { status: vi.fn().mockReturnThis(), send: vi.fn() }
    );

    // sendActivity should be called with plain text (no CardFactory)
    expect(mockSendActivity).toHaveBeenCalled();
    const sentArg = mockSendActivity.mock.calls[0][0];
    expect(sentArg.text).toBeDefined();
    expect(sentArg.textFormat).toBe("markdown");
  });
});
