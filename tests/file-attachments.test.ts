/**
 * Tests for file attachment handling (WOP-115)
 *
 * Tests:
 * - downloadAttachment fetches file from contentUrl
 * - downloadAttachment returns null for missing attachment
 * - downloadAttachment returns null when no contentUrl
 * - buildFileCard creates valid file info card
 * - File attachments in messages are logged
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createMockContext } from "./mocks/wopr-context.js";

const mockAxiosGet = vi.fn();
const mockAxiosPost = vi.fn();

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
        return { conversation: activity.conversation, bot: activity.recipient };
      }
    },
    CardFactory: {
      adaptiveCard: vi.fn((card: any) => ({
        contentType: "application/vnd.microsoft.card.adaptive",
        content: card,
      })),
    },
    MessageFactory: {
      attachment: vi.fn((a: any) => ({ type: "message", attachments: [a] })),
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
  default: {
    get: (...args: any[]) => mockAxiosGet(...args),
    post: (...args: any[]) => mockAxiosPost(...args),
  },
}));

describe("file attachments", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    delete process.env.MSTEAMS_APP_ID;
    delete process.env.MSTEAMS_APP_PASSWORD;
    delete process.env.MSTEAMS_TENANT_ID;
  });

  afterEach(() => {
    vi.resetModules();
  });

  describe("downloadAttachment", () => {
    it("downloads file from contentUrl", async () => {
      const { downloadAttachment } = await import("../src/index.js");

      mockAxiosGet.mockResolvedValue({
        data: Buffer.from("file content"),
      });

      const activity = {
        attachments: [
          {
            contentUrl: "https://example.com/files/report.pdf",
            name: "report.pdf",
            contentType: "application/pdf",
          },
        ],
      };

      const result = await downloadAttachment(activity as any);

      expect(result).not.toBeNull();
      expect(result!.filename).toBe("report.pdf");
      expect(result!.contentType).toBe("application/pdf");
      expect(result!.content).toEqual(Buffer.from("file content"));
    });

    it("returns null when no attachments", async () => {
      const { downloadAttachment } = await import("../src/index.js");

      const activity = { attachments: undefined };
      const result = await downloadAttachment(activity as any);
      expect(result).toBeNull();
    });

    it("returns null when attachment index out of range", async () => {
      const { downloadAttachment } = await import("../src/index.js");

      const activity = {
        attachments: [
          { contentUrl: "https://example.com/file.txt", name: "file.txt" },
        ],
      };

      const result = await downloadAttachment(activity as any, 5);
      expect(result).toBeNull();
    });

    it("returns null when attachment has no contentUrl", async () => {
      const { downloadAttachment } = await import("../src/index.js");

      const activity = {
        attachments: [
          { name: "file.txt", contentType: "text/plain" },
        ],
      };

      const result = await downloadAttachment(activity as any);
      expect(result).toBeNull();
    });

    it("uses default filename when name missing", async () => {
      const { downloadAttachment } = await import("../src/index.js");

      mockAxiosGet.mockResolvedValue({
        data: Buffer.from("data"),
      });

      const activity = {
        attachments: [
          { contentUrl: "https://example.com/files/unnamed" },
        ],
      };

      const result = await downloadAttachment(activity as any);
      expect(result!.filename).toBe("attachment");
      expect(result!.contentType).toBe("application/octet-stream");
    });

    it("downloads from specific attachment index", async () => {
      const { downloadAttachment } = await import("../src/index.js");

      mockAxiosGet.mockResolvedValue({
        data: Buffer.from("second file"),
      });

      const activity = {
        attachments: [
          { contentUrl: "https://example.com/first.txt", name: "first.txt" },
          { contentUrl: "https://example.com/second.txt", name: "second.txt", contentType: "text/plain" },
        ],
      };

      const result = await downloadAttachment(activity as any, 1);
      expect(result!.filename).toBe("second.txt");
      expect(mockAxiosGet).toHaveBeenCalledWith(
        "https://example.com/second.txt",
        expect.any(Object)
      );
    });
  });

  describe("buildFileCard", () => {
    it("creates file info card with all fields", async () => {
      const { buildFileCard } = await import("../src/index.js");

      const card = buildFileCard("report.pdf", "https://example.com/report.pdf", 1024);

      expect(card.contentType).toBe("application/vnd.microsoft.teams.card.file.info");
      expect(card.name).toBe("report.pdf");
      expect(card.contentUrl).toBe("https://example.com/report.pdf");
      expect(card.content.fileType).toBe("pdf");
      expect(card.content.fileSize).toBe(1024);
    });

    it("creates file info card without file size", async () => {
      const { buildFileCard } = await import("../src/index.js");

      const card = buildFileCard("image.png", "https://example.com/image.png");

      expect(card.name).toBe("image.png");
      expect(card.content.fileType).toBe("png");
      expect(card.content.fileSize).toBeUndefined();
    });

    it("handles files without extension", async () => {
      const { buildFileCard } = await import("../src/index.js");

      const card = buildFileCard("README", "https://example.com/README");

      expect(card.name).toBe("README");
      expect(card.content.fileType).toBe("README");
    });
  });

  describe("file attachments in messages", () => {
    it("processes messages with file attachments without error", async () => {
      const mod = await import("../src/index.js");
      const plugin = mod.default;

      const mockCtx = createMockContext({
        appId: "test-id",
        appPassword: "test-pass",
        tenantId: "test-tenant",
        dmPolicy: "open",
        useAdaptiveCards: false,
      });

      await plugin.init(mockCtx as any);

      const activity = {
        type: "message",
        text: "Here is a file",
        from: { id: "user-1", name: "Alice" },
        recipient: { id: "bot-1", name: "Bot" },
        conversation: { id: "conv-1", conversationType: "personal", name: "Chat" },
        attachments: [
          {
            contentUrl: "https://example.com/doc.pdf",
            name: "doc.pdf",
            contentType: "application/pdf",
          },
        ],
      };

      await mod.handleWebhook(
        { __activity: activity },
        { status: vi.fn().mockReturnThis(), send: vi.fn() }
      );

      // Message with attachment should still be injected
      expect(mockCtx.inject).toHaveBeenCalled();
    });

    it("filters out card-type attachments from file logging", async () => {
      const mod = await import("../src/index.js");
      const plugin = mod.default;

      const mockCtx = createMockContext({
        appId: "test-id",
        appPassword: "test-pass",
        tenantId: "test-tenant",
        dmPolicy: "open",
        useAdaptiveCards: false,
      });

      await plugin.init(mockCtx as any);

      const activity = {
        type: "message",
        text: "A card message",
        from: { id: "user-1", name: "Alice" },
        recipient: { id: "bot-1", name: "Bot" },
        conversation: { id: "conv-1", conversationType: "personal", name: "Chat" },
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: {},
          },
        ],
      };

      // Should not throw and should process normally
      await mod.handleWebhook(
        { __activity: activity },
        { status: vi.fn().mockReturnThis(), send: vi.fn() }
      );

      expect(mockCtx.inject).toHaveBeenCalled();
    });
  });
});
