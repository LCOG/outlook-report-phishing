import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";

import { PhishingReportService } from "../../services/phishing-report";

import type { GraphClient } from "../../services/graph-client";
import type { Message, User } from "@microsoft/microsoft-graph-types";

// Test helper functions
function createMockUser(overrides: Partial<User> = {}): User {
  return {
    mail: "test@example.com",
    displayName: "Test User",
    id: "user-123",
    ...overrides,
  };
}

function createMockMessage(overrides: Partial<Message> = {}): Message {
  return {
    id: "msg-123",
    subject: "Test Subject",
    body: {
      content: "Test content",
    },
    ...overrides,
  };
}

interface MockGraphClient {
  getMessage: ReturnType<typeof vi.fn>;
  forwardMessage: ReturnType<typeof vi.fn>;
  moveMessage: ReturnType<typeof vi.fn>;
}

function createMockGraphClient(): MockGraphClient {
  return {
    getMessage: vi.fn(),
    forwardMessage: vi.fn(),
    moveMessage: vi.fn(),
  };
}

describe("PhishingReportService", () => {
  let service: PhishingReportService;
  let mockGraphClient: MockGraphClient;

  const TEST_API_URL = "https://test-api.example.com/report";

  beforeEach(() => {
    mockGraphClient = createMockGraphClient();

    service = new PhishingReportService(mockGraphClient as unknown as GraphClient, TEST_API_URL);

    global.fetch = vi.fn();
  });

  afterEach(() => {
    vi.clearAllMocks();
  });

  describe("constructor", () => {
    it("throws error when API URL is not provided", () => {
      expect(() => {
        new PhishingReportService(mockGraphClient as unknown as GraphClient, "");
      }).toThrow("Report API URL is required");
    });

    it("creates service successfully with valid parameters", () => {
      expect(service).toBeInstanceOf(PhishingReportService);
    });
  });

  describe("buildForwardComment", () => {
    it("returns basic comment without additional info", () => {
      const result = service.buildForwardComment("John Doe", "phishing");

      expect(result).toBe(
        "John Doe forwarded a suspicious email (phishing) via the Report Phish add-in."
      );
    });

    it("includes additional info when provided", () => {
      const result = service.buildForwardComment("Jane Smith", "spam", "This looks like a scam");

      expect(result).toContain("Jane Smith forwarded a suspicious email (spam)");
      expect(result).toContain("Additional details: This looks like a scam");
    });

    it("ignores empty additional info", () => {
      const result = service.buildForwardComment("John Doe", "phishing", "   ");

      expect(result).not.toContain("Additional details");
      expect(result).toBe(
        "John Doe forwarded a suspicious email (phishing) via the Report Phish add-in."
      );
    });

    it("ignores undefined additional info", () => {
      const result = service.buildForwardComment("John Doe", "phishing", undefined);

      expect(result).not.toContain("Additional details");
      expect(result).toBe(
        "John Doe forwarded a suspicious email (phishing) via the Report Phish add-in."
      );
    });

    it("handles different report types", () => {
      const phishing = service.buildForwardComment("User", "phishing");
      const spam = service.buildForwardComment("User", "spam");
      const malware = service.buildForwardComment("User", "malware");

      expect(phishing).toContain("(phishing)");
      expect(spam).toContain("(spam)");
      expect(malware).toContain("(malware)");
    });

    it("formats multiline additional info correctly", () => {
      const result = service.buildForwardComment("User", "phishing", "Line 1\nLine 2");

      expect(result).toContain("\n\nAdditional details: Line 1\nLine 2");
    });
  });

  describe("logPhishingReport", () => {
    const mockMessage = createMockMessage({
      body: { content: "Test email content" },
    });

    it("sends correct data to backend API", async () => {
      vi.mocked(global.fetch).mockResolvedValue({
        ok: true,
        status: 200,
      } as Response);

      await service.logPhishingReport("user@example.com", mockMessage);

      expect(global.fetch).toHaveBeenCalledWith(
        TEST_API_URL,
        expect.objectContaining({
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            employee_email: "user@example.com",
            email_message: "Test email content",
          }),
        })
      );
    });

    it("converts email to lowercase", async () => {
      vi.mocked(global.fetch).mockResolvedValue({
        ok: true,
        status: 200,
      } as Response);

      await service.logPhishingReport("USER@EXAMPLE.COM", mockMessage);

      const callBody = JSON.parse(vi.mocked(global.fetch).mock.calls[0][1]?.body as string);
      expect(callBody.employee_email).toBe("user@example.com");
    });

    it("handles missing message body", async () => {
      vi.mocked(global.fetch).mockResolvedValue({
        ok: true,
        status: 200,
      } as Response);

      const messageWithoutBody: Message = {
        subject: "No body",
      };

      await service.logPhishingReport("user@example.com", messageWithoutBody);

      const callBody = JSON.parse(vi.mocked(global.fetch).mock.calls[0][1]?.body as string);
      expect(callBody.email_message).toBe("No content");
    });

    it("handles missing message body content", async () => {
      vi.mocked(global.fetch).mockResolvedValue({
        ok: true,
        status: 200,
      } as Response);

      const messageWithEmptyBody: Message = {
        subject: "Empty body",
        body: {},
      };

      await service.logPhishingReport("user@example.com", messageWithEmptyBody);

      const callBody = JSON.parse(vi.mocked(global.fetch).mock.calls[0][1]?.body as string);
      expect(callBody.email_message).toBe("No content");
    });

    it("does not throw error when API call fails", async () => {
      vi.mocked(global.fetch).mockResolvedValue({
        ok: false,
        status: 500,
        statusText: "Internal Server Error",
      } as Response);

      await expect(
        service.logPhishingReport("user@example.com", mockMessage)
      ).resolves.not.toThrow();
    });

    it("does not throw error when fetch fails", async () => {
      vi.mocked(global.fetch).mockRejectedValue(new Error("Network error"));

      await expect(
        service.logPhishingReport("user@example.com", mockMessage)
      ).resolves.not.toThrow();
    });

    it("logs error to console when API call fails", async () => {
      const consoleSpy = vi.spyOn(console, "error").mockImplementation(() => {});
      vi.mocked(global.fetch).mockResolvedValue({
        ok: false,
        status: 500,
        statusText: "Internal Server Error",
      } as Response);

      await service.logPhishingReport("user@example.com", mockMessage);

      expect(consoleSpy).toHaveBeenCalledWith(
        "Failed to log phishing report to backend:",
        expect.any(Error)
      );
      consoleSpy.mockRestore();
    });
  });

  describe("reportPhishing", () => {
    const mockUser = createMockUser({
      mail: "reporter@example.com",
      displayName: "Test Reporter",
    });

    const mockMessage = createMockMessage({
      id: "msg-123",
      subject: "Phishing Email",
      body: { content: "Suspicious content" },
    });

    const reportOptions = {
      messageId: "msg-123",
      user: mockUser,
      reportType: "phishing",
      forwardToEmail: "security@example.com",
    };

    beforeEach(() => {
      mockGraphClient.getMessage.mockResolvedValue(mockMessage);
      mockGraphClient.forwardMessage.mockResolvedValue(undefined);
      vi.mocked(global.fetch).mockResolvedValue({
        ok: true,
        status: 200,
      } as Response);
    });

    it("successfully reports phishing email", async () => {
      const result = await service.reportPhishing(reportOptions);

      expect(result.success).toBe(true);
      expect(result.error).toBeUndefined();
    });

    it("calls getMessage with correct parameters", async () => {
      await service.reportPhishing(reportOptions);

      expect(mockGraphClient.getMessage).toHaveBeenCalledWith("msg-123", "?$select=subject,body");
    });

    it("calls logPhishingReport with correct data", async () => {
      await service.reportPhishing(reportOptions);

      expect(global.fetch).toHaveBeenCalledWith(
        TEST_API_URL,
        expect.objectContaining({
          method: "POST",
          body: expect.stringContaining("reporter@example.com"),
        })
      );
    });

    it("forwards message with correct structure", async () => {
      await service.reportPhishing(reportOptions);

      expect(mockGraphClient.forwardMessage).toHaveBeenCalledWith(
        "msg-123",
        expect.objectContaining({
          comment: expect.stringContaining("Test Reporter"),
          toRecipients: [
            {
              emailAddress: {
                name: "Service Desk",
                address: "security@example.com",
              },
            },
          ],
        })
      );
    });

    it("includes report type in forward comment", async () => {
      await service.reportPhishing({
        ...reportOptions,
        reportType: "spam",
      });

      const forwardCall = mockGraphClient.forwardMessage.mock.calls[0][1];
      expect(forwardCall.comment).toContain("(spam)");
    });

    it("includes additional info in forward comment when provided", async () => {
      await service.reportPhishing({
        ...reportOptions,
        additionalInfo: "Looks like a scam",
      });

      const forwardCall = mockGraphClient.forwardMessage.mock.calls[0][1];
      expect(forwardCall.comment).toContain("Additional details: Looks like a scam");
    });

    it("handles user without display name", async () => {
      const userWithoutName: User = {
        mail: "user@example.com",
        id: "user-123",
      };

      await service.reportPhishing({
        ...reportOptions,
        user: userWithoutName,
      });

      const forwardCall = mockGraphClient.forwardMessage.mock.calls[0][1];
      expect(forwardCall.comment).toContain("User forwarded");
    });

    it("handles user without email", async () => {
      const userWithoutEmail: User = {
        displayName: "Test User",
        id: "user-123",
      };

      await service.reportPhishing({
        ...reportOptions,
        user: userWithoutEmail,
      });

      const fetchCall = vi.mocked(global.fetch).mock.calls[0];
      const fetchBody = JSON.parse(fetchCall[1]?.body as string);
      expect(fetchBody.employee_email).toBe("");
    });

    it("returns error result when getMessage fails", async () => {
      mockGraphClient.getMessage.mockRejectedValue(new Error("Graph API error"));

      const result = await service.reportPhishing(reportOptions);

      expect(result.success).toBe(false);
      expect(result.error).toBe("Graph API error");
    });

    it("returns error result when forwardMessage fails", async () => {
      mockGraphClient.forwardMessage.mockRejectedValue(new Error("Failed to forward"));

      const result = await service.reportPhishing(reportOptions);

      expect(result.success).toBe(false);
      expect(result.error).toBe("Failed to forward");
    });

    it("continues even if logging to backend fails", async () => {
      vi.mocked(global.fetch).mockRejectedValue(new Error("Backend down"));

      const result = await service.reportPhishing(reportOptions);

      expect(result.success).toBe(true);
      expect(mockGraphClient.forwardMessage).toHaveBeenCalled();
    });

    it("handles non-Error exceptions", async () => {
      mockGraphClient.getMessage.mockRejectedValue("String error");

      const result = await service.reportPhishing(reportOptions);

      expect(result.success).toBe(false);
      expect(result.error).toBe("Unknown error occurred");
    });

    it("logs error to console when reporting fails", async () => {
      const consoleSpy = vi.spyOn(console, "error").mockImplementation(() => {});
      mockGraphClient.getMessage.mockRejectedValue(new Error("API failed"));

      await service.reportPhishing(reportOptions);

      expect(consoleSpy).toHaveBeenCalledWith("Failed to report phishing email:", "API failed");
      consoleSpy.mockRestore();
    });

    it("successfully reports email with all details", async () => {
      const fullUser = createMockUser({
        mail: "reporter@company.com",
        displayName: "Jane Reporter",
      });
      const fullMessage = createMockMessage({
        subject: "Suspicious Email",
        body: { content: "Click this link!" },
      });

      mockGraphClient.getMessage.mockResolvedValue(fullMessage);

      const options = {
        messageId: "msg-456",
        user: fullUser,
        reportType: "phishing",
        additionalInfo: "Looks like credential harvesting",
        forwardToEmail: "security@company.com",
      };

      const result = await service.reportPhishing(options);

      expect(result.success).toBe(true);
      expect(result.error).toBeUndefined();

      expect(mockGraphClient.getMessage).toHaveBeenCalledOnce();
      expect(mockGraphClient.getMessage).toHaveBeenCalledWith("msg-456", "?$select=subject,body");

      expect(mockGraphClient.forwardMessage).toHaveBeenCalledOnce();
      const forwardBody = mockGraphClient.forwardMessage.mock.calls[0][1];
      expect(forwardBody.comment).toContain("Jane Reporter");
      expect(forwardBody.comment).toContain("(phishing)");
      expect(forwardBody.comment).toContain("credential harvesting");
      expect(forwardBody.toRecipients[0].emailAddress.address).toBe("security@company.com");

      expect(global.fetch).toHaveBeenCalledOnce();
      const fetchCall = vi.mocked(global.fetch).mock.calls[0];
      expect(fetchCall[0]).toBe(TEST_API_URL);
      const fetchBody = JSON.parse(fetchCall[1]?.body as string);
      expect(fetchBody.employee_email).toBe("reporter@company.com");
      expect(fetchBody.email_message).toBe("Click this link!");
    });
  });

  describe("moveToJunk", () => {
    it("calls moveMessage with correct parameters", async () => {
      mockGraphClient.moveMessage.mockResolvedValue({} as Message);

      await service.moveToJunk("msg-123");

      expect(mockGraphClient.moveMessage).toHaveBeenCalledWith("msg-123", "junkemail");
    });

    it("throws error when moveMessage fails", async () => {
      mockGraphClient.moveMessage.mockRejectedValue(new Error("Move failed"));

      await expect(service.moveToJunk("msg-123")).rejects.toThrow("Move failed");
    });

    it("returns successfully when message is moved", async () => {
      const movedMessage = createMockMessage({ id: "msg-123" });
      mockGraphClient.moveMessage.mockResolvedValue(movedMessage);

      await expect(service.moveToJunk("msg-123")).resolves.not.toThrow();
    });
  });
});
