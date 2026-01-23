import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";

import { OfficeMailService } from "../../services/office-mail";
import { mockOfficeGlobal, createSuccessResult, createFailureResult } from "../mocks/office-mocks";

describe("OfficeMailService", () => {
  let service: OfficeMailService;

  beforeEach(() => {
    mockOfficeGlobal();
    service = new OfficeMailService();
  });

  afterEach(() => {
    vi.clearAllMocks();
  });

  describe("getSubject", () => {
    it("returns email subject on success", async () => {
      const mockSubject = "Test Email Subject";
      const subjectMock = vi.mocked(Office.context.mailbox.item!.subject.getAsync);
      subjectMock.mockImplementation((callback) => {
        callback(createSuccessResult(mockSubject));
      });

      const result = await service.getSubject();
      expect(result).toBe(mockSubject);
      expect(subjectMock).toHaveBeenCalledOnce();
    });

    it("rejects with error on failure", async () => {
      const subjectMock = vi.mocked(Office.context.mailbox.item!.subject.getAsync);
      subjectMock.mockImplementation((callback) => {
        callback(createFailureResult("Failed to retrieve subject"));
      });

      await expect(service.getSubject()).rejects.toThrow("Failed to retrieve subject");
    });
  });

  describe("getFrom", () => {
    it("returns sender email address", async () => {
      const mockFrom = {
        displayName: "John Doe",
        emailAddress: "john@example.com",
      } as Office.EmailAddressDetails;

      const fromMock = vi.mocked(Office.context.mailbox.item!.from.getAsync);
      fromMock.mockImplementation((callback) => {
        callback!(createSuccessResult(mockFrom));
      });

      const result = await service.getFrom();
      expect(result).toEqual({
        address: "john@example.com",
        name: "John Doe",
      });
    });
  });

  describe("getToRecipients", () => {
    it("returns array of recipients", async () => {
      const mockRecipients = [
        { displayName: "Alice", emailAddress: "alice@example.com" },
        { displayName: "Bob", emailAddress: "bob@example.com" },
      ] as Office.EmailAddressDetails[];

      const toMock = vi.mocked(Office.context.mailbox.item!.to.getAsync);
      toMock.mockImplementation((callback) => {
        callback(createSuccessResult(mockRecipients));
      });

      const result = await service.getToRecipients();
      expect(result).toHaveLength(2);
      expect(result[0]).toEqual({
        address: "alice@example.com",
        name: "Alice",
      });
    });
  });

  describe("getCcRecipients", () => {
    it("returns array of CC recipients", async () => {
      const mockRecipients = [
        { displayName: "Charlie", emailAddress: "charlie@example.com" },
      ] as Office.EmailAddressDetails[];

      const ccMock = vi.mocked(Office.context.mailbox.item!.cc.getAsync);
      ccMock.mockImplementation((callback) => {
        callback(createSuccessResult(mockRecipients));
      });

      const result = await service.getCcRecipients();
      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({
        address: "charlie@example.com",
        name: "Charlie",
      });
    });
  });

  describe("getRestItemId", () => {
    it("converts EWS ID to REST ID", async () => {
      const mockRestId = "rest-id-123";
      vi.mocked(Office.context.mailbox.convertToRestId).mockReturnValue(mockRestId);

      const result = await service.getRestItemId();
      expect(result).toBe(mockRestId);
      expect(Office.context.mailbox.convertToRestId).toHaveBeenCalledWith(
        "test-item-id",
        Office.MailboxEnums.RestVersion.v2_0
      );
    });

    it("returns itemId directly on iOS", async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (Office.context.mailbox.diagnostics as any).hostName = "OutlookIOS";
      const result = await service.getRestItemId();
      expect(result).toBe("test-item-id");
      expect(Office.context.mailbox.convertToRestId).not.toHaveBeenCalled();
    });
  });

  describe("showNotification", () => {
    it("adds notification message", async () => {
      const addAsyncMock = vi.mocked(Office.context.mailbox.item!.notificationMessages.addAsync);
      addAsyncMock.mockImplementation((_key, _options, callback) => {
        if (callback) {
          callback(createSuccessResult(undefined as unknown as void));
        }
      });

      await service.showNotification("Success message");
      expect(addAsyncMock).toHaveBeenCalledWith(
        "status",
        expect.objectContaining({
          message: "Success message",
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        }),
        expect.any(Function)
      );
    });
  });

  describe("getEmailItem", () => {
    it("returns complete email details", async () => {
      const mockSubject = "Test Subject";
      const mockFrom = {
        displayName: "Sender",
        emailAddress: "sender@example.com",
      } as Office.EmailAddressDetails;
      const mockTo = [
        { displayName: "To", emailAddress: "to@example.com" },
      ] as Office.EmailAddressDetails[];
      const mockCc = [
        { displayName: "Cc", emailAddress: "cc@example.com" },
      ] as Office.EmailAddressDetails[];

      const item = Office.context.mailbox.item!;
      vi.mocked(item.subject.getAsync).mockImplementation((cb) => {
        cb(createSuccessResult(mockSubject));
      });
      vi.mocked(item.from.getAsync).mockImplementation((cb) => {
        cb!(createSuccessResult(mockFrom));
      });
      vi.mocked(item.to.getAsync).mockImplementation((cb) => {
        cb(createSuccessResult(mockTo));
      });
      vi.mocked(item.cc.getAsync).mockImplementation((cb) => {
        cb(createSuccessResult(mockCc));
      });

      const emailItem = await service.getEmailItem();

      expect(emailItem.subject).toBe(mockSubject);
      expect(emailItem.from.address).toBe("sender@example.com");
      expect(emailItem.to).toHaveLength(1);
      expect(emailItem.cc).toHaveLength(1);
      expect(emailItem.itemId).toBe("test-item-id");
      expect(emailItem.conversationId).toBe("conv-id");
    });
  });
});
