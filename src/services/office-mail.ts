import { isAsyncResultSuccess, hasMailboxContext } from "../utils/office-type-guards";

export interface EmailAddress {
  address: string;
  name: string;
}

export interface EmailItem {
  itemId: string;
  subject: string;
  from: EmailAddress;
  to: EmailAddress[];
  cc: EmailAddress[];
  sender: EmailAddress;
  conversationId: string;
  dateTimeCreated: Date;
  dateTimeModified: Date;
  internetMessageId: string;
}

export interface AttachmentDetails {
  id: string;
  name: string;
  size: number;
  attachmentType: "file" | "item" | "cloud";
  contentType: string;
  isInline: boolean;
}

/**
 * Service to interact with Office.js Mailbox APIs using Promises and strong typing
 */
export class OfficeMailService {
  /**
   * Safe access to Office.context.mailbox
   */
  private get mailbox(): Office.Mailbox {
    if (!hasMailboxContext()) {
      throw new Error("Office Mailbox context is not available");
    }
    return Office.context.mailbox;
  }

  /**
   * Safe access to the current mail item
   */
  private get item(): Office.MessageRead {
    const item = this.mailbox.item;
    if (!item) {
      throw new Error("No mail item is available in the current context");
    }
    return item as unknown as Office.MessageRead;
  }

  /**
   * Get the current email's subject
   */
  async getSubject(): Promise<string> {
    return new Promise((resolve, reject) => {
      // Use cast to any to access getAsync which might be missing in some type versions
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const subject = this.item.subject as any;
      if (typeof subject.getAsync === "function") {
        subject.getAsync((result: Office.AsyncResult<string>) => {
          if (isAsyncResultSuccess(result)) {
            resolve(result.value);
          } else {
            reject(new Error(`Failed to get subject: ${result.error.message}`));
          }
        });
      } else {
        resolve(String(subject));
      }
    });
  }

  /**
   * Get the sender's email address and display name
   */
  async getFrom(): Promise<EmailAddress> {
    return new Promise((resolve, reject) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const from = this.item.from as any;
      if (typeof from.getAsync === "function") {
        from.getAsync((result: Office.AsyncResult<Office.EmailAddressDetails>) => {
          if (isAsyncResultSuccess(result)) {
            resolve({
              address: result.value.emailAddress,
              name: result.value.displayName,
            });
          } else {
            reject(new Error(`Failed to get From field: ${result.error.message}`));
          }
        });
      } else {
        resolve({
          address: String(from.emailAddress ?? ""),
          name: String(from.displayName ?? ""),
        });
      }
    });
  }

  /**
   * Get the "To" recipients
   */
  async getToRecipients(): Promise<EmailAddress[]> {
    return new Promise((resolve, reject) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const to = this.item.to as any;
      if (typeof to.getAsync === "function") {
        to.getAsync((result: Office.AsyncResult<Office.EmailAddressDetails[]>) => {
          if (isAsyncResultSuccess(result)) {
            resolve(
              result.value.map((rec) => ({
                address: rec.emailAddress,
                name: rec.displayName,
              }))
            );
          } else {
            reject(new Error(`Failed to get To recipients: ${result.error.message}`));
          }
        });
      } else {
        const recipients = Array.isArray(to) ? to : [];
        resolve(
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          recipients.map((rec: any) => ({
            address: String(rec.emailAddress ?? ""),
            name: String(rec.displayName ?? ""),
          }))
        );
      }
    });
  }

  /**
   * Get the "CC" recipients
   */
  async getCcRecipients(): Promise<EmailAddress[]> {
    return new Promise((resolve, reject) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const cc = this.item.cc as any;
      if (typeof cc.getAsync === "function") {
        cc.getAsync((result: Office.AsyncResult<Office.EmailAddressDetails[]>) => {
          if (isAsyncResultSuccess(result)) {
            resolve(
              result.value.map((rec) => ({
                address: rec.emailAddress,
                name: rec.displayName,
              }))
            );
          } else {
            reject(new Error(`Failed to get Cc recipients: ${result.error.message}`));
          }
        });
      } else {
        const recipients = Array.isArray(cc) ? cc : [];
        resolve(
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          recipients.map((rec: any) => ({
            address: String(rec.emailAddress ?? ""),
            name: String(rec.displayName ?? ""),
          }))
        );
      }
    });
  }

  /**
   * Get the current item's ID in Office.js format
   */
  getItemId(): string {
    return this.item.itemId;
  }

  /**
   * Get the REST-formatted ID for use with Microsoft Graph API
   */
  async getRestItemId(): Promise<string> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const diagnostics = this.mailbox.diagnostics as any;
    if (diagnostics?.hostName === "OutlookIOS") {
      // In Outlook on iOS, itemId is already in REST format
      return this.item.itemId;
    }

    return new Promise((resolve, reject) => {
      const ewsId = this.item.itemId;
      try {
        const restId = this.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
        resolve(restId);
      } catch (error) {
        reject(
          new Error(
            `Failed to convert Item ID to REST format: ${error instanceof Error ? error.message : String(error)}`
          )
        );
      }
    });
  }

  /**
   * Get all attachments for the current item
   */
  async getAttachments(): Promise<AttachmentDetails[]> {
    return new Promise((resolve, reject) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const item = this.item as any;
      if (typeof item.getAttachmentsAsync !== "function") {
        reject(new Error("getAttachmentsAsync is not supported on this item"));
        return;
      }

      item.getAttachmentsAsync((result: Office.AsyncResult<unknown[]>) => {
        if (isAsyncResultSuccess(result)) {
          resolve(
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            result.value.map((attr: any) => ({
              id: String(attr.id),
              name: String(attr.name),
              size: Number(attr.size),
              attachmentType: attr.attachmentType as "file" | "item" | "cloud",
              contentType: String(attr.contentType ?? "application/octet-stream"),
              isInline: Boolean(attr.isInline),
            }))
          );
        } else {
          reject(new Error(`Failed to get attachments: ${result.error.message}`));
        }
      });
    });
  }

  /**
   * Display a dialog using the Office Dialog API
   */
  async displayDialog(
    url: string,
    options: Office.DialogOptions = { height: 60, width: 30 }
  ): Promise<Office.Dialog> {
    return new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(url, options, (result) => {
        if (isAsyncResultSuccess(result)) {
          resolve(result.value);
        } else {
          reject(new Error(`Failed to display dialog: ${result.error.message}`));
        }
      });
    });
  }

  /**
   * Get internet headers by name
   */
  async getInternetHeaders(names: string[]): Promise<Record<string, string>> {
    return new Promise((resolve, reject) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const item = this.item as any;
      const headers = item.internetHeaders;
      // eslint-disable-next-line @typescript-eslint/strict-boolean-expressions
      if (!headers || typeof headers.getAsync !== "function") {
        reject(new Error("Internet headers are not available on this item"));
        return;
      }

      headers.getAsync(names, (result: Office.AsyncResult<Record<string, string>>) => {
        if (isAsyncResultSuccess(result)) {
          resolve(result.value);
        } else {
          reject(new Error(`Failed to get internet headers: ${result.error.message}`));
        }
      });
    });
  }

  /**
   * Show a notification message to the user
   */
  async showNotification(message: string, key: string = "status"): Promise<void> {
    return new Promise((resolve, reject) => {
      this.item.notificationMessages.addAsync(
        key,
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: message,
          icon: "icon-16",
          persistent: false,
        },
        (result) => {
          if (isAsyncResultSuccess(result)) {
            resolve();
          } else {
            reject(new Error(`Failed to show notification: ${result.error.message}`));
          }
        }
      );
    });
  }

  /**
   * Get a consolidated EmailItem object with all common properties
   */
  async getEmailItem(): Promise<EmailItem> {
    const [subject, from, to, cc] = await Promise.all([
      this.getSubject(),
      this.getFrom(),
      this.getToRecipients(),
      this.getCcRecipients(),
    ]);

    return {
      itemId: this.item.itemId,
      subject,
      from,
      to,
      cc,
      sender: from, // In most read cases, from and sender are the same for our purposes
      conversationId: this.item.conversationId,
      dateTimeCreated: this.item.dateTimeCreated,
      dateTimeModified: this.item.dateTimeModified,
      internetMessageId: this.item.internetMessageId,
    };
  }
}
