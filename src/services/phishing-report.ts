import type { GraphClient } from "./graph-client";
import type { Message, User } from "@microsoft/microsoft-graph-types";

/**
 * Options for reporting a phishing email
 */
export interface ReportPhishingOptions {
  messageId: string;
  user: User;
  reportType: string;
  additionalInfo?: string;
  forwardToEmail: string;
}

/**
 * Result of a phishing report operation
 */
export interface PhishingReportResult {
  success: boolean;
  error?: string;
}

/**
 * Service for handling phishing email reports
 * Encapsulates the business logic for reporting and managing suspicious emails:
 * - construct a comment to attach to the forwarded suspicious email based on the
 *   user's category selection and optional text input
 * - forward the reported email
 * - log the report in Team App
 * - move the reported email to junk
 */
export class PhishingReportService {
  constructor(
    private readonly graphClient: GraphClient,
    private readonly reportApiUrl: string
  ) {
    if (!reportApiUrl) {
      throw new Error("Report API URL is required");
    }
  }

  /**
   * Build the comment text for a forwarded phishing report
   * @param displayName - Name of the user reporting the email
   * @param reportType - Type of report (phishing, spam, etc.)
   * @param additionalInfo - Optional additional details from the user
   * @returns Formatted comment string
   */
  buildForwardComment(displayName: string, reportType: string, additionalInfo?: string): string {
    const baseComment = `${displayName} forwarded a suspicious email (${reportType}) via the Report Phish add-in.`;

    if (additionalInfo && additionalInfo.trim() !== "") {
      return `${baseComment}\n\nAdditional details: ${additionalInfo}`;
    }

    return baseComment;
  }

  /**
   * Log a phishing report to the Team App backend API
   * @param employeeEmail - Email address of the employee reporting
   * @param message - The message being reported
   */
  async logPhishingReport(employeeEmail: string, message: Message): Promise<void> {
    try {
      const response = await fetch(this.reportApiUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          employee_email: employeeEmail.toLowerCase(),
          email_message: message.body?.content ?? "No content",
        }),
      });

      if (!response.ok) {
        throw new Error(`Failed to log report: ${response.status} ${response.statusText}`);
      }
    } catch (error) {
      // Log but don't fail the entire operation if backend logging fails
      console.error("Failed to log phishing report to backend:", error);
    }
  }

  /**
   * Report a phishing email by forwarding it to the service desk
   * @param options - Configuration for the phishing report
   * @returns Result indicating success or failure with error details
   */
  async reportPhishing(options: ReportPhishingOptions): Promise<PhishingReportResult> {
    try {
      // Get the original message content
      const message = await this.graphClient.getMessage(options.messageId, "?$select=subject,body");

      // Log to Team App API (non-blocking - won't fail the operation)
      await this.logPhishingReport(options.user.mail ?? "", message);

      // Build the forward comment
      const forwardComment = this.buildForwardComment(
        options.user.displayName ?? "User",
        options.reportType,
        options.additionalInfo
      );

      // Construct the forward request body
      const forwardBody = {
        comment: forwardComment,
        toRecipients: [
          {
            emailAddress: {
              name: "Service Desk",
              address: options.forwardToEmail,
            },
          },
        ],
      };

      // Forward the message to service desk
      await this.graphClient.forwardMessage(options.messageId, forwardBody);

      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
      console.error("Failed to report phishing email:", errorMessage);

      return {
        success: false,
        error: errorMessage,
      };
    }
  }

  /**
   * Move a message to the Junk Email folder
   * @param messageId - ID of the message to move
   */
  async moveToJunk(messageId: string): Promise<void> {
    await this.graphClient.moveMessage(messageId, "junkemail");
  }
}
