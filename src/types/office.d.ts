/**
 * Extended type definitions for Office.js missing or incomplete in @types/office-js
 */

declare namespace Office {
  interface MessageRead {
    /**
     * Gets the attachments of the message.
     * @param callback Optional. The method to call when the operation is complete.
     */
    getAttachmentsAsync(callback?: (result: AsyncResult<AttachmentDetailsCompose[]>) => void): void;
    getAttachmentsAsync(
      options: Office.AsyncContextOptions,
      callback?: (result: AsyncResult<AttachmentDetailsCompose[]>) => void
    ): void;
  }
}
