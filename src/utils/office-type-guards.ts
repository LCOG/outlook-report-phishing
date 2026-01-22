/**
 * Utility functions to safely interact with Office.js APIs
 */

/**
 * Check if the Office object is available and initialized
 */
export function isOfficeInitialized(): boolean {
  return typeof Office !== "undefined";
}

/**
 * Check if the mailbox context is available
 */
export function hasMailboxContext(): boolean {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return isOfficeInitialized() && (Office as any).context?.mailbox !== undefined;
}

/**
 * Check if a mail item is currently selected/available
 */
export function hasMailItem(): boolean {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return hasMailboxContext() && (Office as any).context?.mailbox?.item !== undefined;
}

/**
 * Ensure Office is ready before performing operations
 * @returns Promise that resolves when Office is ready
 */
export async function ensureOfficeReady(): Promise<void> {
  if (typeof Office === "undefined") {
    throw new Error("Office.js library is not loaded");
  }
  await Office.onReady();
}

/**
 * Type guard for Office.AsyncResult success status
 */
export function isAsyncResultSuccess<T>(
  result: Office.AsyncResult<T>
): result is Office.AsyncResult<T> & { status: Office.AsyncResultStatus.Succeeded } {
  return result.status === Office.AsyncResultStatus.Succeeded;
}
