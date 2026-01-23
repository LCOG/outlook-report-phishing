// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Constructs a local URL for the web page for the given path.
 * @param path The path to construct a local URL for.
 * @returns
 */
export function createLocalUrl(path: string): string {
  return `${window.location.origin}/${path}`;
}
