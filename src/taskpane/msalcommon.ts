// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides common MSAL functions for use in the add-in project.

import {
  createNestablePublicClientApplication,
  type IPublicClientApplication,
  type RedirectRequest,
  type SilentRequest,
  type PopupRequest,
} from "@azure/msal-browser";

import { msalConfig } from "./msalconfig";

/**
 * Gets a token request for a given account context.
 * @param accountContext The account context to get the token request for.
 * @returns The token request.
 */
export function getTokenRequest(
  scopes: string[],
  selectAccount: boolean,
  redirectUri?: string
): RedirectRequest & SilentRequest & PopupRequest {
  const request: RedirectRequest & SilentRequest & PopupRequest = {
    scopes,
  };

  if (selectAccount) {
    request.prompt = "select_account";
  }

  if (redirectUri !== undefined && redirectUri !== "") {
    request.redirectUri = redirectUri;
  }

  return request;
}

let _publicClientApp: IPublicClientApplication | undefined;

/**
 * Returns the existing public client application. Returns a new public client application if it did not exist.
 * @returns The nested public client application.
 */
export async function ensurePublicClient(): Promise<IPublicClientApplication> {
  _publicClientApp ??= await createNestablePublicClientApplication(msalConfig);
  return _publicClientApp;
}
