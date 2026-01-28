/* eslint-disable no-console */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides the default MSAL configuration for the add-in project.

import { LogLevel, type Configuration } from "@azure/msal-browser";

import { createLocalUrl } from "./util";

// Entra app registration client IDs for development and production
export const clientIdDev = "81fbcf38-230e-421d-b5fd-35ad67e8db7d";
export const clientIdProd = "7ad80a52-3622-4bbd-ae71-fc5344c525a3";
export const clientId = process.env.NODE_ENV === "production" ? clientIdProd : clientIdDev;

export const msalConfig: Configuration = {
  auth: {
    clientId,
    redirectUri: createLocalUrl("auth.html"),
    postLogoutRedirectUri: createLocalUrl("auth.html"),
  },
  cache: {
    cacheLocation: "localStorage",
  },
  system: {
    loggerOptions: {
      logLevel: LogLevel.Verbose,
      loggerCallback: (level: LogLevel, message: string) => {
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
        }
      },
      piiLoggingEnabled: true,
    },
  },
};

// Default scopes to use in the fallback dialog.
export const defaultScopes = ["user.read", "files.read"];
