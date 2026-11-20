/* eslint-disable no-console */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { LogLevel, type Configuration } from "@azure/msal-browser";

import { createLocalUrl } from "./util";

// Entra app registration client IDs
export const clientIdDev = "81fbcf38-230e-421d-b5fd-35ad67e8db7d";
export const clientIdStag = "7e82c950-e246-4863-b2f3-3f5f2579af7c";
export const clientIdProdLcog = "7ad80a52-3622-4bbd-ae71-fc5344c525a3";
export const clientIdProdCG = "ac3385f6-ec21-4037-bae1-bd1270b8dfb3";

type AgencyKey = "dev" | "stag" | "prod-lcog" | "prod-cg";

const hostnameToAgency: Record<string, AgencyKey> = {
  localhost: "dev",
  "127.0.0.1": "dev",
  "freebee-existing-timid.ngrok-free.dev": "dev",
  "reportphish-staging.lcog-or.gov": "stag",
  "reportphish.lcog-or.gov": "prod-lcog",
  "cg-reportphish.lcog-or.gov": "prod-cg",
};

const clientIdsByAgency: Record<AgencyKey, string> = {
  dev: clientIdDev,
  stag: clientIdStag,
  "prod-lcog": clientIdProdLcog,
  "prod-cg": clientIdProdCG,
};

function resolveAgencyKey(): AgencyKey {
  const host = window.location.hostname.toLowerCase();
  const hostMatch = hostnameToAgency[host];
  if (hostMatch) return hostMatch;

  if (process.env.NODE_ENV === "staging") return "stag";
  if (process.env.NODE_ENV !== "production") return "dev";

  console.warn(`Unknown hostname "${host}" in production; defaulting to prod-lcog`);
  return "prod-lcog";
}

function buildMsalConfig(clientId: string): Configuration {
  return {
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
}

let msalConfigPromise: Promise<Configuration> | undefined;

export async function getMsalConfig(): Promise<Configuration> {
  msalConfigPromise ??= Promise.resolve().then(() => {
    const agency = resolveAgencyKey();
    const clientId = clientIdsByAgency[agency];

    if (!clientId) {
      throw new Error(`No clientId configured for agency "${agency}"`);
    }

    console.log("MSAL agency:", agency);
    console.log("Using client ID:", clientId);

    return buildMsalConfig(clientId);
  });

  return msalConfigPromise;
}

// Default scopes to use in the fallback dialog.
export const defaultScopes = ["user.read", "files.read"];
