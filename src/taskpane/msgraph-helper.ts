// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides the provides functionality to get Microsoft Graph data.

/* global console fetch HeadersInit BodyInit Headers */

/**
 *  Calls a Microsoft Graph API and returns the response.
 *
 * @param accessToken The access token to use for the request.
 * @param path Path component of the URI, e.g., "/me". Should start with "/".
 * @param queryParams Query parameters, e.g., "?$select=name,id". Should start with "?". Optional.
 * @param additionalHeaders Any additional headers to use in the request. Optional.
 * @param body The body element to include in the request. Optional.
 * @returns
 */
export async function makeGraphRequest({
  accessToken,
  path,
  queryParams,
  additionalHeaders,
  body,
}: {
  accessToken: string;
  path: string;
  queryParams?: string;
  additionalHeaders?: HeadersInit;
  body?: BodyInit;
}): Promise<any> {
  if (!path) throw new Error("path is required.");
  if (!path.startsWith("/")) throw new Error("path must start with '/'.");
  if (queryParams && !queryParams.startsWith("?"))
    throw new Error("queryParams must start with '?'.");

  const headers = new Headers(additionalHeaders);
  headers.set("Authorization", accessToken);

  const response = await fetch(`https://graph.microsoft.com/v1.0${path}${queryParams ?? ""}`, {
    method: "GET",
    headers: headers,
    body: body,
  });

  if (response.ok) {
    const data = await response.json();
    console.log("printing data from msgraph-helper:");
    console.log(data);
    return data;
  } else {
    throw new Error(response.statusText);
  }
}

export async function makePostGraphRequest({
  accessToken,
  path,
  additionalHeaders,
  body,
}: {
  accessToken: string;
  path: string;
  additionalHeaders?: HeadersInit;
  body?: BodyInit;
}): Promise<any> {
  if (!path) throw new Error("path is required.");
  if (!path.startsWith("/")) throw new Error("path must start with '/'.");

  const headers = new Headers(additionalHeaders);
  headers.set("Authorization", accessToken);

  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: "POST",
    headers: headers,
    body: body,
  });

  return response;
}
