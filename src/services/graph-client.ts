import type { User, Message } from "@microsoft/microsoft-graph-types";

/**
 * Interface for Graph API error responses.
 */
export interface GraphError {
  error: {
    code: string;
    message: string;
    innerError?: {
      code: string;
      message: string;
      requestId?: string;
      date?: string;
    };
  };
}

/**
 * Interface for Graph API paginated responses.
 */
export interface GraphResponse<T> {
  value?: T[];
  "@odata.nextLink"?: string;
  "@odata.context"?: string;
}

/**
 * A type-safe wrapper for Microsoft Graph API requests.
 */
export class GraphClient {
  private readonly baseUrl = "https://graph.microsoft.com/v1.0";

  constructor(private readonly accessToken: string) {
    if (!accessToken) {
      throw new Error("Access token is required for GraphClient");
    }
  }

  /**
   * Generic request method for Graph API.
   * @param endpoint The API endpoint (e.g., "/me").
   * @param options Fetch options.
   * @returns The typed response data.
   */
  private async request<T>(endpoint: string, options: RequestInit = {}): Promise<T> {
    const url = `${this.baseUrl}${endpoint.startsWith("/") ? endpoint : `/${endpoint}`}`;

    const headers = new Headers(options.headers);
    if (!headers.has("Authorization")) {
      headers.set("Authorization", `Bearer ${this.accessToken}`);
    }

    const response = await fetch(url, {
      ...options,
      headers,
    });

    if (!response.ok) {
      let errorData: GraphError;
      try {
        errorData = await response.json();
      } catch {
        throw new Error(`Graph API error: ${response.status} ${response.statusText}`);
      }
      throw new Error(`Graph API error (${errorData.error.code}): ${errorData.error.message}`);
    }

    // Some Graph API calls (like /forward or /move) might return 202 Accepted or 204 No Content with no body
    if (response.status === 204 || response.headers.get("Content-Length") === "0") {
      return {} as T;
    }

    return response.json() as Promise<T>;
  }

  /**
   * GET request to Graph API.
   */
  async get<T>(endpoint: string, options: RequestInit = {}): Promise<T> {
    return this.request<T>(endpoint, { ...options, method: "GET" });
  }

  /**
   * POST request to Graph API.
   */
  async post<T>(endpoint: string, body: unknown, options: RequestInit = {}): Promise<T> {
    return this.request<T>(endpoint, {
      ...options,
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...options.headers,
      },
      body: JSON.stringify(body),
    });
  }

  /**
   * Gets the current user's profile.
   */
  async getUser(): Promise<User> {
    return this.get<User>("/me");
  }

  /**
   * Gets a specific message by ID.
   */
  async getMessage(id: string, queryParams: string = ""): Promise<Message> {
    return this.get<Message>(`/me/messages/${id}${queryParams}`);
  }

  /**
   * Forwards a message.
   */
  async forwardMessage(id: string, forwardData: unknown): Promise<void> {
    await this.post<void>(`/me/messages/${id}/forward`, forwardData);
  }

  /**
   * Moves a message to another folder.
   */
  async moveMessage(id: string, destinationId: string): Promise<Message> {
    return this.post<Message>(`/me/messages/${id}/move`, { destinationId });
  }
}
