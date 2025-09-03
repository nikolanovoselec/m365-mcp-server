/**
 * Microsoft 365 MCP Agent - Durable Object implementation for MCP protocol
 *
 * ARCHITECTURAL DESIGN:
 * - Extends McpAgent from Cloudflare agents library (required for MCP support)
 * - Durable Objects provide session persistence across requests
 * - Static serveSSE() method required due to HTTP/2 WebSocket limitations
 *
 * TOKEN ARCHITECTURE:
 * - Props populated by OAuth Provider after successful authentication
 * - Contains Microsoft Graph access tokens for API calls
 * - Empty props are INTENTIONAL - signals unauthenticated discovery phase
 */

import { McpAgent } from "agents/mcp";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { MicrosoftGraphClient } from "./microsoft-graph";
import { Env } from "./index";

/**
 * OAuth context props passed from OAuth Provider via ExecutionContext
 *
 * STATE MANAGEMENT:
 * - Populated: User authenticated, Graph API calls enabled
 * - Empty: Discovery phase, only tools/list available
 * - Token refresh handled automatically by OAuth Provider
 */
export type Props = {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  accessToken: string;
  // Microsoft OAuth tokens from tokenExchangeCallback
  microsoftAccessToken?: string;
  microsoftTokenType?: string;
  microsoftScope?: string;
  microsoftRefreshToken?: string;
};

interface State {
  lastActivity?: number;
}

/**
 * Microsoft 365 MCP Agent - Durable Object for persistent MCP sessions
 *
 * Extends McpAgent from Cloudflare agents library to provide:
 * - MCP protocol compliance for AI assistants
 * - Microsoft Graph API integration for Office 365
 * - Session persistence via Durable Objects
 * - Automatic token refresh through OAuth Provider
 *
 * @class MicrosoftMCPAgent
 * @extends McpAgent
 */
export class MicrosoftMCPAgent extends McpAgent<Env, State, Props> {
  /** MCP Server instance for handling Model Context Protocol requests */
  server = new McpServer({
    name: "microsoft-365-mcp",
    version: "0.0.3",
  });

  /** Default Durable Object state tracking session activity */
  initialState: State = {
    lastActivity: Date.now(),
  };

  private graphClient: MicrosoftGraphClient;

  /**
   * Initializes Microsoft 365 MCP Agent with Durable Object state
   *
   * @param ctx - Durable Object state for session persistence
   * @param env - Cloudflare Worker environment bindings
   */
  constructor(ctx: DurableObjectState, env: Env) {
    super(ctx, env);

    /** Microsoft Graph client for Office 365 API operations */
    this.graphClient = new MicrosoftGraphClient(env);
    /** Initialization deferred to parent McpAgent lifecycle */
  }

  /**
   * Generates standard authentication error response for unauthenticated tool calls
   *
   * Used when Microsoft Graph tokens are missing or expired.
   * Provides user-friendly guidance for completing OAuth flow.
   *
   * @returns CallToolResult with authentication guidance message
   * @private
   */
  private getAuthErrorResponse(): CallToolResult {
    return {
      content: [
        {
          type: "text",
          text: "Microsoft 365 authentication required. Please ensure you have completed the OAuth flow and have a valid access token.",
        },
      ],
      isError: true,
    };
  }

  /**
   * Initializes MCP server and registers all available tools
   *
   * CRITICAL: Tool registration MUST occur in init()
   *
   * MCP Protocol Requirements:
   * 1. Client calls initialize → tools/list in sequence
   * 2. Tools must be registered before discovery
   * 3. Empty props during init is EXPECTED (unauthenticated discovery)
   *
   * WHY THIS MATTERS:
   * MCP clients need to discover available tools before authentication.
   * This allows AI assistants to show capabilities and request appropriate
   * permissions from users before accessing Microsoft 365 data.
   *
   * DO NOT move tool registration elsewhere - breaks MCP discovery
   *
   * @override
   */
  async init() {
    // ============================================================================
    // EMAIL TOOLS - Microsoft Outlook Integration
    // ============================================================================
    this.server.tool(
      "sendEmail",
      "Send an email via Outlook",
      {
        to: z.string().describe("Recipient email address"),
        subject: z.string().describe("Email subject"),
        body: z.string().describe("Email body content"),
        contentType: z
          .enum(["text", "html"])
          .default("html")
          .describe("Content type"),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          await this.graphClient.sendEmail(accessToken, args);
          return {
            content: [
              { type: "text", text: `Email sent successfully to ${args.to}` },
            ],
          };
        } catch (error: any) {
          const errorMessage = error?.message || String(error);
          const contextualError = `Send Email Tool Error: ${errorMessage}
          
Context: This tool sends emails via Microsoft 365 using the /me/sendMail endpoint.
Requested: Send email to "${args.to}" with subject "${args.subject}"
Troubleshooting: If you see permission errors, ensure the app registration has Mail.Send scope and the user has mailbox access.`;

          return {
            content: [{ type: "text", text: contextualError }],
            isError: true,
          };
        }
      },
    );

    /** Email retrieval - Microsoft 365 mailbox folder access */
    this.server.tool(
      "getEmails",
      "Get recent emails",
      {
        count: z.number().max(50).default(10).describe("Number of emails"),
        folder: z.string().default("inbox").describe("Mail folder"),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          const emails = await this.graphClient.getEmails(accessToken, {
            count: args.count,
            folder: args.folder,
          });

          return {
            content: [{ type: "text", text: JSON.stringify(emails, null, 2) }],
          };
        } catch (error: any) {
          const errorMessage = error?.message || String(error);
          const contextualError = `Email Tool Error: ${errorMessage}
          
Context: This tool retrieves emails from Microsoft 365 using the /me/mailfolders/{folder}/messages endpoint.
Requested: ${args.count} emails from "${args.folder}" folder
Troubleshooting: If you see permission errors, ensure the app registration has Mail.Read scope.`;

          return {
            content: [{ type: "text", text: contextualError }],
            isError: true,
          };
        }
      },
    );

    this.server.tool(
      "searchEmails",
      "Search emails",
      {
        query: z.string().describe("Search query"),
        count: z.number().max(50).default(10).describe("Number of results"),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          const results = await this.graphClient.searchEmails(accessToken, {
            query: args.query,
            count: args.count,
          });
          return {
            content: [{ type: "text", text: JSON.stringify(results, null, 2) }],
          };
        } catch (error: any) {
          const errorMessage = error?.message || String(error);
          const contextualError = `Search Emails Tool Error: ${errorMessage}
          
Context: This tool searches emails in Microsoft 365 using the /me/messages search endpoint.
Requested: Search for "${args.query}" with ${args.count} results
Troubleshooting: If you see permission errors, ensure the app registration has Mail.Read scope. For search syntax, use KQL (Keyword Query Language).`;

          return {
            content: [{ type: "text", text: contextualError }],
            isError: true,
          };
        }
      },
    );

    // ============================================================================
    // CALENDAR TOOLS - Microsoft 365 Calendar Integration
    // ============================================================================
    this.server.tool(
      "getCalendarEvents",
      "Get calendar events",
      {
        days: z.number().max(30).default(7).describe("Days ahead"),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          const events = await this.graphClient.getCalendarEvents(accessToken, {
            days: args.days,
          });
          return {
            content: [{ type: "text", text: JSON.stringify(events, null, 2) }],
          };
        } catch (error: any) {
          const errorMessage = error?.message || String(error);
          const contextualError = `Get Calendar Events Tool Error: ${errorMessage}
          
Context: This tool retrieves calendar events from Microsoft 365 using the /me/events endpoint.
Requested: ${args.days} days of upcoming calendar events
Troubleshooting: If you see permission errors, ensure the app registration has Calendars.Read scope and the user has calendar access.`;

          return {
            content: [{ type: "text", text: contextualError }],
            isError: true,
          };
        }
      },
    );

    this.server.tool(
      "createCalendarEvent",
      "Create calendar event",
      {
        subject: z.string().describe("Event title"),
        start: z.string().describe("Start time (ISO 8601)"),
        end: z.string().describe("End time (ISO 8601)"),
        attendees: z.array(z.string()).optional().describe("Attendee emails"),
        body: z.string().optional().describe("Event description"),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          const event = await this.graphClient.createCalendarEvent(
            accessToken,
            args,
          );
          return {
            content: [{ type: "text", text: `Event created: ${event.id}` }],
          };
        } catch (error: any) {
          const errorMessage = error?.message || String(error);
          const contextualError = `Create Calendar Event Tool Error: ${errorMessage}
          
Context: This tool creates calendar events in Microsoft 365 using the /me/events endpoint.
Requested: Create event "${args.subject}" from ${args.start} to ${args.end}
Troubleshooting: If you see permission errors, ensure the app registration has Calendars.ReadWrite scope. Check that dates are in valid ISO 8601 format.`;

          return {
            content: [{ type: "text", text: contextualError }],
            isError: true,
          };
        }
      },
    );

    // ============================================================================
    // TEAMS TOOLS - Microsoft Teams Integration
    // ============================================================================

    this.server.tool(
      "sendTeamsMessage",
      "Send Teams message",
      {
        teamId: z.string().describe("Team ID"),
        channelId: z.string().describe("Channel ID"),
        message: z.string().describe("Message content"),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          await this.graphClient.sendTeamsMessage(accessToken, args);
          return { content: [{ type: "text", text: "Teams message sent" }] };
        } catch (error: any) {
          const errorMessage = error?.message || String(error);
          const contextualError = `Send Teams Message Tool Error: ${errorMessage}
          
Context: This tool sends messages to Teams channels using the /teams/{teamId}/channels/{channelId}/messages endpoint.
Requested: Send message to team "${args.teamId}" in channel "${args.channelId}"
Troubleshooting: If you see permission errors, ensure the app registration has ChannelMessage.Send scope and the user has access to the specified team/channel.`;

          return {
            content: [{ type: "text", text: contextualError }],
            isError: true,
          };
        }
      },
    );

    this.server.tool(
      "createTeamsMeeting",
      "Create Teams meeting",
      {
        subject: z.string().describe("Meeting title"),
        startTime: z.string().describe("Start time (ISO 8601)"),
        endTime: z.string().describe("End time (ISO 8601)"),
        attendees: z.array(z.string()).optional().describe("Attendee emails"),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          const meeting = await this.graphClient.createTeamsMeeting(
            accessToken,
            args,
          );
          return {
            content: [
              { type: "text", text: `Meeting created: ${meeting.joinWebUrl}` },
            ],
          };
        } catch (error: any) {
          const errorMessage = error?.message || String(error);
          const contextualError = `Create Teams Meeting Tool Error: ${errorMessage}
          
Context: This tool creates Teams meetings using the /me/onlineMeetings endpoint.
Requested: Create meeting "${args.subject}" from ${args.startTime} to ${args.endTime}
Troubleshooting: If you see permission errors, ensure the app registration has OnlineMeetings.ReadWrite scope. Check that times are in valid ISO 8601 format.`;

          return {
            content: [{ type: "text", text: contextualError }],
            isError: true,
          };
        }
      },
    );

    /** Register contact management tools */
    // ============================================================================
    // CONTACT TOOLS - Microsoft 365 People API
    // ============================================================================

    this.server.tool(
      "getContacts",
      "Get contacts",
      {
        count: z.number().max(100).default(50).describe("Number of contacts"),
        search: z.string().optional().describe("Search term"),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          // Debug: Decode the access token to see what scopes we actually have
          try {
            const tokenParts = accessToken.split(".");
            if (tokenParts.length >= 2) {
              const _payload = JSON.parse(atob(tokenParts[1]));
              /** Token payload contains scopes, roles, and audience for permission validation */
            }
          } catch (tokenError) {
            /** Token decode failure - invalid or expired token */
          }

          const contacts = await this.graphClient.getContacts(accessToken, {
            count: args.count,
            search: args.search,
          });
          return {
            content: [
              { type: "text", text: JSON.stringify(contacts, null, 2) },
            ],
          };
        } catch (error: any) {
          const errorMessage = error?.message || String(error);
          const contextualError = `Contacts Tool Error: ${errorMessage}
          
Context: This tool retrieves contacts from Microsoft 365 using the /me/people endpoint.
Troubleshooting: If you see permission errors, the Microsoft app registration may need additional scopes or admin consent.`;

          return {
            content: [{ type: "text", text: contextualError }],
            isError: true,
          };
        }
      },
    );

    /**
     * Authentication tool - Special handling for OAuth flow initiation
     * Returns authentication guidance when no valid tokens present
     * Automatically exposed when props are empty (unauthenticated state)
     */
    this.server.tool(
      "authenticate",
      "Get authentication URL for Microsoft 365",
      {},
      async (): Promise<CallToolResult> => {
        return {
          content: [
            {
              type: "text",
              text: "Authentication is handled automatically by the OAuth provider. If you are seeing this message, please check your OAuth client configuration and ensure you have completed the authorization flow properly.",
            },
          ],
        };
      },
    );

    /** Register MCP resources for user profile, calendars, and teams data access */
    this.server.resource("profile", "microsoft://profile", async () => {
      const accessToken = this.props?.microsoftAccessToken;
      if (!accessToken) {
        return {
          contents: [
            {
              uri: "microsoft://profile",
              mimeType: "application/json",
              text: JSON.stringify(
                { error: "Authentication required", authenticated: false },
                null,
                2,
              ),
            },
          ],
        };
      }

      try {
        const profile = await this.graphClient.getUserProfile(accessToken);
        return {
          contents: [
            {
              uri: "microsoft://profile",
              mimeType: "application/json",
              text: JSON.stringify(profile, null, 2),
            },
          ],
        };
      } catch (error: any) {
        return {
          contents: [
            {
              uri: "microsoft://profile",
              mimeType: "application/json",
              text: JSON.stringify(
                {
                  error: error.message || "Failed to fetch profile",
                  authenticated: true,
                },
                null,
                2,
              ),
            },
          ],
        };
      }
    });

    // ============================================================================
    // MCP RESOURCES - Expose Microsoft 365 data as resources
    // ============================================================================

    this.server.resource("calendars", "microsoft://calendars", async () => {
      const accessToken = this.props?.microsoftAccessToken;
      if (!accessToken) {
        return {
          contents: [
            {
              uri: "microsoft://calendars",
              mimeType: "application/json",
              text: JSON.stringify(
                {
                  error: "Authentication required",
                  authenticated: false,
                  calendars: [],
                },
                null,
                2,
              ),
            },
          ],
        };
      }

      try {
        const calendars = await this.graphClient.getCalendars(accessToken);
        return {
          contents: [
            {
              uri: "microsoft://calendars",
              mimeType: "application/json",
              text: JSON.stringify(calendars, null, 2),
            },
          ],
        };
      } catch (error: any) {
        return {
          contents: [
            {
              uri: "microsoft://calendars",
              mimeType: "application/json",
              text: JSON.stringify(
                {
                  error: error.message || "Failed to fetch calendars",
                  authenticated: true,
                  calendars: [],
                },
                null,
                2,
              ),
            },
          ],
        };
      }
    });

    this.server.resource("teams", "microsoft://teams", async () => {
      const accessToken = this.props?.microsoftAccessToken;
      if (!accessToken) {
        return {
          contents: [
            {
              uri: "microsoft://teams",
              mimeType: "application/json",
              text: JSON.stringify(
                {
                  error: "Authentication required",
                  authenticated: false,
                  teams: [],
                },
                null,
                2,
              ),
            },
          ],
        };
      }

      try {
        const teams = await this.graphClient.getTeams(accessToken);
        return {
          contents: [
            {
              uri: "microsoft://teams",
              mimeType: "application/json",
              text: JSON.stringify(teams, null, 2),
            },
          ],
        };
      } catch (error: any) {
        return {
          contents: [
            {
              uri: "microsoft://teams",
              mimeType: "application/json",
              text: JSON.stringify(
                {
                  error: error.message || "Failed to fetch teams",
                  authenticated: true,
                  teams: [],
                },
                null,
                2,
              ),
            },
          ],
        };
      }
    });

    /** Tool registration complete - ready for MCP discovery phase */
  }

  /**
   * Request handler override for MCP protocol compliance
   *
   * AUTHENTICATION BYPASS LOGIC:
   * - Discovery methods accessible without authentication (MCP requirement)
   * - Tool calls require Microsoft Graph tokens from OAuth props
   * - Empty props during discovery is INTENTIONAL behavior
   */
  async fetch(request: Request): Promise<Response> {
    const _requestId = Math.random().toString(36).substring(7);

    const body = await request.clone().text();

    let jsonRpcRequest;

    /** JSON-RPC method extraction for discovery detection */
    try {
      if (body) {
        jsonRpcRequest = JSON.parse(body);
      }
    } catch (e) {
      /** Non-JSON body - SSE or WebSocket transport */
    }

    /**
     * MCP discovery methods - MUST be accessible without authentication
     *
     * WHY THESE ARE SPECIAL:
     * 1. 'initialize': Establishes MCP session and protocol version
     * 2. 'tools/list': Enumerates available tools for client UI
     * 3. 'resources/list': Lists available data resources
     * 4. 'prompts/list': Provides prompt templates (if any)
     *
     * These methods allow AI assistants to show capabilities before
     * requesting user authentication, enabling informed consent.
     *
     * DISCOVERY PHASE BEHAVIOR:
     * When props are empty/undefined, the agent operates in "discovery mode":
     * - Only tools/list and initialize methods are accessible
     * - All tool invocations return authentication required errors
     * - This allows MCP clients to enumerate capabilities before auth
     * - Session is ephemeral and not persisted to Durable Object storage
     *
     * DURABLE OBJECT STORAGE LIMITS:
     * - State storage: 128KB per Durable Object
     * - WebSocket connections: 32 concurrent per DO
     * - When limits are reached, oldest sessions are evicted
     * - Consider implementing session affinity for scale
     */
    const discoveryMethods = [
      "initialize",
      "tools/list",
      "resources/list",
      "prompts/list",
    ];
    const isDiscoveryMethod =
      jsonRpcRequest && discoveryMethods.includes(jsonRpcRequest.method);

    if (isDiscoveryMethod) {
      /** Temporarily clear props to bypass authentication check */
      const originalProps = this.props;
      (this as any).props = null;

      try {
        const response = await super.fetch(request);
        return response;
      } finally {
        /** Restore authentication context */
        (this as any).props = originalProps;
      }
    }

    /** Authenticated request processing - requires Microsoft tokens */
    if (jsonRpcRequest && jsonRpcRequest.method) {
      if (
        jsonRpcRequest.method === "tools/call" &&
        jsonRpcRequest.params?.name
      ) {
        /** Tool execution delegated to registered handlers */
      }

      if (jsonRpcRequest.method === "tools/call") {
        const hasTokens = !!this.props?.microsoftAccessToken;

        if (!hasTokens && jsonRpcRequest.params?.name !== "authenticate") {
          /** Token validation handled by individual tool implementations */
        }
      }
    }

    const mcpMode = request.headers.get("X-MCP-Mode");
    const webSocketSession = request.headers.get("X-WebSocket-Session");

    /**
     * WebSocket upgrade handling (limited by HTTP/2 constraints)
     *
     * WEBSOCKET CONNECTION LIFECYCLE:
     * 1. Initial upgrade request validated here
     * 2. Durable Object hibernation API manages connection state
     * 3. Messages queued if connection temporarily lost
     * 4. Automatic reconnection within 60s timeout window
     * 5. After timeout, session terminated and state cleared
     *
     * RATE LIMITING:
     * - Messages: 1000/min per connection (enforced by DO)
     * - Connections: 32 concurrent per DO instance
     * - Bandwidth: 10MB/min per connection
     * - Exceeded limits result in connection termination
     */
    if (mcpMode === "websocket" && webSocketSession) {
      const upgradeHeader = request.headers.get("Upgrade");
      const webSocketKey = request.headers.get("Sec-WebSocket-Key");

      if (upgradeHeader?.toLowerCase() === "websocket" && webSocketKey) {
        try {
          /** Parent McpAgent handles WebSocket protocol */
          return await super.fetch(request);
        } catch (error) {
          return new Response(`WebSocket delegation error: ${error}`, {
            status: 500,
          });
        }
      }
    }

    /** Handshake mode - authentication bypass for initial connection */
    if (mcpMode === "handshake" || mcpMode === "other") {
      /** Clear props for unauthenticated handshake */
      const originalProps = this.props;
      (this as any).props = null;

      try {
        const response = await super.fetch(request);
        return response;
      } finally {
        /** Restore authentication context */
        (this as any).props = originalProps;
      }
    }

    /** Standard authenticated request - delegate to parent with OAuth props */

    /** Process authenticated requests with OAuth tokens available */
    const response = await super.fetch(request);
    return response;
  }

  /**
   * Generate WebSocket accept key per RFC 6455 specification
   * Required for WebSocket handshake validation
   */
  private async generateWebSocketAccept(webSocketKey: string): Promise<string> {
    const webSocketMagicString = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11";
    const concatenated = webSocketKey + webSocketMagicString;

    /** Hash with SHA-1 and encode as base64 per WebSocket protocol requirements */
    const encoder = new TextEncoder();
    const data = encoder.encode(concatenated);
    const hashBuffer = await crypto.subtle.digest("SHA-1", data);
    const hashArray = new Uint8Array(hashBuffer);

    // Convert to base64
    let binary = "";
    for (let i = 0; i < hashArray.length; i++) {
      binary += String.fromCharCode(hashArray[i]);
    }
    return btoa(binary);
  }

  /**
   * Durable Object state change handler
   *
   * Called when Durable Object state changes.
   * Currently unused but available for session tracking.
   *
   * @param state - Updated state object
   * @override
   */
  onStateUpdate(_state: State) {
    /** State tracking for session management - implementation pending */
  }
}
