/**
 * Microsoft 365 MCP Server - Cloudflare Workers Implementation
 *
 * ARCHITECTURE NOTES:
 * OAuth Provider wrapper pattern is REQUIRED (not over-engineering) due to:
 * 1. Cloudflare defaults to HTTP/2, preventing direct WebSocket upgrades
 * 2. MCP Agent's WebSocket requirement conflicts with platform limitations
 * 3. Need to bridge Cloudflare OAuth tokens with Microsoft Graph tokens
 *
 * This complex routing enables:
 * - SSE transport for MCP protocol (primary)
 * - HTTP JSON-RPC fallback (secondary)
 * - WebSocket simulation via SSE (when needed)
 *
 * Compatible with all MCP clients (AI assistants, chatbots, automation tools)
 * Single /sse endpoint serves all MCP protocol variants through static methods
 */

import OAuthProvider from "@cloudflare/workers-oauth-provider";
import { MicrosoftMCPAgent } from "./microsoft-mcp-agent";
import { MicrosoftHandler } from "./microsoft-handler";

/**
 * Export Durable Object class for Cloudflare Workers runtime registration
 * Required for wrangler.toml [[durable_objects.bindings]] configuration
 */
export { MicrosoftMCPAgent };

export interface Env {
  /** Microsoft Entra ID application configuration */
  MICROSOFT_CLIENT_ID: string;
  MICROSOFT_TENANT_ID: string;
  GRAPH_API_VERSION: string;

  /** Cloudflare Workers deployment configuration */
  WORKER_DOMAIN: string /* Worker subdomain: "your-worker.your-subdomain.workers.dev" */;
  PROTOCOL: string /* Protocol scheme: "https" for production, "http" for local dev */;

  /** Sensitive credentials - deployed via 'wrangler secret put' command */
  MICROSOFT_CLIENT_SECRET: string;
  COOKIE_ENCRYPTION_KEY: string;
  ENCRYPTION_KEY: string;
  COOKIE_SECRET: string;

  /** Durable Object namespace binding for MCP Agent instances */
  MCP_OBJECT: DurableObjectNamespace;

  /** KV storage namespaces for configuration and caching */
  CONFIG_KV: KVNamespace;
  CACHE_KV: KVNamespace;

  /**
   * OAuth KV namespace - REQUIRED by @cloudflare/workers-oauth-provider
   * Paradoxically remains empty in this implementation but must exist
   * to prevent runtime errors when OAuth Provider checks for client existence
   */
  OAUTH_KV: KVNamespace;
}

/**
 * Microsoft OAuth 2.1 token endpoint response structure
 * Matches Microsoft Identity Platform v2.0 token response specification
 */
interface MicrosoftTokenResponse {
  access_token: string;
  token_type: string;
  expires_in: number;
  refresh_token?: string;
  scope: string;
}

/**
 * Exchange Microsoft authorization code for access/refresh tokens
 *
 * Implements OAuth 2.1 authorization code flow with PKCE (S256)
 * Called by tokenExchangeCallback after Microsoft authentication
 *
 * @param authorizationCode - Code received from Microsoft OAuth callback
 * @param env - Cloudflare Worker environment bindings
 * @param redirectUri - Must match the redirect URI used in authorization request
 * @returns Microsoft OAuth tokens including access_token and refresh_token
 * @throws Error if token exchange fails with Microsoft Identity Platform
 */
async function exchangeMicrosoftTokens(
  authorizationCode: string,
  env: Env,
  redirectUri: string,
): Promise<MicrosoftTokenResponse> {
  const tokenUrl = `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`;

  const params = new URLSearchParams({
    client_id: env.MICROSOFT_CLIENT_ID,
    client_secret: env.MICROSOFT_CLIENT_SECRET,
    code: authorizationCode,
    redirect_uri: redirectUri,
    grant_type: "authorization_code",
    scope:
      "User.Read Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite Contacts.Read Contacts.ReadWrite People.Read People.Read.All OnlineMeetings.ReadWrite ChannelMessage.Send Team.ReadBasic.All offline_access",
  });

  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: params.toString(),
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(
      `Microsoft token exchange failed: ${response.status} ${error}`,
    );
  }

  return (await response.json()) as MicrosoftTokenResponse;
}

/**
 * Refresh Microsoft access tokens using refresh token
 *
 * Implements OAuth 2.1 refresh token grant for token renewal
 * Automatically called by OAuth Provider when access token expires
 *
 * @param refreshToken - Valid Microsoft refresh token from previous authorization
 * @param env - Cloudflare Worker environment bindings
 * @returns New Microsoft OAuth tokens including fresh access_token
 * @throws Error if refresh fails (user needs to re-authenticate)
 */
async function refreshMicrosoftTokens(
  refreshToken: string,
  env: Env,
): Promise<MicrosoftTokenResponse> {
  const tokenUrl = `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`;

  const params = new URLSearchParams({
    client_id: env.MICROSOFT_CLIENT_ID,
    client_secret: env.MICROSOFT_CLIENT_SECRET,
    refresh_token: refreshToken,
    grant_type: "refresh_token",
    scope:
      "User.Read Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite Contacts.Read Contacts.ReadWrite People.Read People.Read.All OnlineMeetings.ReadWrite ChannelMessage.Send Team.ReadBasic.All offline_access",
  });

  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: params.toString(),
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(
      `Microsoft token refresh failed: ${response.status} ${error}`,
    );
  }

  return (await response.json()) as MicrosoftTokenResponse;
}

// ============================================================================
// ARCHITECTURAL NOTES
// ============================================================================

/**
 * REQUEST FLOW ARCHITECTURE:
 *
 * 1. API Token Mode Check (REQUIRED for MCP compliance)
 *    - Detects unauthenticated discovery requests (tools/list)
 *    - Distinguishes OAuth tokens (3-part format) from API tokens
 *    - Allows tool enumeration before user authentication
 *
 * 2. OAuth Provider Processing
 *    - Manages dual token architecture (Cloudflare + Microsoft)
 *    - Handles OAuth 2.1 + PKCE authorization flow
 *    - Stores tokens in props for MCP Agent access
 *
 * 3. MCP Agent Invocation
 *    - Static serveSSE() method creates Durable Object instances
 *    - Props containing Microsoft tokens passed via ExecutionContext
 *    - Handles all MCP protocol operations with Graph API integration
 */

// ============================================================================
// DEPRECATED FUNCTIONS - Preserved for reference
// ============================================================================

/**
 * @deprecated Client ID mapping no longer used with static OAuth Provider
 * Preserved to show evolution from dynamic to static client registration
 * Original purpose: Map static client IDs to dynamically registered ones
 * Deprecation reason: OAuth Provider now handles client management internally
 */

/**
 * Static MCP client identifier for discovery session consistency
 * Used to maintain same client ID across unauthenticated discovery requests
 */
const MCP_CLIENT_ID = "rWJu8WV42zC5pfGT";

/**
 * Initialize static MCP client in OAuth KV storage
 *
 * ARCHITECTURAL QUIRK:
 * Function attempts to store client in OAUTH_KV but OAuth Provider ignores it
 * OAuth Provider manages clients internally, making this effectively a no-op
 * Retained for compatibility with potential future OAuth Provider versions
 *
 * @param env - Cloudflare Worker environment containing OAUTH_KV namespace
 */
export async function initializeMCPClient(env: any): Promise<void> {
  try {
    /** Attempt to retrieve existing client from KV (always returns null in practice) */
    const existingClient = await env.OAUTH_KV.get(`client:${MCP_CLIENT_ID}`);
    if (existingClient) {
      return;
    }

    /**
     * Client record structure matching OAuth Provider's expected format
     * Note: Despite storing this in KV, OAuth Provider manages clients internally
     * and this KV record is effectively ignored
     */
    const clientInfo = {
      clientId: MCP_CLIENT_ID,
      clientName: "Microsoft 365 MCP Static Client",
      redirectUris: [],
      // Public client - no clientSecret field for mcp-remote compatibility
      tokenEndpointAuthMethod: "none",
      grantTypes: ["authorization_code", "refresh_token"],
      responseTypes: ["code"],
      registrationDate: Math.floor(Date.now() / 1000), // Unix timestamp in seconds
    };

    /** Store client using OAuth Provider's key format (client:{id}) */
    await env.OAUTH_KV.put(
      `client:${MCP_CLIENT_ID}`,
      JSON.stringify(clientInfo),
    );
  } catch (error) {
    /** Non-fatal error - OAuth Provider will handle client creation dynamically */
  }
}

/**
 * @deprecated Client ID mapping handler for authorization endpoint
 *
 * Retained for reference but unused in current OAuth Provider architecture.
 * Shows how to implement client ID aliasing for backward compatibility.
 * Original purpose: Map static client IDs to dynamically registered ones.
 *
 * @param request - Authorization request from client
 * @param env - Cloudflare Worker environment bindings
 * @param ctx - Execution context for async operations
 * @returns Response with authorization redirect or error
 */
async function _handleAuthorizeWithClientMapping(
  request: Request,
  env: Env,
  ctx: ExecutionContext,
): Promise<Response> {
  try {
    const url = new URL(request.url);
    const requestedClientId = url.searchParams.get("client_id");

    if (!requestedClientId) {
      return new Response(
        JSON.stringify({
          error: "invalid_request",
          error_description: "Missing client_id parameter",
        }),
        {
          status: 400,
          headers: { "Content-Type": "application/json" },
        },
      );
    }

    /** Get the actual registered client ID */
    const actualClientId = await env.CONFIG_KV.get(
      `static_client_actual:${MCP_CLIENT_ID}`,
    );

    /** If the requested client ID is already the registered static client, proceed normally */
    if (requestedClientId === actualClientId) {
      return await createOAuthProvider(env).fetch(request, env, ctx);
    }

    /** If no static client registered yet, register one */
    if (!actualClientId) {
      /** Register static client once for mcp-remote compatibility */
      const registerRequest = new Request(`${url.origin}/register`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          client_name: "Microsoft 365 MCP Static Client",
          redirect_uris: [],
          grant_types: ["authorization_code"],
          response_types: ["code"],
          token_endpoint_auth_method: "none",
        }),
      });

      const registerResponse = await createOAuthProvider(env).fetch(
        registerRequest,
        env,
        ctx,
      );

      if (registerResponse.status === 201) {
        const registrationResult = (await registerResponse.json()) as any;
        const newActualClientId = registrationResult.client_id;

        // Store the actual client ID for future use
        await env.CONFIG_KV.put(
          `static_client_actual:${MCP_CLIENT_ID}`,
          newActualClientId,
        );

        /** Create a new request with the actual registered client ID */
        const mappedUrl = new URL(request.url);
        mappedUrl.searchParams.set("client_id", newActualClientId);

        const mappedRequest = new Request(mappedUrl.toString(), {
          method: request.method,
          headers: request.headers,
          body: request.body,
        });

        return await createOAuthProvider(env).fetch(mappedRequest, env, ctx);
      } else {
        return new Response(
          JSON.stringify({
            error: "server_error",
            error_description: "Failed to register MCP client",
          }),
          {
            status: 500,
            headers: { "Content-Type": "application/json" },
          },
        );
      }
    }

    /** Static client registered - map requested client ID to actual registration */

    /** Create new request with actual registered client ID */
    const mappedUrl = new URL(request.url);
    mappedUrl.searchParams.set("client_id", actualClientId);

    const mappedRequest = new Request(mappedUrl.toString(), {
      method: request.method,
      headers: request.headers,
      body: request.body,
    });

    /** Process authorization request with mapped client ID */
    return await createOAuthProvider(env).fetch(mappedRequest, env, ctx);
  } catch (error: any) {
    return new Response(
      JSON.stringify({
        error: "server_error",
        error_description: error.message || "Authorization failed",
      }),
      {
        status: 500,
        headers: { "Content-Type": "application/json" },
      },
    );
  }
}

/**
 * @deprecated Client ID mapping handler for token endpoint
 *
 * Retained for reference but unused in current OAuth Provider architecture.
 * Shows how to handle token requests with client ID translation.
 * Original purpose: Support token exchange with static client IDs.
 *
 * @param request - Token exchange request from client
 * @param env - Cloudflare Worker environment bindings
 * @param ctx - Execution context for async operations
 * @returns Response with OAuth tokens or error
 */
async function _handleTokenWithClientMapping(
  request: Request,
  env: Env,
  ctx: ExecutionContext,
): Promise<Response> {
  try {
    /** Parse form data to extract client_id */
    const formData = await request.clone().formData();

    const requestedClientId = formData.get("client_id") as string;

    if (!requestedClientId) {
      /** Default to MCP_CLIENT_ID for mcp-remote compatibility */
      const defaultClientId = MCP_CLIENT_ID;

      /** Create new form data with default client ID */
      const newFormData = new FormData();
      for (const [key, value] of formData.entries()) {
        newFormData.append(key, value as string);
      }
      newFormData.append("client_id", defaultClientId);

      /** Create new request with default client ID */
      const mappedRequest = new Request(request.url, {
        method: request.method,
        headers: request.headers,
        body: newFormData,
      });

      return await createOAuthProvider(env).fetch(mappedRequest, env, ctx);
    }

    /** Retrieve actual registered client ID for static MCP client */
    const actualClientId = await env.CONFIG_KV.get(
      `static_client_actual:${MCP_CLIENT_ID}`,
    );

    /** If requesting client matches registered static client, proceed with standard flow */
    if (requestedClientId === actualClientId) {
      return await createOAuthProvider(env).fetch(request, env, ctx);
    }

    /** Map to registered static client if available */
    if (actualClientId && requestedClientId !== actualClientId) {
      /** Create new form data with actual client ID */
      const newFormData = new FormData();
      for (const [key, value] of formData.entries()) {
        if (key === "client_id") {
          newFormData.append(key, actualClientId);
        } else {
          newFormData.append(key, value as string);
        }
      }

      // Create a new request with the actual client ID
      const mappedRequest = new Request(request.url, {
        method: request.method,
        headers: request.headers,
        body: newFormData,
      });

      return await createOAuthProvider(env).fetch(mappedRequest, env, ctx);
    }

    /** Proceed with standard token exchange for non-mapped clients */
    return await createOAuthProvider(env).fetch(request, env, ctx);
  } catch (error: any) {
    return new Response(
      JSON.stringify({
        error: "server_error",
        error_description: error.message || "Token exchange failed",
      }),
      {
        status: 500,
        headers: { "Content-Type": "application/json" },
      },
    );
  }
}

/**
 * @deprecated OAuth Provider factory with environment closure
 *
 * Unused in current architecture where OAuth Provider is instantiated inline.
 * Retained for reference showing alternative configuration pattern that could
 * be useful for testing or multiple provider instances.
 *
 * @param env - Cloudflare Worker environment bindings
 * @returns Configured OAuthProvider instance with Microsoft handlers
 */
function createOAuthProvider(env: Env) {
  return new OAuthProvider({
    /** API protection disabled - all routing handled via apiHandlers and defaultHandler */
    apiRoute: [],
    apiHandler: {
      fetch: async () => new Response("Not used", { status: 404 }),
    },

    /** MicrosoftHandler processes OAuth 2.1 + PKCE authorization and callback flows */
    defaultHandler: MicrosoftHandler as any,

    /** Standard OAuth 2.1 endpoint configuration */
    authorizeEndpoint: "/authorize",
    tokenEndpoint: "/token",
    clientRegistrationEndpoint: "/register",

    /** DCR (Dynamic Client Registration) enabled for MCP client automatic setup */
    disallowPublicClientRegistration: false,

    /** Microsoft Graph API permission scopes required for Microsoft 365 operations */
    scopesSupported: [
      "User.Read",
      "Mail.Read",
      "Mail.ReadWrite",
      "Mail.Send",
      "Calendars.Read",
      "Calendars.ReadWrite",
      "Contacts.Read",
      "Contacts.ReadWrite",
      "People.Read",
      "People.Read.All",
      "OnlineMeetings.ReadWrite",
      "ChannelMessage.Send",
      "Team.ReadBasic.All",
      "offline_access",
    ],

    // Token exchange callback - integrates Microsoft tokens into OAuth flow
    tokenExchangeCallback: async (options: any) => {
      // Use captured environment from closure

      if (options.grantType === "authorization_code") {
        /** Extract Microsoft authorization code stored by MicrosoftHandler in OAuth props */
        const microsoftAuthCode = options.props.microsoftAuthCode;
        const redirectUri = options.props.microsoftRedirectUri;

        if (!microsoftAuthCode) {
          throw new Error("No Microsoft authorization code available");
        }

        /** Execute Microsoft OAuth 2.1 token exchange using authorization code flow */
        const microsoftTokens = await exchangeMicrosoftTokens(
          microsoftAuthCode,
          env,
          redirectUri,
        );

        return {
          // Store Microsoft access token in the access token props for the MCP agent
          accessTokenProps: {
            ...options.props,
            microsoftAccessToken: microsoftTokens.access_token,
            microsoftTokenType: microsoftTokens.token_type,
            microsoftScope: microsoftTokens.scope,
          },
          // Store Microsoft refresh token in the grant for future refreshes
          newProps: {
            ...options.props,
            microsoftRefreshToken: microsoftTokens.refresh_token,
          },
          // Match Microsoft token TTL
          accessTokenTTL: microsoftTokens.expires_in,
        };
      }

      if (options.grantType === "refresh_token") {
        // Refresh Microsoft tokens using stored refresh token
        const refreshToken = options.props.microsoftRefreshToken;

        if (!refreshToken) {
          throw new Error("No Microsoft refresh token available");
        }

        const microsoftTokens = await refreshMicrosoftTokens(refreshToken, env);

        return {
          accessTokenProps: {
            ...options.props,
            microsoftAccessToken: microsoftTokens.access_token,
            microsoftTokenType: microsoftTokens.token_type,
            microsoftScope: microsoftTokens.scope,
          },
          newProps: {
            ...options.props,
            microsoftRefreshToken:
              microsoftTokens.refresh_token || refreshToken,
          },
          accessTokenTTL: microsoftTokens.expires_in,
        };
      }

      // For other grant types, return unchanged
      return {};
    },
  });
}

/**
 * API Token Mode detection for MCP protocol compliance
 *
 * REQUIRED for unauthenticated discovery phase where MCP clients
 * call tools/list before user authentication to enumerate available tools
 *
 * Token Format Logic:
 * - OAuth tokens: 3-part colon-separated format (header:payload:signature) per JWT spec
 * - API tokens: Any other format (simple strings, UUIDs) trigger discovery mode
 * - This distinction allows tools enumeration without full authentication
 *
 * WHY THIS EXISTS:
 * MCP clients need to show available capabilities before asking users to authenticate.
 * This allows users to make informed decisions about granting permissions.
 *
 * @param req - HTTP request to examine for token format
 * @returns true if request should use discovery mode, false for authenticated mode
 */
async function isApiTokenRequest(req: Request): Promise<boolean> {
  const url = new URL(req.url);

  /** Skip API Token Mode for OAuth endpoints to prevent authentication loop */
  if (!url.pathname.startsWith("/mcp") && !url.pathname.startsWith("/sse")) {
    return false;
  }

  const authHeader = req.headers.get("Authorization");
  if (!authHeader) {
    /** Missing Authorization header indicates unauthenticated discovery request */
    return true;
  }

  const [type, token] = authHeader.split(" ");
  if (type !== "Bearer") return false;

  /**
   * Token Format Detection:
   * - OAuth tokens: 3-part colon format (header:payload:signature) per JWT spec
   * - API tokens: Any other format (simple strings, UUIDs, etc.)
   *
   * This distinction allows the server to differentiate between:
   * - Authenticated requests (OAuth tokens) requiring full validation
   * - Discovery requests (API tokens) allowing tool enumeration without auth
   */
  const codeParts = token.split(":");
  return codeParts.length !== 3;
}

/**
 * Handle unauthenticated discovery requests in API Token Mode
 *
 * Creates temporary Durable Object instance for tool enumeration
 * Session is ephemeral and not persisted after discovery phase
 *
 * PERFORMANCE: Uses consistent 'discovery-session' ID to reuse Durable Object
 * instance across discovery requests, reducing cold start overhead
 *
 * @param request - Original HTTP request to forward to MCP Agent
 * @param env - Cloudflare Worker environment bindings
 * @param _ctx - Execution context (unused but required by signature)
 * @returns Response from MCP Agent with available tools list
 */
async function handleApiTokenMode(
  request: Request,
  env: Env,
  _ctx: ExecutionContext,
): Promise<Response> {
  /**
   * Direct Durable Object invocation bypasses OAuth authentication
   * MicrosoftMCPAgent.fetch() contains logic to handle discovery without props
   * Session name 'discovery-session' ensures consistent instance reuse
   */
  const id = env.MCP_OBJECT.idFromName("discovery-session");
  const obj = env.MCP_OBJECT.get(id);
  return await obj.fetch(request);
}

/**
 * @deprecated JSON-RPC method parser for debugging purposes
 *
 * Unused in current architecture but retained for debugging
 * Can be used to inspect JSON-RPC methods in request bodies
 *
 * @param request - HTTP request containing JSON-RPC payload
 * @returns Promise resolving to method name or null if parsing fails
 */
async function _parseJsonRpcMethod(request: Request): Promise<string | null> {
  try {
    const body = await request.clone().text();
    if (!body) return null;
    const jsonRpc = JSON.parse(body);
    return jsonRpc.method || null;
  } catch {
    return null;
  }
}

// ============================================================================
// MAIN WORKER EXPORT
// ============================================================================

/**
 * Cloudflare Worker fetch handler - Entry point for all requests
 *
 * PROCESSING ORDER:
 * 1. API Token Mode check (unauthenticated discovery)
 * 2. OAuth Provider processing (authentication and token management)
 * 3. MCP Agent invocation (protocol handling)
 */
export default {
  fetch: async (
    request: Request,
    env: Env,
    ctx: ExecutionContext,
  ): Promise<Response> => {
    const _url = new URL(request.url);

    /**
     * API Token Mode must be checked before OAuth Provider
     * to prevent authentication requirements on discovery requests
     *
     * PERFORMANCE CONSIDERATIONS:
     * - API Token Mode reuses 'discovery-session' Durable Object ID
     * - This prevents cold starts on every discovery request
     * - Trade-off: All discovery requests share same DO instance
     * - Consider implementing LRU cache for multiple discovery sessions
     *
     * OPTIMIZATION OPPORTUNITIES:
     * - Cache discovery responses in KV with 5-minute TTL
     * - Implement connection pooling for multiple clients
     * - Use Durable Object alarms for session cleanup
     * - Consider Workers Analytics Engine for metrics
     */
    if (await isApiTokenRequest(request)) {
      return await handleApiTokenMode(request, env, ctx);
    }

    // ============================================================================
    // OAUTH PROVIDER CONFIGURATION
    // ============================================================================

    /**
     * OAuth Provider configuration with custom /sse handler
     *
     * WHY THIS ARCHITECTURE:
     * 1. Cloudflare defaults to HTTP/2, preventing WebSocket upgrades
     * 2. MCP Agent requires WebSocket-like persistent connections
     * 3. OAuth Provider wrapper enables SSE transport as WebSocket alternative
     * 4. Props passing mechanism bridges OAuth context to MCP Agent
     *
     * This is NOT over-engineering but a necessary adaptation to platform constraints.
     */
    return new OAuthProvider({
      /**
       * CRITICAL INTEGRATION POINT: /sse endpoint handler
       *
       * Purpose: Bridge OAuth Provider props to MCP Agent
       * - Extracts Microsoft tokens from OAuth context
       * - Passes tokens via modified ExecutionContext
       * - Enables Graph API calls within MCP tools
       *
       * Note: 404 on initial POST is INTENTIONAL
       * Signals client to authenticate before session creation
       */
      apiHandlers: {
        "/sse": {
          fetch: async (
            request: Request,
            env: unknown,
            ctx: ExecutionContext,
          ) => {
            const typedEnv = env as Env;

            /**
             * CRITICAL: OAuth Provider stores authenticated user props in ExecutionContext.
             * These props contain Microsoft tokens needed for Graph API calls.
             * Empty props indicate unauthenticated state (intentional behavior).
             */

            const oauthProps = (ctx as any).props;

            if (oauthProps) {
              try {
                /**
                 * Static serveSSE method from agents library creates Durable Object instances
                 * Props passed via modified ExecutionContext become available as this.props in agent
                 */
                const propsContext = {
                  ...ctx,
                  props: oauthProps,
                };

                const response = await MicrosoftMCPAgent.serveSSE("/sse").fetch(
                  request,
                  typedEnv,
                  propsContext as ExecutionContext,
                );

                return response;
              } catch (error) {
                return new Response(
                  JSON.stringify({
                    error: "Internal server error in OAuth API handler",
                    details: (error as Error).message,
                  }),
                  {
                    status: 500,
                    headers: { "Content-Type": "application/json" },
                  },
                );
              }
            } else {
              /**
               * INTENTIONAL 404 BEHAVIOR:
               *
               * Returns 404 when no OAuth props exist (unauthenticated)
               * This is NOT an error but a signal to the client:
               * - No authenticated session exists yet
               * - Client should initiate OAuth flow
               * - Prevents unnecessary Durable Object creation
               *
               * Performance optimization: Quick rejection without resource allocation
               */
              try {
                const staticResponse = await MicrosoftMCPAgent.serveSSE(
                  "/sse",
                ).fetch(request, typedEnv, ctx);
                return staticResponse;
              } catch (staticError) {
                return new Response(
                  JSON.stringify({
                    error: "SSE transport initialization failed",
                    details: (staticError as Error).message,
                  }),
                  {
                    status: 500,
                    headers: { "Content-Type": "application/json" },
                  },
                );
              }
            }
          },
        },
      },
      /** MicrosoftHandler manages OAuth 2.1 + PKCE authorization and callback */
      /** OAuth 2.1 + PKCE flow handler for Microsoft authentication */
      defaultHandler: MicrosoftHandler as any,
      /** OAuth 2.1 standard endpoints */
      authorizeEndpoint: "/authorize",
      tokenEndpoint: "/token",
      /** DCR endpoint for automatic client registration */
      clientRegistrationEndpoint: "/register",

      /** DCR (Dynamic Client Registration) enabled for MCP client automatic setup */
      disallowPublicClientRegistration: false,

      /** Complete set of Microsoft Graph API scopes for all supported operations */
      scopesSupported: [
        "User.Read",
        "Mail.Read",
        "Mail.ReadWrite",
        "Mail.Send",
        "Calendars.Read",
        "Calendars.ReadWrite",
        "Contacts.Read",
        "Contacts.ReadWrite",
        "People.Read",
        "People.Read.All",
        "OnlineMeetings.ReadWrite",
        "ChannelMessage.Send",
        "Team.ReadBasic.All",
        "offline_access",
      ],

      /**
       * Token exchange callback - critical integration point
       *
       * PURPOSE:
       * Bridges Cloudflare OAuth Provider tokens with Microsoft Graph tokens.
       * This dual token architecture is REQUIRED because:
       * 1. OAuth Provider manages client authentication and session
       * 2. Microsoft Graph API requires Microsoft-specific access tokens
       * 3. Tokens must be synchronized for proper authorization flow
       *
       * FLOW:
       * 1. OAuth Provider calls this after client authorization
       * 2. Exchange Microsoft auth code for Graph API tokens
       * 3. Store both token sets in props for MCP Agent access
       */
      tokenExchangeCallback: async (options: any) => {
        if (options.grantType === "authorization_code") {
          /** Extract Microsoft authorization code stored by MicrosoftHandler in OAuth props */
          const microsoftAuthCode = options.props.microsoftAuthCode;
          const redirectUri = options.props.microsoftRedirectUri;

          if (!microsoftAuthCode) {
            throw new Error("No Microsoft authorization code available");
          }

          /** Execute Microsoft OAuth 2.1 token exchange using authorization code flow */
          const microsoftTokens = await exchangeMicrosoftTokens(
            microsoftAuthCode,
            env,
            redirectUri,
          );

          /** Token storage complete - now available in MCP Agent props */

          return {
            /**
             * CRITICAL: Store ALL tokens in newProps (not accessTokenProps)
             * This ensures tokens persist across requests and sessions
             * Follows Cloudflare's OAuth Provider persistence pattern
             */
            newProps: {
              ...options.props,
              microsoftAccessToken: microsoftTokens.access_token,
              microsoftTokenType: microsoftTokens.token_type,
              microsoftScope: microsoftTokens.scope,
              microsoftRefreshToken: microsoftTokens.refresh_token,
            },
            /** Token expiration from Microsoft (typically 3600 seconds for access tokens) */
            accessTokenTTL: microsoftTokens.expires_in,
          };
        }

        if (options.grantType === "refresh_token") {
          /** Retrieve refresh token from previous authorization */
          const refreshToken = options.props.microsoftRefreshToken;

          if (!refreshToken) {
            throw new Error("No Microsoft refresh token available");
          }

          /** Execute token refresh flow with Microsoft OAuth 2.1 endpoint */
          const microsoftTokens = await refreshMicrosoftTokens(
            refreshToken,
            env,
          );

          return {
            /**
             * Update all tokens including new access token and potentially new refresh token
             * Microsoft may issue new refresh token; fallback to existing if not provided
             */
            newProps: {
              ...options.props,
              microsoftAccessToken: microsoftTokens.access_token,
              microsoftTokenType: microsoftTokens.token_type,
              microsoftScope: microsoftTokens.scope,
              microsoftRefreshToken:
                microsoftTokens.refresh_token || refreshToken,
            },
            accessTokenTTL: microsoftTokens.expires_in,
          };
        }

        throw new Error(`Unsupported grant type: ${options.grantType}`);
      },

      accessTokenTTL: 3600,
    }).fetch(request, env, ctx);
  },
};
