/**
 * Microsoft OAuth Handler - Manages OAuth 2.1 + PKCE authorization flow
 * 
 * ARCHITECTURE:
 * - Handles OAuth client approval dialog before Microsoft authentication
 * - Implements client type detection (AI assistant vs MCP remote)
 * - Stores Microsoft authorization codes in OAuth Provider props
 * - Integrates with Cloudflare Workers OAuth Provider framework
 * 
 * FLOW:
 * 1. GET /authorize - Show approval dialog or redirect to Microsoft
 * 2. User approves via form submission
 * 3. POST /authorize - Redirect to Microsoft with client type detection
 * 4. GET /callback - Handle Microsoft callback, store auth code, redirect to client
 * 
 * Environment bindings accessed through Hono context (c.env)
 */
import type { AuthRequest, OAuthHelpers } from '@cloudflare/workers-oauth-provider';
import { Hono } from 'hono';
import { getUpstreamAuthorizeUrl } from './utils';
import {
  clientIdAlreadyApproved,
  parseRedirectApproval,
  renderApprovalDialog,
} from './workers-oauth-utils';
import { Env, initializeMCPClient } from './index';

/**
 * Hono app instance with typed environment bindings
 * Provides OAuth helper methods through OAUTH_PROVIDER binding
 */
const app = new Hono<{ Bindings: Env & { OAUTH_PROVIDER: OAuthHelpers } }>();

// ============================================================================
// AUTHORIZATION ENDPOINT - OAuth 2.1 + PKCE Flow Initiation
// ============================================================================

/**
 * GET /authorize - OAuth authorization endpoint
 * 
 * Handles initial authorization request from MCP client.
 * Shows approval dialog on first visit, bypasses if already approved.
 * 
 * @param c - Hono context with request and environment
 * @returns Response with approval dialog or redirect to Microsoft
 */
app.get('/authorize', async c => {
  /** Initialize static MCP client for session consistency */
  await initializeMCPClient(c.env);

  const oauthReqInfo = await c.env.OAUTH_PROVIDER.parseAuthRequest(c.req.raw);
  const { clientId } = oauthReqInfo;
  if (!clientId) {
    return c.text('Invalid request', 400);
  }

  if (
    await clientIdAlreadyApproved(c.req.raw, oauthReqInfo.clientId, c.env.COOKIE_ENCRYPTION_KEY)
  ) {
    return redirectToMicrosoft(c.req.raw, oauthReqInfo, c.env);
  }

  /**
   * Render approval dialog for first-time client authorization
   * User consent required before redirecting to Microsoft
   */
  return renderApprovalDialog(c.req.raw, {
    client: await c.env.OAUTH_PROVIDER.lookupClient(clientId),
    server: {
      description: 'Microsoft 365 MCP Server - Access your Office 365 data through AI tools.',
      logo: 'https://upload.wikimedia.org/wikipedia/commons/thumb/4/44/Microsoft_logo.svg/240px-Microsoft_logo.svg.png',
      name: 'Microsoft 365 MCP Server',
    },
    state: { oauthReqInfo },
  });
});

/**
 * POST /authorize - Process user approval and redirect to Microsoft
 * 
 * Receives form submission from approval dialog.
 * Parses encrypted state cookie and redirects to Microsoft Identity Platform.
 * 
 * @param c - Hono context with form data and cookies
 * @returns Response redirect to Microsoft authorization endpoint
 */
app.post('/authorize', async c => {
  const { state, headers } = await parseRedirectApproval(c.req.raw, c.env.COOKIE_ENCRYPTION_KEY);
  if (!state.oauthReqInfo) {
    return c.text('Invalid request', 400);
  }

  return redirectToMicrosoft(c.req.raw, state.oauthReqInfo, c.env, headers);
});

/**
 * Redirects user to Microsoft Identity Platform for authentication
 * 
 * Implements OAuth 2.1 authorization code flow with client type detection
 * based on redirect URI patterns (AI assistants vs MCP remote clients).
 * 
 * CLIENT TYPE DETECTION:
 * - AI Assistant: Contains '.ai', 'assistant', or 'chat' in redirect URI
 * - MCP Remote: Standard MCP client (default)
 * This affects metadata labeling and potential future feature differentiation.
 * 
 * @param request - Original request for URL construction
 * @param oauthReqInfo - OAuth request parameters from client
 * @param env - Environment containing Microsoft app registration details
 * @param headers - Additional headers (e.g., approval cookies)
 * @returns Response with 302 redirect to Microsoft authorization endpoint
 */
async function redirectToMicrosoft(
  request: Request,
  oauthReqInfo: AuthRequest,
  env: Env,
  headers: Record<string, string> = {}
) {
  /**
   * Client type detection based on redirect URI patterns
   * 
   * Detects whether the client is a web-based service or a local MCP client
   * based on the redirect URI. This is used for metadata labeling in OAuth.
   * Local clients (like mcp-remote) use localhost redirect URIs.
   */
  const isLocalClient = oauthReqInfo.redirectUri?.includes('localhost') || 
                        oauthReqInfo.redirectUri?.includes('127.0.0.1');
  const clientType = isLocalClient ? 'mcp-remote' : 'web-connector';

  /**
   * Redirect to Microsoft Identity Platform for user authentication
   * Uses OAuth 2.1 + PKCE authorization code flow
   * State parameter preserves original OAuth request info
   */
  return new Response(null, {
    headers: {
      ...headers,
      location: getUpstreamAuthorizeUrl({
        client_id: env.MICROSOFT_CLIENT_ID,
        redirect_uri: new URL('/callback', request.url).href,
        scope:
          'User.Read Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite Contacts.Read Contacts.ReadWrite People.Read People.Read.All OnlineMeetings.ReadWrite ChannelMessage.Send Team.ReadBasic.All offline_access',
        state: btoa(JSON.stringify({ ...oauthReqInfo, clientType })),
        upstream_url: `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/oauth2/v2.0/authorize`,
      }),
    },
    status: 302,
  });
}

// ============================================================================
// CALLBACK ENDPOINT - Microsoft OAuth Response Handler
// ============================================================================

/**
 * GET /callback - OAuth callback endpoint from Microsoft
 * 
 * CRITICAL FLOW:
 * 1. Extracts state from Microsoft callback (contains original OAuth request)
 * 2. Validates authorization code received from Microsoft
 * 3. Stores Microsoft auth code in OAuth Provider props (NOT exchanged here)
 * 4. Completes OAuth Provider authorization (triggers tokenExchangeCallback)
 * 5. Redirects client to original callback URL with OAuth access token
 * 
 * IMPORTANT: Token exchange happens in tokenExchangeCallback, not here.
 * This separation allows OAuth Provider to manage token lifecycle.
 * 
 * @param c - Hono context with query parameters from Microsoft
 * @returns Response redirect to client's original callback URL
 */
app.get('/callback', async c => {
  /**
   * Extract OAuth request info from state parameter
   * 
   * State contains original OAuth request encoded in base64.
   * Includes client type for metadata labeling.
   */
  const stateData = JSON.parse(atob(c.req.query('state') as string)) as AuthRequest & {
    clientType?: string;
  };
  const { clientType = 'mcp-remote', ...oauthReqInfo } = stateData;

  if (!oauthReqInfo.clientId) {
    return c.text('Invalid state', 400);
  }

  /**
   * Microsoft authorization code from query parameter
   * 
   * This code will be exchanged for tokens in tokenExchangeCallback.
   * Stored in props to pass through OAuth Provider flow.
   */
  const microsoftAuthCode = c.req.query('code');
  if (!microsoftAuthCode) {
    return c.text('No authorization code received from Microsoft', 400);
  }

  const redirectUri = new URL('/callback', c.req.url).href;

  /**
   * Complete OAuth authorization with OAuth Provider
   * 
   * CRITICAL: Microsoft auth code stored in props for token exchange callback.
   * Actual Microsoft tokens obtained via tokenExchangeCallback in index.ts.
   * 
   * Props flow:
   * 1. Store Microsoft auth code here
   * 2. OAuth Provider triggers tokenExchangeCallback
   * 3. Callback exchanges code for Microsoft tokens
   * 4. Tokens stored in newProps for persistence
   */
  const { redirectTo } = await c.env.OAUTH_PROVIDER.completeAuthorization({
    metadata: {
      label: `Microsoft 365 User (${clientType})`,
    },
    /**
     * Pass Microsoft authorization code for token exchange
     * 
     * These props will be available in tokenExchangeCallback.
     * microsoftRedirectUri must match the one used in authorization.
     */
    props: {
      microsoftAuthCode,
      microsoftRedirectUri: redirectUri,
      clientType,
    } as any,
    request: oauthReqInfo,
    scope: oauthReqInfo.scope,
    /** Temporary user ID - will be updated after token exchange with real Microsoft user ID */
    userId: 'microsoft_' + Date.now(),
  });

  return Response.redirect(redirectTo);
});

export { app as MicrosoftHandler };
