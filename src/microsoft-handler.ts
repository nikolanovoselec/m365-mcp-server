// No need to import env separately, we'll use c.env from Hono context
import type { AuthRequest, OAuthHelpers } from '@cloudflare/workers-oauth-provider';
import { Hono } from 'hono';
import { getUpstreamAuthorizeUrl } from './utils';
import {
  clientIdAlreadyApproved,
  parseRedirectApproval,
  renderApprovalDialog,
} from './workers-oauth-utils';
import { Env, initializeMCPClient } from './index';

const app = new Hono<{ Bindings: Env & { OAUTH_PROVIDER: OAuthHelpers } }>();

app.get('/authorize', async c => {
  // Initialize static MCP client if it doesn't exist
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

app.post('/authorize', async c => {
  const { state, headers } = await parseRedirectApproval(c.req.raw, c.env.COOKIE_ENCRYPTION_KEY);
  if (!state.oauthReqInfo) {
    return c.text('Invalid request', 400);
  }

  return redirectToMicrosoft(c.req.raw, state.oauthReqInfo, c.env, headers);
});

async function redirectToMicrosoft(
  request: Request,
  oauthReqInfo: AuthRequest,
  env: Env,
  headers: Record<string, string> = {}
) {
  // Detect client type based on redirect URI
  const isClaudeDesktop = oauthReqInfo.redirectUri?.includes('claude.ai');
  const clientType = isClaudeDesktop ? 'claude-desktop' : 'mcp-remote';
  
  console.log(`OAuth flow detected: ${clientType} (redirect: ${oauthReqInfo.redirectUri})`);
  
  return new Response(null, {
    headers: {
      ...headers,
      location: getUpstreamAuthorizeUrl({
        client_id: env.MICROSOFT_CLIENT_ID,
        redirect_uri: new URL('/callback', request.url).href,
        scope:
          'User.Read Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite Contacts.ReadWrite OnlineMeetings.ReadWrite ChannelMessage.Send Team.ReadBasic.All offline_access',
        state: btoa(JSON.stringify({ ...oauthReqInfo, clientType })),
        upstream_url: `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/oauth2/v2.0/authorize`,
      }),
    },
    status: 302,
  });
}

/**
 * OAuth Callback Endpoint
 *
 * This route handles the callback from Microsoft after user authentication.
 * It exchanges the temporary code for an access token, then stores some
 * user metadata & the auth token as part of the 'props' on the token passed
 * down to the client. It ends by redirecting the client back to _its_ callback URL
 */
app.get('/callback', async c => {
  // Get the oauthReqInfo out of state (now includes clientType)
  const stateData = JSON.parse(atob(c.req.query('state') as string)) as AuthRequest & { clientType?: string };
  const { clientType = 'mcp-remote', ...oauthReqInfo } = stateData;
  
  if (!oauthReqInfo.clientId) {
    return c.text('Invalid state', 400);
  }

  console.log(`OAuth callback for ${clientType}`);

  // Get the Microsoft authorization code
  const microsoftAuthCode = c.req.query('code');
  if (!microsoftAuthCode) {
    return c.text('No authorization code received from Microsoft', 400);
  }

  const redirectUri = new URL('/callback', c.req.url).href;

  // Store the Microsoft auth code in props - the tokenExchangeCallback will handle the actual token exchange
  // For now, we'll use a placeholder user ID - the actual user info will be obtained during token exchange
  const { redirectTo } = await c.env.OAUTH_PROVIDER.completeAuthorization({
    metadata: {
      label: `Microsoft 365 User (${clientType})`, // Include client type for debugging
    },
    // Store the Microsoft authorization code, redirect URI, and client type for the tokenExchangeCallback
    props: {
      microsoftAuthCode,
      microsoftRedirectUri: redirectUri,
      clientType,
    } as any,
    request: oauthReqInfo,
    scope: oauthReqInfo.scope,
    userId: 'microsoft_' + Date.now(), // Temporary user ID - will be updated
  });

  return Response.redirect(redirectTo);
});

export { app as MicrosoftHandler };
