/**
 * Microsoft 365 MCP Server - Unified Endpoint Architecture
 * Supports Claude Desktop direct connector + mcp-remote with single /sse endpoint
 */

import OAuthProvider from '@cloudflare/workers-oauth-provider';
import { MicrosoftMCPAgent } from './microsoft-mcp-agent';
import { MicrosoftHandler } from './microsoft-handler';

export interface Env {
  // Environment variables
  MICROSOFT_CLIENT_ID: string;
  MICROSOFT_TENANT_ID: string;
  GRAPH_API_VERSION: string;
  
  // Deployment configuration
  WORKER_DOMAIN: string;  // e.g., "your-worker.your-subdomain.workers.dev"
  PROTOCOL: string;       // "https" or "http" for development

  // Secrets
  MICROSOFT_CLIENT_SECRET: string;
  COOKIE_ENCRYPTION_KEY: string;
  ENCRYPTION_KEY: string;
  COOKIE_SECRET: string;

  // Durable Objects
  MICROSOFT_MCP_AGENT: DurableObjectNamespace;

  // KV Namespaces
  CONFIG_KV: KVNamespace;
  CACHE_KV: KVNamespace;

  // OAuth KV (required by OAuthProvider)
  OAUTH_KV: KVNamespace;
}

// Microsoft OAuth token response type
interface MicrosoftTokenResponse {
  access_token: string;
  token_type: string;
  expires_in: number;
  refresh_token?: string;
  scope: string;
}

// Microsoft OAuth token exchange functions
async function exchangeMicrosoftTokens(
  authorizationCode: string,
  env: Env,
  redirectUri: string
): Promise<MicrosoftTokenResponse> {
  const tokenUrl = `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`;

  const params = new URLSearchParams({
    client_id: env.MICROSOFT_CLIENT_ID,
    client_secret: env.MICROSOFT_CLIENT_SECRET,
    code: authorizationCode,
    redirect_uri: redirectUri,
    grant_type: 'authorization_code',
    scope:
      'User.Read Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite Contacts.ReadWrite OnlineMeetings.ReadWrite ChannelMessage.Send Team.ReadBasic.All offline_access',
  });

  const response = await fetch(tokenUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: params.toString(),
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Microsoft token exchange failed: ${response.status} ${error}`);
  }

  return (await response.json()) as MicrosoftTokenResponse;
}

async function refreshMicrosoftTokens(
  refreshToken: string,
  env: Env
): Promise<MicrosoftTokenResponse> {
  const tokenUrl = `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`;

  const params = new URLSearchParams({
    client_id: env.MICROSOFT_CLIENT_ID,
    client_secret: env.MICROSOFT_CLIENT_SECRET,
    refresh_token: refreshToken,
    grant_type: 'refresh_token',
    scope:
      'User.Read Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite Contacts.ReadWrite OnlineMeetings.ReadWrite ChannelMessage.Send Team.ReadBasic.All offline_access',
  });

  const response = await fetch(tokenUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: params.toString(),
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Microsoft token refresh failed: ${response.status} ${error}`);
  }

  return (await response.json()) as MicrosoftTokenResponse;
}

// 🔥 UNIFIED ENDPOINT ARCHITECTURE
// Single /sse endpoint supporting GET, POST, and WebSocket for all clients
// - Claude Desktop: GET (SSE validation) + WebSocket upgrade
// - mcp-remote: WebSocket with OAuth flow
// - Direct testing: POST JSON-RPC

// Hybrid MCP handler supporting GET (SSE), POST (JSON-RPC), and WebSocket
async function handleHybridMcp(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
  const url = new URL(request.url);
  
  // Protocol 1: WebSocket upgrade requests (check FIRST - WebSocket upgrades are GET requests)
  const upgradeHeader = request.headers.get('Upgrade');
  const webSocketKey = request.headers.get('Sec-WebSocket-Key');
  const webSocketVersion = request.headers.get('Sec-WebSocket-Version');
  
  if ((upgradeHeader && upgradeHeader.toLowerCase() === 'websocket') || 
      (webSocketKey && webSocketVersion)) {
    console.log('WebSocket upgrade detected - using simplified WebSocket handler');
    
    // For now, return the old behavior that worked (delegate to OAuth provider)  
    // TODO: Implement proper WebSocket handling following Cloudflare template
    try {
      // Generate unique session ID for WebSocket connection
      const sessionId = crypto.randomUUID();
      const id = env.MICROSOFT_MCP_AGENT.idFromName(`ws-${sessionId}`);
      const stub = env.MICROSOFT_MCP_AGENT.get(id);
      
      // Add required headers for MCP agent WebSocket handling
      const modifiedRequest = new Request(request, {
        headers: {
          ...Object.fromEntries(request.headers.entries()),
          'X-WebSocket-Session': sessionId,
          'X-MCP-Mode': 'websocket'
        }
      });
      
      return stub.fetch(modifiedRequest);
    } catch (error) {
      console.error('WebSocket upgrade error:', error);
      return new Response(`WebSocket upgrade failed: ${error}`, { status: 500 });
    }
  }
  
  // Protocol 2: GET requests with SSE headers (Claude Desktop MCP over SSE)
  if (request.method === 'GET') {
    const acceptHeader = request.headers.get('Accept');
    if (acceptHeader && acceptHeader.includes('text/event-stream')) {
      console.log('SSE MCP connection requested');
      return handleSSEMcp(request, env);
    }
    return new Response('GET requests must include Accept: text/event-stream for SSE or provide WebSocket headers', { status: 400 });
  }
  
  // Protocol 3: POST requests (Direct JSON-RPC for testing)
  if (request.method === 'POST') {
    return handleDirectJsonRpc(request, env);
  }
  
  return new Response('Unsupported method - Use GET (SSE), POST (JSON-RPC), or WebSocket upgrade', { status: 405 });
}

// Client ID mapping for MCP compatibility
const CLIENT_ID_MAPPING: Record<string, string> = {};

// Pre-registered client ID for mcp-remote compatibility (MCP spec allows hardcoded client IDs)
const MCP_CLIENT_ID = 'rWJu8WV42zC5pfGT';

// Initialize static MCP client using OAuth helpers when available
export async function initializeMCPClient(env: any): Promise<void> {
  try {
    console.log(`Checking if MCP client exists: ${MCP_CLIENT_ID}`);
    
    // Check if client already exists in KV
    const existingClient = await env.OAUTH_KV.get(`client:${MCP_CLIENT_ID}`);
    if (existingClient) {
      console.log(`Static MCP client already exists: ${MCP_CLIENT_ID}`);
      return;
    }
    
    console.log(`Creating static MCP client: ${MCP_CLIENT_ID}`);
    
    // Manually create client record with the specific client ID we need
    // The OAuth provider's createClient() generates random IDs, which doesn't work for our static use case
    // Format must match exactly what the OAuth provider expects
    const clientInfo = {
      clientId: MCP_CLIENT_ID,
      clientName: 'Microsoft 365 MCP Static Client',
      redirectUris: [],
      // Public client - no clientSecret field for mcp-remote compatibility
      tokenEndpointAuthMethod: 'none',
      grantTypes: ['authorization_code', 'refresh_token'],
      responseTypes: ['code'],
      registrationDate: Math.floor(Date.now() / 1000) // Unix timestamp in seconds
    };
    
    // Store client in KV using the same key format as OAuthProvider
    await env.OAUTH_KV.put(`client:${MCP_CLIENT_ID}`, JSON.stringify(clientInfo));
    
    console.log(`Successfully created static MCP client: ${MCP_CLIENT_ID}`);
  } catch (error) {
    console.error('Failed to initialize MCP client:', error);
    // Don't throw - let the worker continue, client might get created later
  }
}

// Handle /authorize with fixed client ID for mcp-remote compatibility
async function handleAuthorizeWithClientMapping(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
  try {
    const url = new URL(request.url);
    const requestedClientId = url.searchParams.get('client_id');
    
    if (!requestedClientId) {
      return new Response(JSON.stringify({
        error: 'invalid_request',
        error_description: 'Missing client_id parameter'
      }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' }
      });
    }

    console.log(`Authorization request for client_id: ${requestedClientId}`);
    
    // Get the actual registered client ID
    const actualClientId = await env.CONFIG_KV.get(`static_client_actual:${MCP_CLIENT_ID}`);
    
    // If the requested client ID is already our registered static client, proceed normally
    if (requestedClientId === actualClientId) {
      console.log(`Client ID ${requestedClientId} is already the registered static client, proceeding normally`);
      return await createOAuthProvider(env).fetch(request, env, ctx);
    }
    
    // If we don't have a static client registered yet, register one
    if (!actualClientId) {
      console.log(`Registering static MCP client: ${MCP_CLIENT_ID}`);
      
      // Register our static client once
      const registerRequest = new Request(`${url.origin}/register`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          client_name: 'Microsoft 365 MCP Static Client',
          redirect_uris: [],
          grant_types: ['authorization_code'],
          response_types: ['code'],
          token_endpoint_auth_method: 'none'
        })
      });
      
      const registerResponse = await createOAuthProvider(env).fetch(registerRequest, env, ctx);
      
      if (registerResponse.status === 201) {
        const registrationResult = await registerResponse.json() as any;
        const newActualClientId = registrationResult.client_id;
        
        // Store the actual client ID for future use
        await env.CONFIG_KV.put(`static_client_actual:${MCP_CLIENT_ID}`, newActualClientId);
        
        console.log(`Registered static MCP client: ${MCP_CLIENT_ID} -> ${newActualClientId}`);
        
        // Now use the newly registered client ID
        console.log(`Using static MCP client: ${requestedClientId} -> ${newActualClientId}`);
        
        // Create a new request with the actual registered client ID
        const mappedUrl = new URL(request.url);
        mappedUrl.searchParams.set('client_id', newActualClientId);
        
        const mappedRequest = new Request(mappedUrl.toString(), {
          method: request.method,
          headers: request.headers,
          body: request.body
        });
        
        return await createOAuthProvider(env).fetch(mappedRequest, env, ctx);
      } else {
        console.error('Failed to register static client:', await registerResponse.text());
        return new Response(JSON.stringify({
          error: 'server_error',
          error_description: 'Failed to register MCP client'
        }), {
          status: 500,
          headers: { 'Content-Type': 'application/json' }
        });
      }
    }
    
    // We have a static client registered, map the requested client ID to it
    console.log(`Using static MCP client: ${requestedClientId} -> ${actualClientId}`);
    
    // Create a new request with the actual registered client ID
    const mappedUrl = new URL(request.url);
    mappedUrl.searchParams.set('client_id', actualClientId);
    
    const mappedRequest = new Request(mappedUrl.toString(), {
      method: request.method,
      headers: request.headers,
      body: request.body
    });
    
    // Process the authorization request with the mapped client ID
    return await createOAuthProvider(env).fetch(mappedRequest, env, ctx);
    
  } catch (error: any) {
    console.error('Authorization error:', error);
    return new Response(JSON.stringify({
      error: 'server_error',
      error_description: error.message || 'Authorization failed'
    }), {
      status: 500,
      headers: { 'Content-Type': 'application/json' }
    });
  }
}


// MCP Agent will be mounted per-request to avoid binding issues

// Handle /token with static client ID mapping for mcp-remote compatibility  
async function handleTokenWithClientMapping(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
  try {
    console.log('=== TOKEN EXCHANGE DEBUG ===');
    console.log('Request method:', request.method);
    console.log('Request URL:', request.url);
    console.log('Request headers:', Object.fromEntries(request.headers.entries()));
    
    // Check content type
    const contentType = request.headers.get('content-type');
    console.log('Content-Type:', contentType);
    
    // Parse the form data to get client_id
    const formData = await request.clone().formData();
    console.log('Form data entries:');
    for (const [key, value] of formData.entries()) {
      console.log(`  ${key}: ${key === 'client_secret' ? '[REDACTED]' : value}`);
    }
    
    const requestedClientId = formData.get('client_id') as string;
    
    if (!requestedClientId) {
      console.log('Token request missing client_id, using default MCP client_id');
      // For mcp-remote compatibility, use the expected client_id
      const defaultClientId = MCP_CLIENT_ID;
      
      // Create new form data with default client ID
      const newFormData = new FormData();
      for (const [key, value] of formData.entries()) {
        newFormData.append(key, value as string);
      }
      newFormData.append('client_id', defaultClientId);
      
      // Create a new request with the default client ID
      const mappedRequest = new Request(request.url, {
        method: request.method,
        headers: request.headers,
        body: newFormData
      });
      
      console.log('Using default client_id for token exchange:', defaultClientId);
      return await createOAuthProvider(env).fetch(mappedRequest, env, ctx);
    }
    
    console.log(`Token request for client_id: ${requestedClientId}`);
    
    // Get the actual registered client ID for our static MCP client
    const actualClientId = await env.CONFIG_KV.get(`static_client_actual:${MCP_CLIENT_ID}`);
    
    // If the requested client ID is already our registered static client, proceed normally
    if (requestedClientId === actualClientId) {
      console.log(`Token client ID ${requestedClientId} is already the registered static client, proceeding normally`);
      return await createOAuthProvider(env).fetch(request, env, ctx);
    }
    
    // If we have a static client registered, map to it
    if (actualClientId && requestedClientId !== actualClientId) {
      console.log(`Using static MCP client for token: ${requestedClientId} -> ${actualClientId}`);
      
      // Create new form data with actual client ID
      const newFormData = new FormData();
      for (const [key, value] of formData.entries()) {
        if (key === 'client_id') {
          newFormData.append(key, actualClientId);
        } else {
          newFormData.append(key, value as string);
        }
      }
      
      // Create a new request with the actual client ID
      const mappedRequest = new Request(request.url, {
        method: request.method,
        headers: request.headers,
        body: newFormData
      });
      
      return await createOAuthProvider(env).fetch(mappedRequest, env, ctx);
    }
    
    // If no mapping needed, proceed normally
    return await createOAuthProvider(env).fetch(request, env, ctx);
    
  } catch (error: any) {
    console.error('Token exchange error:', error);
    return new Response(JSON.stringify({
      error: 'server_error',
      error_description: error.message || 'Token exchange failed'
    }), {
      status: 500,
      headers: { 'Content-Type': 'application/json' }
    });
  }
}

// Handle MCP over SSE for Claude Desktop
async function handleSSEMcp(request: Request, env: Env): Promise<Response> {
  console.log('Setting up SSE MCP connection');
  
  // Create a streaming response for SSE
  const { readable, writable } = new TransformStream();
  const writer = writable.getWriter();
  const encoder = new TextEncoder();
  
  // Start SSE connection
  const sseResponse = new Response(readable, {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Accept, Authorization'
    }
  });
  
  // Handle the MCP connection asynchronously
  handleMCPConnection(writer, encoder, env).catch(console.error);
  
  return sseResponse;
}

// Handle MCP protocol over SSE
async function handleMCPConnection(writer: WritableStreamDefaultWriter, encoder: TextEncoder, env: Env) {
  try {
    console.log('Starting MCP handshake over SSE');
    
    // Wait a moment for client to be ready
    await new Promise(resolve => setTimeout(resolve, 100));
    
    // Send server info (initialize response)
    const serverInfo = {
      jsonrpc: '2.0',
      id: 1,
      result: {
        protocolVersion: '2024-11-05',
        capabilities: {
          tools: {},
          resources: {},
          prompts: {},
          logging: {}
        },
        serverInfo: {
          name: 'Microsoft 365 MCP Server',
          version: '0.3.0'
        }
      }
    };
    
    await writer.write(encoder.encode(`data: ${JSON.stringify(serverInfo)}\n\n`));
    
    // Send tools list
    const toolsList = {
      jsonrpc: '2.0',
      id: 2,
      result: {
        tools: [
          {
            name: 'read_emails',
            description: 'Read emails from Microsoft Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                folder: { type: 'string', description: 'Email folder (inbox, sent, drafts)' },
                limit: { type: 'number', description: 'Number of emails to retrieve' }
              }
            }
          },
          {
            name: 'send_email',
            description: 'Send an email via Microsoft Outlook', 
            inputSchema: {
              type: 'object',
              properties: {
                to: { type: 'string', description: 'Recipient email address' },
                subject: { type: 'string', description: 'Email subject' },
                body: { type: 'string', description: 'Email body' }
              },
              required: ['to', 'subject', 'body']
            }
          }
        ]
      }
    };
    
    await writer.write(encoder.encode(`data: ${JSON.stringify(toolsList)}\n\n`));
    
    console.log('MCP handshake completed - sent server info and tools list');
    
    // Keep connection alive with periodic pings
    const pingInterval = setInterval(async () => {
      try {
        const pingMessage = `data: ${JSON.stringify({ type: 'ping', timestamp: Date.now() })}\n\n`;
        await writer.write(encoder.encode(pingMessage));
      } catch (error) {
        console.error('SSE ping failed:', error);
        clearInterval(pingInterval);
        writer.close();
      }
    }, 30000); // Ping every 30 seconds
    
    // Handle connection cleanup
    setTimeout(() => {
      console.log('Closing MCP SSE connection after timeout');
      clearInterval(pingInterval);
      writer.close();
    }, 300000); // Close after 5 minutes
    
  } catch (error) {
    console.error('MCP SSE connection error:', error);
    writer.close();
  }
}

// Create OAuth provider factory function to capture environment
function createOAuthProvider(env: Env) {
  return new OAuthProvider({
  // Protect API endpoints - token endpoint needs special handling for client ID mapping
  apiRoute: [],  // Don't protect any routes, handle mapping at the defaultHandler level
  apiHandler: { fetch: async () => new Response('Not used', { status: 404 }) }, // Not used

  // Default handler for OAuth flows - route to MicrosoftHandler
  defaultHandler: MicrosoftHandler as any,

  // OAuth endpoints
  authorizeEndpoint: '/authorize',
  tokenEndpoint: '/token',
  clientRegistrationEndpoint: '/register',
  
  // Enable dynamic client registration for Claude Desktop support
  disallowPublicClientRegistration: false,

  // Supported scopes
  scopesSupported: [
    'User.Read',
    'Mail.Read',
    'Mail.ReadWrite',
    'Mail.Send',
    'Calendars.Read',
    'Calendars.ReadWrite',
    'Contacts.ReadWrite',
    'OnlineMeetings.ReadWrite',
    'ChannelMessage.Send',
    'Team.ReadBasic.All',
  ],

  // Token exchange callback - integrates Microsoft tokens into OAuth flow
  tokenExchangeCallback: async (options: any) => {
    // Use captured environment from closure

    if (options.grantType === 'authorization_code') {
      // The MicrosoftHandler has stored the Microsoft authorization code in props
      const microsoftAuthCode = options.props.microsoftAuthCode;
      const redirectUri = options.props.microsoftRedirectUri;

      if (!microsoftAuthCode) {
        throw new Error('No Microsoft authorization code available');
      }

      try {
        // Exchange Microsoft authorization code for access tokens
        const microsoftTokens = await exchangeMicrosoftTokens(microsoftAuthCode, env, redirectUri);

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
      } catch (error) {
        console.error('Microsoft token exchange failed:', error);
        throw error;
      }
    }

    if (options.grantType === 'refresh_token') {
      // Refresh Microsoft tokens using stored refresh token
      const refreshToken = options.props.microsoftRefreshToken;

      if (!refreshToken) {
        throw new Error('No Microsoft refresh token available');
      }

      try {
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
            microsoftRefreshToken: microsoftTokens.refresh_token || refreshToken,
          },
          accessTokenTTL: microsoftTokens.expires_in,
        };
      } catch (error) {
        console.error('Microsoft token refresh failed:', error);
        throw error;
      }
    }

    // For other grant types, return unchanged
    return {};
  },
  });
}

// Unified handler supporting GET, POST, and WebSocket on single /sse endpoint
const unifiedHandler = {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    const url = new URL(request.url);
    
    // DEBUG: Log all requests to see what we're getting
    console.log(`=== REQUEST DEBUG ===`);
    console.log(`Method: ${request.method}`);
    console.log(`URL: ${request.url}`);
    console.log(`Pathname: ${url.pathname}`);
    
    // Route 1: MCP endpoint - hybrid handler for multiple protocols
    if (url.pathname === '/sse' || url.pathname.startsWith('/sse/')) {
      return handleHybridMcp(request, env, ctx);
    }
    
    // Route 2: Health check
    if (url.pathname === '/health') {
      return new Response(JSON.stringify({
        status: 'healthy',
        service: 'Microsoft 365 MCP Server - Unified Endpoint',
        timestamp: new Date().toISOString(),
        architecture: 'Single /sse endpoint with multi-protocol support',
        protocols: {
          'GET': 'SSE validation for Claude Desktop',
          'POST': 'Direct JSON-RPC for testing',
          'WebSocket': 'Full MCP protocol with OAuth'
        },
        endpoints: {
          'mcp-server': '/sse (all protocols)',
          'health': '/health',
          'authorization': '/authorize'
        },
        documentation: 'https://github.com/nikolanovoselec/m365-mcp-server'
      }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }
    
    // Route 3: Service info
    if (url.pathname === '/' || url.pathname === '/info') {
      return new Response(JSON.stringify({
        service: 'Microsoft 365 MCP Server',
        version: '0.3.0',
        architecture: 'Unified Endpoint',
        configurations: {
          claude_desktop_direct: {
            url: `${env.PROTOCOL}://${env.WORKER_DOMAIN}/sse`,
            description: 'Direct web connector'
          },
          mcp_remote: {
            command: 'npx',
            args: ['mcp-remote', `${env.PROTOCOL}://${env.WORKER_DOMAIN}/sse`],
            description: 'Traditional MCP-remote configuration'
          }
        },
        note: 'Both configurations use the same unified endpoint with automatic protocol detection',
        documentation: 'https://github.com/nikolanovoselec/m365-mcp-server'
      }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }
    
    // Route 4: OAuth Server Metadata Discovery (for Claude Desktop native remote MCP)
    if (url.pathname === '/.well-known/oauth-authorization-server') {
      console.log('OAuth server metadata discovery request');
      return handleOAuthMetadata(request, env);
    }
    
    // Route 5: All other OAuth routes - delegate to OAuth provider (client initialization happens in handlers)
    console.log(`OAuth route: ${url.pathname}`);
    
    const oauthProvider = createOAuthProvider(env);
    return oauthProvider.fetch(request, env, ctx);
  }
};

// Handle OAuth Server Metadata Discovery (RFC8414) for Claude Desktop
async function handleOAuthMetadata(request: Request, env: Env): Promise<Response> {
  const url = new URL(request.url);
  const baseUrl = `${url.protocol}//${url.host}`;
  
  const metadata = {
    issuer: baseUrl,
    authorization_endpoint: `${baseUrl}/authorize`,
    token_endpoint: `${baseUrl}/token`, 
    registration_endpoint: `${baseUrl}/register`,
    grant_types_supported: [
      "authorization_code",
      "refresh_token"
    ],
    response_types_supported: [
      "code"
    ],
    code_challenge_methods_supported: [
      "S256"
    ],
    token_endpoint_auth_methods_supported: [
      "none"
    ],
    scopes_supported: [
      "claudeai"
    ]
  };
  
  return new Response(JSON.stringify(metadata, null, 2), {
    headers: {
      'Content-Type': 'application/json',
      'Cache-Control': 'public, max-age=3600',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET',
      'Access-Control-Allow-Headers': 'Content-Type'
    }
  });
}

// Handle unified MCP endpoint with GET, POST, and WebSocket support
async function handleUnifiedMcp(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
  
  // DEBUG: Log all headers to understand what we're receiving
  const allHeaders = Object.fromEntries(request.headers.entries());
  console.log('Request details:', {
    method: request.method,
    url: request.url,
    headers: allHeaders
  });
  
  // Protocol 1: WebSocket upgrade requests (mcp-remote and OAuth-capable clients)
  // Check this FIRST because WebSocket upgrades are also GET requests
  const upgradeHeader = request.headers.get('Upgrade');
  const connectionHeader = request.headers.get('Connection');
  const webSocketKey = request.headers.get('Sec-WebSocket-Key');
  const webSocketVersion = request.headers.get('Sec-WebSocket-Version');
  
  console.log('WebSocket check:', { upgradeHeader, connectionHeader, webSocketKey, webSocketVersion });
  
  // All WebSocket requests: delegate to OAuth provider (handles WebSocket natively)
  if ((upgradeHeader && upgradeHeader.toLowerCase() === 'websocket') ||
      (connectionHeader && connectionHeader.toLowerCase().includes('upgrade')) ||
      (webSocketKey && webSocketVersion)) {
    console.log('WebSocket upgrade detected - delegating to OAuth provider with native WebSocket support');
    const oauthProvider = createOAuthProvider(env);
    return oauthProvider.fetch(request, env, ctx);
  }
  
  // Protocol 2: GET requests with SSE headers
  if (request.method === 'GET') {
    const acceptHeader = request.headers.get('Accept');
    if (acceptHeader && acceptHeader.includes('text/event-stream')) {
      // All SSE requests get the same validation response
      // mcp-remote will then switch to WebSocket for the actual MCP protocol
      console.log('SSE validation request (Claude Desktop or mcp-remote)');
      return handleSseValidation(request, env);
    }
    return new Response('GET requests must include Accept: text/event-stream for SSE or Upgrade: websocket', { status: 400 });
  }
  
  // Protocol 3: POST requests (Direct JSON-RPC for testing/debugging)
  if (request.method === 'POST') {
    return handleDirectJsonRpc(request, env);
  }
  
  return new Response('Unsupported method - Use GET (SSE), POST (JSON-RPC), or WebSocket upgrade', { status: 405 });
}

// Handle SSE validation for Claude Desktop web connector
async function handleSseValidation(request: Request, env: Env): Promise<Response> {
  console.log('Claude Desktop SSE validation - returning SSE headers');
  
  return new Response('', {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Accept, Upgrade, Connection'
    }
  });
}

// Handle direct JSON-RPC requests for testing
async function handleDirectJsonRpc(request: Request, env: Env): Promise<Response> {
  try {
    const body = await request.text();
    const message = JSON.parse(body);
    
    console.log(`Received POST method: ${message.method}`);
    
    if (!message.method) {
      return new Response(JSON.stringify({
        jsonrpc: '2.0',
        id: message.id || null,
        error: { code: -32600, message: 'Invalid request - missing method' }
      }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' }
      });
    }
    
    // Allow MCP handshake methods without authentication - using slash notation for Claude Desktop compatibility
    const handshakeMethods = ['initialize', 'initialized', 'tools/list', 'resources/list', 'prompts/list', 'notifications/initialized', 'notifications/cancelled'];
    
    if (handshakeMethods.includes(message.method)) {
      console.log(`Direct JSON-RPC handshake: ${message.method}`);
      
      switch (message.method) {
        case 'initialize':
          return new Response(JSON.stringify({
            jsonrpc: '2.0',
            id: message.id,
            result: {
              protocolVersion: '2024-11-05',
              serverInfo: {
                name: 'microsoft-365-mcp',
                version: '0.3.0'
              },
              capabilities: {
                tools: {},
                resources: {},
                prompts: {}
              }
            }
          }), {
            headers: { 'Content-Type': 'application/json' }
          });
          
        case 'list_tools':
          return new Response(JSON.stringify({
            jsonrpc: '2.0',
            id: message.id,
            result: {
              tools: [
                {
                  name: 'sendEmail',
                  description: 'Send an email via Outlook',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      to: { type: 'string', description: 'Recipient email address' },
                      subject: { type: 'string', description: 'Email subject' },
                      body: { type: 'string', description: 'Email body content' },
                      contentType: { type: 'string', enum: ['text', 'html'], default: 'html' }
                    },
                    required: ['to', 'subject', 'body']
                  }
                },
                {
                  name: 'getEmails',
                  description: 'Get recent emails',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      count: { type: 'number', maximum: 50, default: 10 },
                      folder: { type: 'string', default: 'inbox' }
                    }
                  }
                },
                {
                  name: 'getCalendarEvents', 
                  description: 'Get calendar events',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      days: { type: 'number', maximum: 30, default: 7 }
                    }
                  }
                },
                {
                  name: 'authenticate',
                  description: 'Get authentication URL for Microsoft 365',
                  inputSchema: {
                    type: 'object',
                    properties: {}
                  }
                }
              ]
            }
          }), {
            headers: { 'Content-Type': 'application/json' }
          });
          

        case 'tools/list':
          return new Response(JSON.stringify({
            jsonrpc: '2.0',
            id: message.id,
            result: {
              tools: [
                {
                  name: 'sendEmail',
                  description: 'Send an email via Outlook',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      to: { type: 'string', description: 'Recipient email address' },
                      subject: { type: 'string', description: 'Email subject' },
                      body: { type: 'string', description: 'Email body content' },
                      contentType: { type: 'string', enum: ['text', 'html'], default: 'html' }
                    },
                    required: ['to', 'subject', 'body']
                  }
                },
                {
                  name: 'getEmails',
                  description: 'Get recent emails',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      count: { type: 'number', maximum: 50, default: 10 },
                      folder: { type: 'string', default: 'inbox' }
                    }
                  }
                },
                {
                  name: 'searchEmails',
                  description: 'Search emails',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      query: { type: 'string', description: 'Search query' },
                      count: { type: 'number', maximum: 50, default: 10, description: 'Number of results' }
                    },
                    required: ['query']
                  }
                },
                {
                  name: 'getCalendarEvents',
                  description: 'Get calendar events',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      days: { type: 'number', maximum: 30, default: 7 }
                    }
                  }
                },
                {
                  name: 'createCalendarEvent',
                  description: 'Create calendar event',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      subject: { type: 'string', description: 'Event title' },
                      start: { type: 'string', description: 'Start date/time (ISO format)' },
                      end: { type: 'string', description: 'End date/time (ISO format)' },
                      body: { type: 'string', description: 'Event description' },
                      attendees: { type: 'string', description: 'Comma-separated email addresses' }
                    },
                    required: ['subject', 'start', 'end']
                  }
                },
                {
                  name: 'sendTeamsMessage',
                  description: 'Send Teams message',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      chatId: { type: 'string', description: 'Teams chat ID' },
                      message: { type: 'string', description: 'Message content' }
                    },
                    required: ['chatId', 'message']
                  }
                },
                {
                  name: 'createTeamsMeeting',
                  description: 'Create Teams meeting',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      subject: { type: 'string', description: 'Meeting title' },
                      start: { type: 'string', description: 'Start date/time (ISO format)' },
                      end: { type: 'string', description: 'End date/time (ISO format)' },
                      attendees: { type: 'string', description: 'Comma-separated email addresses' }
                    },
                    required: ['subject', 'start', 'end']
                  }
                },
                {
                  name: 'getContacts',
                  description: 'Get contacts',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      count: { type: 'number', maximum: 50, default: 10, description: 'Number of contacts' }
                    }
                  }
                },
                {
                  name: 'authenticate',
                  description: 'Get authentication URL for Microsoft 365',
                  inputSchema: {
                    type: 'object',
                    properties: {}
                  }
                }
              ]
            }
          }), {
            headers: { 'Content-Type': 'application/json' }
          });

        case 'resources/list':
        case 'prompts/list':
          return new Response(JSON.stringify({
            jsonrpc: '2.0',
            id: message.id,
            result: { [message.method.split('/')[0]]: [] }
          }), {
            headers: { 'Content-Type': 'application/json' }
          });

        case 'initialized':
        case 'notifications/initialized':
        case 'notifications/cancelled':
          // These are notifications that don't require responses
          return new Response('', { status: 200 });
          
        default:
          return new Response(JSON.stringify({
            jsonrpc: '2.0',
            id: message.id,
            error: { code: -32601, message: 'Method not found' }
          }), {
            status: 404,
            headers: { 'Content-Type': 'application/json' }
          });
      }
    }
    
    // Tool calls require authentication
    if (message.method === 'tools/call') {
      const host = request.headers.get('host') || env.WORKER_DOMAIN;
      
      return new Response(JSON.stringify({
        jsonrpc: '2.0',
        id: message.id,
        error: {
          code: -32001,
          message: 'Authentication required',
          data: {
            error: 'microsoft_oauth_required',
            auth_url: `https://${host}/authorize`,
            instructions: [
              'Microsoft 365 authentication is required to use tools.',
              'Please visit the auth_url to complete OAuth authentication.',
              'After authentication, tools will be available for use.'
            ]
          }
        }
      }), {
        status: 401,
        headers: { 'Content-Type': 'application/json' }
      });
    }
    
    // Other methods
    return new Response(JSON.stringify({
      jsonrpc: '2.0',
      id: message.id,
      error: { code: -32601, message: 'Method not supported in direct mode' }
    }), {
      status: 404,
      headers: { 'Content-Type': 'application/json' }
    });
    
  } catch (error) {
    console.error('JSON-RPC parsing error:', error);
    
    return new Response(JSON.stringify({
      jsonrpc: '2.0',
      id: null,
      error: { 
        code: -32700, 
        message: 'Parse error',
        data: { details: error instanceof Error ? error.message : 'Unknown error' }
      }
    }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' }
    });
  }
}


export default unifiedHandler;

// Export Durable Object classes
export { MicrosoftMCPAgent } from './microsoft-mcp-agent';