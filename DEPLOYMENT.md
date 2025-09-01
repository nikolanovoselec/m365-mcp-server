# Deployment Guide

Complete deployment and development setup for the Microsoft 365 MCP Server on Cloudflare Workers.

## Table of Contents

- [Prerequisites](#prerequisites)
  - [Required Accounts & Services](#required-accounts--services)
  - [Install Wrangler CLI](#install-wrangler-cli)
- [Step 1: Cloudflare Setup](#step-1-cloudflare-setup)
  - [Authentication](#authentication)
  - [Verify Authentication](#verify-authentication)
- [Step 2: Project Setup](#step-2-project-setup)
  - [Clone Repository](#clone-repository)
  - [Environment Configuration](#environment-configuration)
- [Step 3: Microsoft Entra ID Configuration](#step-3-microsoft-entra-id-configuration)
  - [Create Application Registration](#create-application-registration)
  - [Configure Redirect URIs](#configure-redirect-uris)
  - [Configure API Permissions](#configure-api-permissions)
  - [Create Client Secret](#create-client-secret)
- [Step 4: Deploy Secrets](#step-4-deploy-secrets)
- [Step 5: Deploy to Production](#step-5-deploy-to-production)
- [Step 6: Verification](#step-6-verification)
- [Development Workflow](#development-workflow)
  - [Local Development](#local-development)
  - [Testing](#testing)
- [Architecture & Dependencies](#architecture--dependencies)
  - [Core Libraries](#core-libraries)
  - [Manual Implementation](#manual-implementation)
- [Troubleshooting](#troubleshooting)
  - [Common Deployment Issues](#common-deployment-issues)
  - [Cloudflare WAF Issues](#cloudflare-waf-issues)
  - [Production Checklist](#production-checklist)
- [Performance Optimization](#performance-optimization)
  - [Cloudflare Workers Optimization](#cloudflare-workers-optimization)
  - [Response Caching](#response-caching)
  - [Monitoring](#monitoring)
- [Support](#support)
- [Technical Deep Dive](#technical-deep-dive)
  - [1. Unified Endpoint Architecture](#1-unified-endpoint-architecture)
  - [2. OAuth 2.1 + PKCE + Dynamic Client Registration Flow](#2-oauth-21--pkce--dynamic-client-registration-flow)
  - [3. Library vs Manual Implementation Analysis](#3-library-vs-manual-implementation-analysis)
  - [4. Durable Objects + KV Storage Architecture](#4-durable-objects--kv-storage-architecture)
  - [5. Microsoft Graph API Integration Patterns](#5-microsoft-graph-api-integration-patterns)
  - [6. Client Detection & Protocol Adaptation](#6-client-detection--protocol-adaptation)
  - [7. Security Implementation Deep Dive](#7-security-implementation-deep-dive)
  - [8. WebSocket Implementation & HTTP/2 Challenges](#8-websocket-implementation--http2-challenges)
  - [9. Environment & Secrets Management Architecture](#9-environment--secrets-management-architecture)
  - [10. Libraries & SDK Integration Analysis](#10-libraries--sdk-integration-analysis)
  - [11. MCP Protocol Implementation Patterns](#11-mcp-protocol-implementation-patterns)
  - [12. Production Deployment Patterns](#12-production-deployment-patterns)

## Prerequisites

### Required Accounts & Services

1. **Cloudflare Account**
   - Free tier sufficient for development
   - Workers Paid plan recommended for production ($5/month)
   - Sign up at [cloudflare.com](https://dash.cloudflare.com/sign-up)

2. **Microsoft Entra ID (Azure AD)**
   - Free tier sufficient
   - Application registration access
   - Sign up at [portal.azure.com](https://portal.azure.com)

3. **Development Environment**
   - Node.js 18.0+ and npm
   - Git version control
   - Text editor or IDE

### Install Wrangler CLI

```bash
# Install globally via npm
npm install -g wrangler

# Verify installation
wrangler --version

# Alternative: Use npx (no global install)
npx wrangler --version
```

## Step 1: Cloudflare Setup

### Authentication

**Interactive Login (Recommended):**
```bash
wrangler login
# Opens browser for authentication
```

**API Token Method:**
1. Visit [Cloudflare Dashboard > API Tokens](https://dash.cloudflare.com/profile/api-tokens)
2. Create token with permissions:
   - Zone:Zone:Read
   - Zone:Zone Settings:Edit  
   - Account:Cloudflare Workers:Edit
   - Account:Account Settings:Read
3. Configure: `export CLOUDFLARE_API_TOKEN="your_token"`

### Verify Authentication

```bash
wrangler auth whoami
# Should show your account details
```

## Step 2: Project Setup

### Clone Repository

```bash
git clone https://github.com/nikolanovoselec/m365-mcp-server
cd m365-mcp-server
npm install
```

### Environment Configuration

**Production (wrangler.toml):**
```bash
# Copy template
cp wrangler.example.toml wrangler.toml
```

Edit `wrangler.toml` with your values:
```toml
account_id = "YOUR_CLOUDFLARE_ACCOUNT_ID"

[vars]
MICROSOFT_CLIENT_ID = "YOUR_ENTRA_ID_CLIENT_ID"
MICROSOFT_TENANT_ID = "YOUR_ENTRA_ID_TENANT_ID"
GRAPH_API_VERSION = "v1.0"
WORKER_DOMAIN = "your-worker.your-subdomain.workers.dev"
PROTOCOL = "https"
```

**Local Development (.dev.vars):**
```bash
# Copy template
cp .dev.vars.example .dev.vars
```

Edit `.dev.vars` with your values:
```bash
MICROSOFT_CLIENT_ID=your_entra_id_client_id
MICROSOFT_TENANT_ID=your_entra_id_tenant_id
MICROSOFT_CLIENT_SECRET=your_entra_id_client_secret
GRAPH_API_VERSION=v1.0
WORKER_DOMAIN=localhost:8787
PROTOCOL=http
ENCRYPTION_KEY=your_32_byte_hex_key
COOKIE_SECRET=your_32_byte_hex_key
```

## Step 3: Microsoft Entra ID Configuration

### Create Application Registration

1. Go to [Azure Portal > App Registrations](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps)
2. Click "New registration"
3. Configure:
   - **Name**: Microsoft 365 MCP Server
   - **Supported account types**: Accounts in any organizational directory and personal Microsoft accounts
   - **Redirect URI**: Web - `https://your-domain.com/callback`

### Configure Redirect URIs

Add these redirect URIs in Authentication settings:
```
https://your-worker-domain.com/callback
https://your-favorite-ai.com/api/mcp/auth_callback
```

### Configure API Permissions

In API permissions, add these Microsoft Graph permissions:
- `User.Read` (delegated)
- `Mail.Read` (delegated)
- `Mail.ReadWrite` (delegated)
- `Mail.Send` (delegated)
- `Calendars.Read` (delegated)
- `Calendars.ReadWrite` (delegated)
- `Contacts.ReadWrite` (delegated)
- `OnlineMeetings.ReadWrite` (delegated)
- `ChannelMessage.Send` (delegated)
- `Team.ReadBasic.All` (delegated)
- `offline_access` (delegated)

Grant admin consent for all permissions.

### Create Client Secret

1. Go to "Certificates & secrets"
2. Click "New client secret"
3. Add description and set expiration
4. Copy the secret value (you won't see it again)

## Step 4: Deploy Secrets

Deploy sensitive values as Cloudflare Worker secrets:

```bash
# Microsoft client secret
npx wrangler secret put MICROSOFT_CLIENT_SECRET
# Enter: your_client_secret_from_azure

# Encryption key (generate: openssl rand -hex 32)
npx wrangler secret put ENCRYPTION_KEY
# Enter: your_32_byte_hex_encryption_key

# Cookie secret (generate: openssl rand -hex 32)
npx wrangler secret put COOKIE_SECRET
# Enter: your_32_byte_hex_cookie_secret
```

## Step 5: Deploy to Production

```bash
# Deploy worker
npx wrangler deploy

# Expected output:
# ✨ Successfully published your Worker to your-worker.your-subdomain.workers.dev
```

## Step 6: Verification

Test your deployment:

```bash
# Health check
curl https://your-worker-domain.com/health

# OAuth metadata
curl https://your-worker-domain.com/.well-known/oauth-authorization-server

# MCP tool discovery
curl -X POST https://your-worker-domain.com/sse \
  -H 'Content-Type: application/json' \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}'
```

## Development Workflow

### Local Development

```bash
# Start local development server
npx wrangler dev

# Server runs at http://localhost:8787
# Test with: curl http://localhost:8787/health
```

### Testing

```bash
# Type checking
npm run type-check

# Build verification  
npm run build

# Lint (if configured)
npm run lint
```

## Architecture & Dependencies

### Core Libraries

**@cloudflare/workers-oauth-provider (v0.0.6)**
- Complete OAuth 2.1 + PKCE server implementation
- Handles protocol compliance and security validation
- Automatic token generation and PKCE challenge validation

**agents (v0.0.113)**
- Cloudflare Workers AI agents framework
- Provides MCP protocol utilities and Durable Objects integration
- Session management and WebSocket handling

**@modelcontextprotocol/sdk (v1.17.4)**
- Official MCP SDK for TypeScript
- Tool schemas and JSON-RPC type definitions
- Protocol compliance utilities

**hono (v4.9.5)**
- Fast web framework for Cloudflare Workers
- Request routing and middleware support
- TypeScript-first design

**zod (v3.22.4)**
- Schema validation for request/response data
- Type-safe input validation for MCP tools
- Error handling with descriptive messages

### Manual Implementation

**60% of functionality manually implemented:**
- Microsoft Graph API integration
- MCP tool handlers and business logic
- Client detection and protocol routing
- Token encryption and session management
- Error handling and user experience

**40% provided by libraries:**
- OAuth protocol implementation
- WebSocket and networking
- Type definitions and validation
- Basic MCP protocol compliance

## Troubleshooting

### Common Deployment Issues

**Authentication Errors:**
```bash
# Check Wrangler auth status
wrangler auth whoami

# Re-authenticate if needed
wrangler logout
wrangler login
```

**Environment Variable Issues:**
```bash
# Check deployed secrets
wrangler secret list

# Update secrets if needed
wrangler secret put SECRET_NAME
```

**Build Failures:**
```bash
# Clear node_modules and reinstall
rm -rf node_modules package-lock.json
npm install

# Check TypeScript compilation
npm run type-check
```

### Cloudflare WAF Issues

If OAuth callbacks are blocked by Cloudflare's Web Application Firewall:

1. **Check Security Events:**
   - Go to Cloudflare Dashboard > Security > Events
   - Look for blocked requests to `/callback`

2. **Create WAF Bypass Rules:**
   - Go to Security > WAF > Custom Rules
   - Create rule: Skip Managed Rules for OAuth endpoints
   - Expression: `(http.request.uri.path eq "/callback")`
   - Action: Skip > Managed Rules

3. **Test OAuth Flow:**
   ```bash
   curl -v "https://your-domain.com/callback?code=test&state=test"
   # Should return 200 instead of 403
   ```

### Production Checklist

**Before deploying to production:**
- [ ] All environment variables configured
- [ ] Secrets deployed to Cloudflare
- [ ] Microsoft Entra ID permissions granted
- [ ] Redirect URIs match production domain
- [ ] Health check returns 200 OK
- [ ] OAuth metadata endpoint accessible
- [ ] Tool discovery works without authentication

**Security validation:**
- [ ] No sensitive data in source code
- [ ] All tokens encrypted at rest
- [ ] HTTPS enforced for all endpoints
- [ ] CORS policies properly configured
- [ ] Rate limiting enabled (Cloudflare default)

## Performance Optimization

### Cloudflare Workers Optimization

```toml
# In wrangler.toml - optimize for performance
[build]
compatibility_date = "2024-11-05"
compatibility_flags = ["nodejs_compat"]

# Enable KV caching
[[kv_namespaces]]
binding = "CACHE_KV"
```

### Response Caching

```typescript
// Cache frequently accessed data
const cacheKey = `emails-${userId}-${folder}`;
const cached = await env.CACHE_KV.get(cacheKey);
if (cached) {
  return JSON.parse(cached);
}

// Fetch fresh data and cache
const data = await fetchFromGraphAPI();
await env.CACHE_KV.put(cacheKey, JSON.stringify(data), { 
  expirationTtl: 300 // 5 minutes
});
```

### Monitoring

**Cloudflare Analytics:**
- Monitor request volume and latency
- Track error rates and response codes  
- Set up alerts for high error rates

**Custom Logging:**
```typescript
// Structured logging in Workers
console.log(JSON.stringify({
  timestamp: new Date().toISOString(),
  level: 'INFO',
  message: 'OAuth callback processed',
  userId: session.userId,
  clientType: 'claude-desktop'
}));
```

## Support

For deployment issues:
1. Check [Cloudflare Workers documentation](https://developers.cloudflare.com/workers/)
2. Review [Microsoft Graph API docs](https://docs.microsoft.com/en-us/graph/)
3. Report issues at [GitHub repository](https://github.com/nikolanovoselec/m365-mcp-server/issues)

## Technical Deep Dive

### 1. Unified Endpoint Architecture

The MCP server implements a sophisticated single-endpoint design that handles multiple protocols through intelligent request analysis.

**Protocol Detection Logic:**
```typescript
// Single /sse endpoint supporting multiple protocols
if (request.headers.get('upgrade') === 'websocket') {
  return handleWebSocketUpgrade(request);
}

if (request.headers.get('accept') === 'text/event-stream') {
  return handleSSEConnection(request);
}

if (request.method === 'POST') {
  return handleJSONRPC(request);
}
```

**Supported Protocols:**
- **WebSocket**: Full bidirectional MCP protocol for mcp-remote
- **Server-Sent Events**: Claude Desktop validation and streaming
- **JSON-RPC over HTTP**: Direct tool testing and debugging

**Client Fingerprinting:**
```typescript
// Detect client type based on headers and redirect URIs
const clientType = detectClientType({
  userAgent: request.headers.get('user-agent'),
  acceptHeader: request.headers.get('accept'),
  redirectUri: oauthContext?.redirectUri,
  upgradeHeader: request.headers.get('upgrade')
});
```

### 2. OAuth 2.1 + PKCE + Dynamic Client Registration Flow

**Complete OAuth Server Implementation:**

**Dynamic Client Registration for Claude Desktop:**
```typescript
// Auto-generate client IDs for Claude Desktop web connectors
const registerRequest = new Request('/register', {
  method: 'POST',
  body: JSON.stringify({
    client_name: 'Microsoft 365 MCP Static Client',
    redirect_uris: ['https://your-ai-platform.com/api/mcp/auth_callback'],
    grant_types: ['authorization_code'],
    response_types: ['code'],
    token_endpoint_auth_method: 'none'
  })
});
```

**Static Client Mapping for mcp-remote:**
```typescript
// Pre-registered client ID for mcp-remote compatibility
const MCP_CLIENT_ID = 'rWJu8WV42zC5pfGT';

// Client ID aliasing for backward compatibility
async function handleAuthorizeWithClientMapping(request) {
  const requestedClientId = url.searchParams.get('client_id');
  const actualClientId = await env.CONFIG_KV.get(`static_client_actual:${MCP_CLIENT_ID}`);
  
  // Rewrite client_id in request for internal processing
  const mappedUrl = new URL(request.url);
  mappedUrl.searchParams.set('client_id', actualClientId);
  return processOAuthRequest(mappedUrl);
}
```

**PKCE Challenge/Verification:**
- Implements RFC 7636 PKCE with S256 method
- Code challenge generation and verification
- State parameter encryption for security

**Token Exchange Callback:**
```typescript
tokenExchangeCallback: async (options) => {
  // Bridge OAuth provider tokens to Microsoft Graph tokens
  const microsoftTokens = await exchangeMicrosoftTokens(
    options.props.microsoftAuthCode, 
    env, 
    redirectUri
  );
  
  return {
    accessTokenProps: {
      microsoftAccessToken: microsoftTokens.access_token,
      microsoftTokenType: microsoftTokens.token_type,
    },
    newProps: {
      microsoftRefreshToken: microsoftTokens.refresh_token,
    },
    accessTokenTTL: microsoftTokens.expires_in,
  };
}
```

### 3. Library vs Manual Implementation Analysis

**Production Dependencies Breakdown:**

**@cloudflare/workers-oauth-provider (v0.0.6)**
- **What it provides**: Complete OAuth 2.1 + PKCE server implementation
- **Automatic features**: Protocol compliance, token validation, client registration
- **Integration pattern**: Custom tokenExchangeCallback for Microsoft Graph bridging

**agents (v0.0.113)**  
- **What it provides**: Cloudflare Workers MCP agent framework
- **Automatic features**: Durable Objects integration, WebSocket utilities, session management
- **Integration pattern**: Extended McpAgent class with Microsoft Graph client

**@modelcontextprotocol/sdk (v1.17.4)**
- **What it provides**: Official MCP protocol types and interfaces
- **Automatic features**: JSON-RPC type safety, tool schema validation
- **Integration pattern**: Server instance with custom tool handlers

**hono (v4.9.5)**
- **What it provides**: Fast web framework for Cloudflare Workers
- **Automatic features**: Request routing, middleware support, context handling
- **Integration pattern**: OAuth handler routes with environment binding

**zod (v3.22.4)**
- **What it provides**: TypeScript-first schema validation
- **Automatic features**: Input validation, type inference, error messages
- **Integration pattern**: Tool parameter validation and sanitization

**Implementation Distribution:**
- **Library-Provided (40%)**: OAuth protocol, WebSocket handling, type definitions, request routing
- **Manual Implementation (60%)**: Microsoft Graph integration, client detection, protocol adaptation, session management, security patterns

**Key Manual Implementations:**
```typescript
// Custom Microsoft Graph token exchange
async function exchangeMicrosoftTokens(authCode, env, redirectUri) {
  const tokenUrl = `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`;
  // Custom token exchange logic with error handling
}

// Multi-protocol endpoint routing
async function handleHybridMcp(request, env, ctx) {
  // Custom protocol detection and routing logic
}

// Client-specific authentication handling
async function handleConditionalAuth(clientType, request) {
  // Custom authentication logic per client type
}
```

### 4. Durable Objects + KV Storage Architecture

**Durable Objects Design:**
```typescript
export class MicrosoftMCPAgent extends McpAgent<Env, State, Props> {
  // Persistent MCP sessions with state management
  constructor(ctx: DurableObjectState, env: Env) {
    super(ctx, env);
    this.graphClient = new MicrosoftGraphClient(env);
  }
}
```

**Three-Tier KV Architecture:**

**OAUTH_KV Namespace:**
```typescript
// OAuth client registration and token storage
await env.OAUTH_KV.put(`client:${clientId}`, JSON.stringify({
  clientId,
  clientName,
  redirectUris,
  tokenEndpointAuthMethod: 'none',
  grantTypes: ['authorization_code', 'refresh_token']
}));
```

**CONFIG_KV Namespace:**
```typescript
// Static client mappings and configuration
await env.CONFIG_KV.put(`static_client_actual:${MCP_CLIENT_ID}`, actualClientId);
```

**CACHE_KV Namespace:**
```typescript  
// Response caching for Microsoft Graph API calls
const cacheKey = `emails-${userId}-${folder}`;
await env.CACHE_KV.put(cacheKey, JSON.stringify(data), { 
  expirationTtl: 300 
});
```

**Session Isolation:**
```typescript
// Unique session per WebSocket connection
const sessionId = crypto.randomUUID();
const id = env.MICROSOFT_MCP_AGENT.idFromName(`ws-${sessionId}`);
const stub = env.MICROSOFT_MCP_AGENT.get(id);
```

### 5. Microsoft Graph API Integration Patterns

**Token Refresh Automation:**
```typescript
async function refreshMicrosoftTokens(refreshToken, env) {
  const tokenUrl = `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`;
  
  const params = new URLSearchParams({
    client_id: env.MICROSOFT_CLIENT_ID,
    client_secret: env.MICROSOFT_CLIENT_SECRET,
    refresh_token: refreshToken,
    grant_type: 'refresh_token',
    scope: 'User.Read Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite...'
  });
  
  // Automatic token refresh with error handling
  const response = await fetch(tokenUrl, { method: 'POST', body: params });
  return response.json();
}
```

**Graph API Endpoint Mapping:**
```typescript
const toolEndpointMapping = {
  'getEmails': 'GET https://graph.microsoft.com/v1.0/me/messages',
  'sendEmail': 'POST https://graph.microsoft.com/v1.0/me/sendMail',  
  'createCalendarEvent': 'POST https://graph.microsoft.com/v1.0/me/events',
  'getCalendarEvents': 'GET https://graph.microsoft.com/v1.0/me/events',
  'sendTeamsMessage': 'POST https://graph.microsoft.com/v1.0/teams/{id}/channels/{id}/messages',
  'createTeamsMeeting': 'POST https://graph.microsoft.com/v1.0/me/onlineMeetings',
  'getContacts': 'GET https://graph.microsoft.com/v1.0/me/contacts',
  'searchEmails': 'GET https://graph.microsoft.com/v1.0/me/messages?$search="query"'
};
```

**Scope Management:**
```typescript
const requiredScopes = [
  'User.Read',
  'Mail.Read', 'Mail.ReadWrite', 'Mail.Send',
  'Calendars.Read', 'Calendars.ReadWrite', 
  'Contacts.ReadWrite',
  'OnlineMeetings.ReadWrite',
  'ChannelMessage.Send', 'Team.ReadBasic.All',
  'offline_access'
];
```

### 6. Client Detection & Protocol Adaptation

**Smart Client Fingerprinting:**
```typescript
function detectClientType(signals) {
  // AI platform detection (e.g., Claude, ChatGPT, etc.)
  if (signals.redirectUri?.includes('your-ai-platform.com')) {
    return 'ai-platform';
  }
  
  // mcp-remote detection via other signals
  if (signals.userAgent?.includes('mcp-remote') || 
      signals.upgradeHeader === 'websocket') {
    return 'mcp-remote';
  }
  
  // WebSocket client detection
  if (signals.upgradeHeader === 'websocket') {
    return 'websocket-client';
  }
  
  return 'unknown';
}
```

**Conditional Authentication:**
```typescript
// Handshake methods work without authentication
const handshakeMethods = [
  'initialize', 'initialized', 'tools/list', 'resources/list', 
  'prompts/list', 'notifications/initialized'
];

if (handshakeMethods.includes(message.method)) {
  // Allow unauthenticated access for discovery
  return handleHandshake(message);
}

// Tool calls require authentication
if (message.method === 'tools/call') {
  return requireAuthentication(message);
}
```

**Adaptive Response Formatting:**
```typescript
// Client-specific response formatting
switch (clientType) {
  case 'claude-desktop':
    return formatForSSE(response);
  case 'mcp-remote':  
    return formatForWebSocket(response);
  default:
    return formatForJSONRPC(response);
}
```

### 7. Security Implementation Deep Dive

**HMAC-Signed Approval Cookies:**
```typescript
async function signApprovalCookie(approvedClients, secret) {
  const payload = JSON.stringify(approvedClients);
  const key = await crypto.subtle.importKey(
    'raw', 
    new TextEncoder().encode(secret),
    { hash: 'SHA-256', name: 'HMAC' },
    false,
    ['sign']
  );
  
  const signature = await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(payload));
  const signatureHex = Array.from(new Uint8Array(signature))
    .map(b => b.toString(16).padStart(2, '0'))
    .join('');
    
  return `${signatureHex}.${btoa(payload)}`;
}
```

**HTML Sanitization:**
```typescript
function sanitizeHtml(unsafe) {
  return unsafe
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
```

**Token Encryption:**
```typescript
// All tokens encrypted at rest using Web Crypto API
const encrypted = await crypto.subtle.encrypt(
  { name: 'AES-GCM', iv: iv },
  key,
  new TextEncoder().encode(tokenData)
);
```

**Cookie Security Attributes:**
```typescript
const secureCookie = `${COOKIE_NAME}=${value}; HttpOnly; Secure; Path=/; SameSite=Lax; Max-Age=${ONE_YEAR_IN_SECONDS}`;
```

### 8. WebSocket Implementation & HTTP/2 Challenges

**WebSocketPair Pattern:**
```typescript
// Cloudflare Workers WebSocket implementation
if (upgradeHeader?.toLowerCase() === 'websocket' && webSocketKey) {
  const webSocketPair = new WebSocketPair();
  const [client, server] = Object.values(webSocketPair);
  
  server.accept();
  
  server.addEventListener('message', async (event) => {
    const message = JSON.parse(event.data);
    const response = await handleMCPMessage(message);
    server.send(JSON.stringify(response));
  });
  
  return new Response(null, { status: 101, webSocket: client });
}
```

**WebSocket Accept Key Generation (RFC 6455):**
```typescript
async function generateWebSocketAccept(webSocketKey) {
  const webSocketMagicString = '258EAFA5-E914-47DA-95CA-C5AB0DC85B11';
  const concatenated = webSocketKey + webSocketMagicString;
  
  const hashBuffer = await crypto.subtle.digest('SHA-1', new TextEncoder().encode(concatenated));
  const hashArray = new Uint8Array(hashBuffer);
  
  let binary = '';
  for (let i = 0; i < hashArray.length; i++) {
    binary += String.fromCharCode(hashArray[i]);
  }
  return btoa(binary);
}
```

**HTTP/2 Compatibility Issues:**
- Cloudflare serves content over HTTP/2 by default
- WebSocket upgrade headers are invalid in HTTP/2
- 'Upgrade: websocket' headers get dropped during protocol conversion
- **Solution**: Detect WebSocket intent via multiple header analysis

**Fallback Strategy:**
```typescript
// Multi-signal WebSocket detection
const isWebSocketRequest = (
  upgradeHeader?.toLowerCase() === 'websocket' ||
  webSocketKey && webSocketVersion ||
  connectionHeader?.toLowerCase().includes('upgrade')
);
```

### 9. Environment & Secrets Management Architecture

**Three-Tier Configuration:**

**Public Configuration (wrangler.toml):**
```toml
[vars]
MICROSOFT_CLIENT_ID = "your-client-id"
MICROSOFT_TENANT_ID = "your-tenant-id" 
GRAPH_API_VERSION = "v1.0"
WORKER_DOMAIN = "your-worker.workers.dev"
PROTOCOL = "https"
```

**Secrets (Cloudflare Workers):**
```bash
# Deploy via Wrangler CLI
wrangler secret put MICROSOFT_CLIENT_SECRET
wrangler secret put ENCRYPTION_KEY
wrangler secret put COOKIE_SECRET
```

**Local Development (.dev.vars):**
```bash
# Local development only
MICROSOFT_CLIENT_ID=local-client-id
MICROSOFT_CLIENT_SECRET=local-secret
WORKER_DOMAIN=localhost:8787
PROTOCOL=http
```

**Environment Variable Mapping:**
```typescript
interface Env {
  // Public configuration
  MICROSOFT_CLIENT_ID: string;
  MICROSOFT_TENANT_ID: string;
  WORKER_DOMAIN: string;
  PROTOCOL: string;
  
  // Secrets
  MICROSOFT_CLIENT_SECRET: string;
  ENCRYPTION_KEY: string;
  COOKIE_SECRET: string;
  
  // Bindings
  MICROSOFT_MCP_AGENT: DurableObjectNamespace;
  OAUTH_KV: KVNamespace;
  CONFIG_KV: KVNamespace;
  CACHE_KV: KVNamespace;
}
```

### 10. Libraries & SDK Integration Analysis

**Dependency Strategy:**

**Why Each Library Was Chosen:**
- **@cloudflare/workers-oauth-provider**: Only OAuth 2.1 + PKCE implementation for Workers
- **agents**: Official Cloudflare MCP framework with Durable Objects support
- **@modelcontextprotocol/sdk**: Standard MCP types and protocol compliance
- **hono**: Lightweight, Workers-optimized web framework
- **zod**: Runtime type validation for API safety

**Integration Patterns:**

**OAuth Provider Extension:**
```typescript
const oauthProvider = new OAuthProvider({
  // Standard OAuth configuration
  authorizeEndpoint: '/authorize',
  tokenEndpoint: '/token',
  
  // Custom Microsoft integration
  tokenExchangeCallback: async (options) => {
    return await exchangeMicrosoftTokens(options);
  }
});
```

**MCP Agent Extension:**
```typescript
class MicrosoftMCPAgent extends McpAgent<Env, State, Props> {
  // Extend base agent with Microsoft-specific tools
  async init() {
    this.server.tool('sendEmail', 'Send email via Outlook', schema, handler);
    this.server.tool('getEmails', 'Get emails', schema, handler);
    // ... additional tools
  }
}
```

**What We Didn't Use (And Why):**
- **Express/Fastify**: Too heavy for Workers, chose Hono instead
- **Passport.js**: Doesn't support Workers runtime, used OAuth provider instead  
- **Axios**: Fetch API sufficient, avoided external HTTP client
- **Lodash**: Modern JS methods sufficient, avoided utility library bloat

### 11. MCP Protocol Implementation Patterns

**Tool Schema Definition:**
```typescript
// Zod schema validation for MCP tools
this.server.tool(
  'sendEmail',
  'Send an email via Outlook',
  {
    to: z.string().describe('Recipient email address'),
    subject: z.string().describe('Email subject'),
    body: z.string().describe('Email body content'),
    contentType: z.enum(['text', 'html']).default('html')
  },
  async (args): Promise<CallToolResult> => {
    // Tool implementation
  }
);
```

**JSON-RPC Message Handling:**
```typescript
// Protocol-compliant message routing
switch (message.method) {
  case 'initialize':
    return {
      jsonrpc: '2.0',
      id: message.id,
      result: {
        protocolVersion: '2024-11-05',
        serverInfo: { name: 'microsoft-365-mcp', version: '0.3.0' },
        capabilities: { tools: {}, resources: {}, prompts: {} }
      }
    };
    
  case 'tools/list':
    return { jsonrpc: '2.0', id: message.id, result: { tools: [...] } };
    
  case 'tools/call':
    return await this.handleToolCall(message.params);
}
```

**Resource Management:**
```typescript
// MCP resources for Microsoft 365 data
this.server.resource('profile', 'microsoft://profile', async () => {
  const profile = await this.graphClient.getUserProfile(accessToken);
  return {
    contents: [{
      uri: 'microsoft://profile',
      mimeType: 'application/json',
      text: JSON.stringify(profile, null, 2)
    }]
  };
});
```

### 12. Production Deployment Patterns

**Cloudflare Workers Optimization:**
```toml
# wrangler.toml optimizations
compatibility_date = "2024-11-05"
compatibility_flags = ["nodejs_compat"]

# Durable Objects configuration
[durable_objects]
bindings = [
  { name = "MICROSOFT_MCP_AGENT", class_name = "MicrosoftMCPAgent" }
]

# Migration strategy
[[migrations]]
tag = "v1"
new_classes = ["MicrosoftMCPAgent"]
```

**KV Namespace Provisioning:**
```bash
# Create KV namespaces
wrangler kv:namespace create "OAUTH_KV"
wrangler kv:namespace create "CONFIG_KV" 
wrangler kv:namespace create "CACHE_KV"

# Preview namespaces for development
wrangler kv:namespace create "OAUTH_KV" --preview
wrangler kv:namespace create "CONFIG_KV" --preview
wrangler kv:namespace create "CACHE_KV" --preview
```

**Deployment Strategy:**
```bash
# Type checking
npm run type-check

# Build verification  
npm run build

# Deploy to production
wrangler deploy

# Deploy secrets
wrangler secret put MICROSOFT_CLIENT_SECRET
wrangler secret put ENCRYPTION_KEY
wrangler secret put COOKIE_SECRET
```

**Custom Domain Routing:**
```toml
# Optional custom domain
routes = [
  { pattern = "mcp.your-domain.com/*", zone_name = "your-domain.com" }
]
```

**Edge Caching Strategy:**
```typescript
// Cache Microsoft Graph responses at the edge
const cacheKey = new Request(url, { method: 'GET' });
const cache = caches.default;

let response = await cache.match(cacheKey);
if (!response) {
  response = await fetchFromMicrosoftGraph(request);
  
  // Cache for 5 minutes
  response.headers.set('Cache-Control', 'public, max-age=300');
  await cache.put(cacheKey, response.clone());
}

return response;
```

---

*Complete deployment guide for production-ready Microsoft 365 MCP Server with enterprise security and performance optimization.*