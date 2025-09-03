# Microsoft 365 MCP Server - Operations Guide

Complete operational documentation for development, deployment, and maintenance of the Microsoft 365 MCP Server.

## Table of Contents

### Part I: Development Operations

- [Development Setup](#development-setup)
- [Project Structure](#project-structure)
- [Development Workflow](#development-workflow)
- [Testing Procedures](#testing-procedures)
- [Adding New Features](#adding-new-features)
- [Code Standards](#code-standards)
- [Debugging Guide](#debugging-guide)
- [Contributing Guidelines](#contributing-guidelines)

### Part II: Deployment Operations

- [Prerequisites](#prerequisites)
- [Production Deployment](#production-deployment)
- [Environment Configuration](#environment-configuration)
- [Security Setup](#security-setup)
- [Monitoring & Operations](#monitoring--operations)
- [Troubleshooting](#troubleshooting)
- [Maintenance Procedures](#maintenance-procedures)

### Part III: Advanced Operations

- [Production Monitoring](#production-monitoring)

---

## Part I: Development Operations

## Development Setup

### Prerequisites

- **Node.js** 18+ with npm
- **Wrangler CLI** for Cloudflare Workers development
- **Git** for version control
- **Microsoft 365 Developer Account** with app registration
- **Cloudflare Account** with Workers enabled

### Initial Setup

1. **Clone the repository:**

```bash
git clone https://github.com/nikolanovoselec/m365-mcp-server.git
cd m365-mcp-server
```

2. **Install dependencies:**

```bash
npm install
```

3. **Install Wrangler CLI globally:**

```bash
npm install -g wrangler
# or use npx: npx wrangler <command>
```

4. **Authenticate with Cloudflare:**

```bash
wrangler auth login
```

5. **Set up environment files:**

```bash
# Copy example files
cp .dev.vars.example .dev.vars
cp wrangler.example.toml wrangler.toml

# Configure with your Microsoft 365 and Cloudflare credentials
```

### Environment Configuration

The development environment requires specific variables for Microsoft 365 and Cloudflare integration. The .dev.vars file stores sensitive credentials locally and is never committed to version control. All secrets must be properly generated for security.

**.dev.vars file structure:**

```bash
# Microsoft 365 Configuration
MICROSOFT_CLIENT_ID=your-app-client-id
MICROSOFT_TENANT_ID=your-tenant-id
GRAPH_API_VERSION=v1.0

# Deployment Configuration
WORKER_DOMAIN=your-worker.your-subdomain.workers.dev
PROTOCOL=https

# Secrets (generate with: openssl rand -hex 32)
MICROSOFT_CLIENT_SECRET=your-client-secret
COOKIE_ENCRYPTION_KEY=your-32-char-encryption-key
ENCRYPTION_KEY=your-32-char-encryption-key
COOKIE_SECRET=your-cookie-secret
```

**wrangler.toml configuration:**

This configuration file defines the Worker's runtime settings, environment bindings, and resource mappings. It specifies the TypeScript entry point, Durable Object classes, and KV namespace bindings that connect the Worker to Cloudflare's distributed storage infrastructure.

```toml
name = "m365-mcp-server-dev"
main = "src/index.ts"
compatibility_date = "2025-01-01"
compatibility_flags = ["nodejs_compat"]

account_id = "your-cloudflare-account-id"

[vars]
GRAPH_API_VERSION = "v1.0"
MICROSOFT_CLIENT_ID = "your-client-id"
MICROSOFT_TENANT_ID = "your-tenant-id"
WORKER_DOMAIN = "your-worker.your-subdomain.workers.dev"
PROTOCOL = "https"

[[durable_objects.bindings]]
name = "MCP_OBJECT"
class_name = "MicrosoftMCPAgent"

[[kv_namespaces]]
binding = "OAUTH_KV"
id = "your-oauth-kv-namespace-id"

[[kv_namespaces]]
binding = "CONFIG_KV"
id = "your-config-kv-namespace-id"

[[kv_namespaces]]
binding = "CACHE_KV"
id = "your-cache-kv-namespace-id"
```

## Project Structure

The codebase follows a modular architecture with clear separation between protocol handling, OAuth flows, and Microsoft Graph integration. Each component has a specific responsibility in the MCP server implementation, enabling maintainable and testable code.

```
src/
├── index.ts                 # Main worker entry point & protocol routing
├── microsoft-mcp-agent.ts   # Durable Object MCP agent implementation
├── microsoft-handler.ts     # OAuth authentication handlers
├── microsoft-graph.ts       # Microsoft Graph API client
├── workers-oauth-utils.ts   # OAuth utility functions
└── utils.ts                 # General utility functions

tests/
├── unit/                    # Unit tests for individual components
├── integration/             # Integration tests with external APIs
└── fixtures/                # Test data and mock responses

docs/
├── README.md               # Project overview and quick start
├── ARCHITECTURE.md         # Technical architecture documentation
├── API_REFERENCE.md        # Complete API documentation
├── DEPLOYMENT.md           # Production deployment guide
└── DEVELOPMENT.md          # This development guide
```

### Key Components

**index.ts** - Main Entry Point:

- Protocol detection and routing
- OAuth provider configuration
- Microsoft token exchange callbacks
- Hybrid endpoint handling

**microsoft-mcp-agent.ts** - Durable Object Agent:

- MCP Server implementation
- Tool definitions and handlers
- Resource providers
- Authentication management

**microsoft-graph.ts** - Graph API Client:

- Microsoft Graph API integration
- Request/response handling
- Error processing and retries
- Response caching logic

**microsoft-handler.ts** - OAuth Handlers:

- Authorization flow handling
- Client approval dialogs
- Callback processing
- Token management

## Development Workflow

### 1. Start Development Server

```bash
# Start local development server with hot reload
npm run dev

# Alternative with specific binding
npx wrangler dev --local
```

The development server will:

- Start on `http://localhost:8787`
- Auto-reload on file changes
- Use local storage for KV namespaces
- Enable debug logging

### 2. Development Testing Flow

**Step 1: Test Tool Discovery**

```bash
curl -X POST http://localhost:8787/sse \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}'
```

**Step 2: Test Authentication Flow**

```bash
# 1. Trigger OAuth flow
curl -X GET "http://localhost:8787/authorize?client_id=test-client&response_type=code&redirect_uri=http://localhost:3000/callback&scope=User.Read"

# 2. Complete flow in browser
# 3. Extract tokens from callback
```

**Step 3: Test Tool Execution**

```bash
curl -X POST http://localhost:8787/sse \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer YOUR_DEV_TOKEN" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"getEmails","arguments":{"count":5}}}'
```

### 3. Hot Reloading

The development server supports hot reloading for rapid iteration during development. Changes to TypeScript files trigger automatic rebuilds while preserving WebSocket connections and KV storage state, eliminating the need to re-authenticate or recreate test data.

```typescript
// Make changes to any .ts file
// Server automatically restarts
// WebSocket connections are preserved
// KV data persists across restarts
```

### 4. Local Storage

Development uses local storage for KV namespaces to simulate Cloudflare's distributed storage locally. The .wrangler/state directory contains SQLite databases that persist data between development sessions, allowing you to test OAuth flows and data caching without affecting production.

```bash
# View local KV data
ls .wrangler/state/v3/kv/

# Clear local storage
rm -rf .wrangler/state/
```

### 4. Code Quality Workflow (NEW - Improved)

This project now includes automated code quality checks to prevent issues:

**Setup Pre-commit Hooks (One-time):**

```bash
# Install git hooks that run validation before each commit
npm run setup-hooks
```

**Development Commands:**

```bash
# Full validation (runs automatically before build/deploy)
npm run validate

# Individual checks
npm run type-check    # TypeScript compilation check
npm run lint          # ESLint code quality check
npm run lint:fix      # Auto-fix ESLint issues
npm run format        # Format code with Prettier
npm run format:check  # Check code formatting

# Clean build environment
npm run clean         # Remove build artifacts and temp files

# Safe build and deploy
npm run build         # Validates, cleans, then builds
npm run deploy        # Validates, then deploys to Cloudflare
```

**Manual Pre-commit Check:**

```bash
# Run the same checks as the git hook
npm run precommit
```

**Code Quality Benefits:**

- Prevents backup files from being committed
- Ensures TypeScript compilation passes
- Enforces consistent code style
- Catches unused variables and imports
- Validates build before deployment
- Automatic cleanup of temporary files

## Testing Procedures

### Unit Tests

Unit tests validate individual components in isolation using mocked dependencies. They ensure that each function, class, and module behaves correctly under various input conditions without requiring external services or network calls.

```bash
# Run all unit tests
npm test

# Run specific test file
npm test -- tests/unit/microsoft-graph.test.ts

# Run tests in watch mode
npm test -- --watch
```

**Test Structure Example:**

This example demonstrates the recommended testing pattern using Vitest framework with proper mocking and assertion strategies. Tests are organized by component, with clear setup, execution, and verification phases for maintainability.

```typescript
// tests/unit/microsoft-graph.test.ts
import { describe, it, expect, beforeEach } from "vitest";
import { MicrosoftGraphClient } from "../../src/microsoft-graph";

describe("MicrosoftGraphClient", () => {
  let client: MicrosoftGraphClient;
  let mockEnv: Env;

  beforeEach(() => {
    mockEnv = {
      GRAPH_API_VERSION: "v1.0",
      MICROSOFT_CLIENT_ID: "test-id",
      // ... other env vars
    };
    client = new MicrosoftGraphClient(mockEnv);
  });

  it("should send email successfully", async () => {
    const mockResponse = { status: 204, json: () => Promise.resolve({}) };
    global.fetch = vi.fn().mockResolvedValue(mockResponse);

    const result = await client.sendEmail("fake-token", {
      to: "test@example.com",
      subject: "Test",
      body: "Hello",
    });

    expect(fetch).toHaveBeenCalledWith(
      "https://graph.microsoft.com/v1.0/me/sendMail",
      expect.objectContaining({
        method: "POST",
        headers: expect.objectContaining({
          Authorization: "Bearer fake-token",
        }),
      }),
    );
  });

  it("should handle Microsoft Graph errors", async () => {
    const mockError = {
      status: 400,
      json: () =>
        Promise.resolve({
          error: { message: "Invalid recipient" },
        }),
    };
    global.fetch = vi.fn().mockResolvedValue(mockError);

    await expect(
      client.sendEmail("fake-token", {
        to: "invalid-email",
        subject: "Test",
        body: "Hello",
      }),
    ).rejects.toThrow("Microsoft Graph API error: 400 - Invalid recipient");
  });
});
```

### Integration Tests

Integration tests verify that components work correctly together with real or simulated external services. These tests validate OAuth flows, Microsoft Graph API interactions, and end-to-end scenarios to ensure the system functions properly in production-like conditions.

```bash
# Run integration tests (requires valid credentials)
npm run test:integration

# Run against local development server
npm run test:integration:local
```

**Integration Test Example:**

This integration test validates the complete OAuth authorization flow from client registration through token exchange. It simulates real client behavior and verifies that all components interact correctly to provide authenticated access.

```typescript
// tests/integration/oauth-flow.test.ts
import { describe, it, expect } from "vitest";

describe("OAuth Integration", () => {
  const BASE_URL = process.env.TEST_BASE_URL || "http://localhost:8787";

  it("should complete full OAuth flow", async () => {
    // 1. Register client
    const registerResponse = await fetch(`${BASE_URL}/register`, {
      method: "POST",
      body: JSON.stringify({
        client_name: "Integration Test Client",
        redirect_uris: ["http://localhost:3000/callback"],
      }),
    });

    const client = await registerResponse.json();
    expect(client.client_id).toBeDefined();

    // 2. Start authorization
    const authUrl = `${BASE_URL}/authorize?client_id=${client.client_id}&response_type=code&redirect_uri=http://localhost:3000/callback`;
    const authResponse = await fetch(authUrl);
    expect(authResponse.status).toBe(302);

    // 3. Complete with Microsoft (mock)
    // ... additional OAuth flow testing
  });
});
```

### Load Testing

Load testing simulates production traffic patterns to identify performance bottlenecks and capacity limits. Using k6 or Artillery, these tests gradually increase concurrent users to measure response times, error rates, and resource utilization under stress.

```bash
# Install k6 for load testing
npm install -g k6

# Run load tests
k6 run tests/load/basic-load.js
```

**Load Test Example:**

This k6 script simulates realistic user behavior with gradual traffic ramp-up and sustained load periods. It measures key performance indicators like response time and success rate to ensure the system meets performance requirements.

```javascript
// tests/load/basic-load.js
import http from "k6/http";
import { check, sleep } from "k6";

export let options = {
  stages: [
    { duration: "2m", target: 10 }, // Ramp up
    { duration: "5m", target: 50 }, // Stay at 50 users
    { duration: "2m", target: 0 }, // Ramp down
  ],
};

export default function () {
  // Test tool discovery
  let response = http.post(
    `${__ENV.BASE_URL}/sse`,
    JSON.stringify({
      jsonrpc: "2.0",
      id: 1,
      method: "tools/list",
      params: {},
    }),
    {
      headers: { "Content-Type": "application/json" },
    },
  );

  check(response, {
    "tool discovery succeeds": (r) => r.status === 200,
    "response time < 500ms": (r) => r.timings.duration < 500,
  });

  sleep(1);
}
```

## Adding New Features

### Adding a New Microsoft Graph Tool

**Step 1: Add to Microsoft Graph Client**

```typescript
// src/microsoft-graph.ts
export interface NewToolParams {
  param1: string;
  param2?: number;
}

export class MicrosoftGraphClient {
  async newTool(accessToken: string, params: NewToolParams): Promise<any> {
    const url = `${this.baseUrl}/me/newEndpoint`;

    const body = {
      property1: params.param1,
      property2: params.param2 || 0,
    };

    const response = await this.makeGraphRequest(
      accessToken,
      url,
      "POST",
      body,
    );
    return response;
  }
}
```

**Step 2: Add Tool Definition**

Define the MCP tool interface with proper type safety, validation, and error handling for the new Microsoft Graph functionality.

```typescript
// src/microsoft-mcp-agent.ts
async init() {
  // ... existing tools

  this.server.tool(
    'newTool',
    'Description of the new tool',
    {
      param1: z.string().describe('Description of param1'),
      param2: z.number().optional().describe('Description of param2'),
    },
    async (args): Promise<CallToolResult> => {
      const accessToken = this.props?.microsoftAccessToken;
      if (!accessToken) {
        return this.getAuthErrorResponse();
      }

      try {
        const result = await this.graphClient.newTool(accessToken, args);
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      } catch (error: any) {
        return {
          content: [
            { type: 'text', text: `Failed to execute newTool: ${error?.message || String(error)}` },
          ],
          isError: true,
        };
      }
    }
  );
}
```

**Step 3: Add Tests**

Create comprehensive unit tests for the new tool including success cases, error scenarios, and edge conditions.

```typescript
// tests/unit/new-tool.test.ts
describe("newTool", () => {
  it("should execute newTool successfully", async () => {
    const mockResponse = { id: "new-item-id", status: "created" };
    global.fetch = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      json: () => Promise.resolve(mockResponse),
    });

    const result = await client.newTool("fake-token", {
      param1: "test-value",
      param2: 42,
    });

    expect(result).toEqual(mockResponse);
    expect(fetch).toHaveBeenCalledWith(
      "https://graph.microsoft.com/v1.0/me/newEndpoint",
      expect.objectContaining({
        method: "POST",
        body: JSON.stringify({
          property1: "test-value",
          property2: 42,
        }),
      }),
    );
  });
});
```

**Step 4: Update Documentation**

Document the new tool in API reference and architecture documents to maintain comprehensive documentation.

```typescript
// API_REFERENCE.md - Add tool documentation
// ARCHITECTURE.md - Update tool mapping table if needed
```

### Adding a New Resource

**Step 1: Add Resource Provider**

```typescript
// src/microsoft-mcp-agent.ts
async init() {
  // ... existing resources

  this.server.resource('newResource', 'microsoft://newResource', async () => {
    const accessToken = this.props?.microsoftAccessToken;
    if (!accessToken) {
      return {
        contents: [
          {
            uri: 'microsoft://newResource',
            mimeType: 'application/json',
            text: JSON.stringify({ error: 'Authentication required', authenticated: false }, null, 2),
          },
        ],
      };
    }

    try {
      const data = await this.graphClient.getNewResourceData(accessToken);
      return {
        contents: [
          {
            uri: 'microsoft://newResource',
            mimeType: 'application/json',
            text: JSON.stringify(data, null, 2),
          },
        ],
      };
    } catch (error: any) {
      return {
        contents: [
          {
            uri: 'microsoft://newResource',
            mimeType: 'application/json',
            text: JSON.stringify({ error: error.message || 'Failed to fetch data', authenticated: true }, null, 2),
          },
        ],
      };
    }
  });
}
```

### Adding Protocol Support

**Step 1: Extend Protocol Detection**

```typescript
// src/index.ts
async function handleHybridMcp(
  request: Request,
  env: Env,
  ctx: ExecutionContext,
): Promise<Response> {
  // ... existing protocol detection

  // Add new protocol detection
  if (request.headers.get("X-Custom-Protocol") === "new-protocol") {
    return handleNewProtocol(request, env);
  }

  // ... rest of handler
}

async function handleNewProtocol(
  request: Request,
  env: Env,
): Promise<Response> {
  // Implement new protocol handling
  return new Response("New protocol handler", { status: 200 });
}
```

## Code Standards

### TypeScript Configuration

The TypeScript compiler configuration enforces strict type checking and modern JavaScript features. This configuration ensures code quality, catches type errors at compile time, and maintains compatibility with Cloudflare Workers runtime environment.

```json
// tsconfig.json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "ES2022",
    "moduleResolution": "node",
    "allowSyntheticDefaultImports": true,
    "esModuleInterop": true,
    "strict": true,
    "skipLibCheck": true,
    "forceConsistentCasingInFileNames": true,
    "types": ["@cloudflare/workers-types"]
  }
}
```

### Code Style Guidelines

**1. Naming Conventions:**

```typescript
// Use camelCase for functions and variables
const accessToken = "token";
async function sendEmail() {}

// Use PascalCase for classes and types
class MicrosoftGraphClient {}
interface EmailParams {}

// Use UPPER_SNAKE_CASE for constants
const MAX_EMAIL_COUNT = 50;
```

**2. Function Structure:**

Standard function organization pattern with JSDoc comments, input validation, error handling, and meaningful error messages.

```typescript
// Always include JSDoc comments for public functions
/**
 * Sends an email via Microsoft Graph API
 * @param accessToken - Microsoft Graph access token
 * @param params - Email parameters including recipient, subject, body
 * @returns Promise resolving to the API response
 * @throws {Error} When API request fails or token is invalid
 */
async function sendEmail(
  accessToken: string,
  params: EmailParams,
): Promise<any> {
  // Validate inputs first
  if (!accessToken) {
    throw new Error("Access token is required");
  }

  if (!params.to || !params.subject || !params.body) {
    throw new Error("Missing required email parameters");
  }

  try {
    // Main logic
    const response = await this.makeGraphRequest(
      accessToken,
      url,
      "POST",
      body,
    );
    return response;
  } catch (error) {
    // Always provide meaningful error context
    console.error("Failed to send email:", error);
    throw new Error(
      `Email sending failed: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
}
```

**3. Error Handling Pattern:**

Consistent error handling approach that provides useful debugging information while avoiding sensitive data exposure.

```typescript
// Use consistent error handling pattern
try {
  const result = await apiCall();
  return { success: true, data: result };
} catch (error) {
  console.error('Operation failed:', error);

  if (error instanceof Microsoft GraphError) {
    // Handle specific API errors
    return { success: false, error: error.message, code: error.code };
  }

  // Handle unknown errors
  return { success: false, error: 'Unknown error occurred' };
}
```

**4. Type Safety:**

TypeScript patterns for ensuring type safety at compile time and runtime with proper type guards and interfaces.

```typescript
// Always define proper types
interface ToolResult {
  success: boolean;
  data?: any;
  error?: string;
  code?: number;
}

// Use type guards for runtime validation
function isEmailParams(obj: any): obj is EmailParams {
  return (
    typeof obj === "object" &&
    typeof obj.to === "string" &&
    typeof obj.subject === "string" &&
    typeof obj.body === "string"
  );
}
```

### Linting Configuration

ESLint configuration enforces consistent code style and catches common programming errors. The rules are designed to prevent bugs, improve code readability, and maintain consistency across the codebase while allowing necessary flexibility for Cloudflare Workers patterns.

```json
// .eslintrc.json
{
  "extends": ["@typescript-eslint/recommended", "prettier"],
  "parser": "@typescript-eslint/parser",
  "plugins": ["@typescript-eslint"],
  "rules": {
    "@typescript-eslint/no-unused-vars": "error",
    "@typescript-eslint/explicit-function-return-type": "warn",
    "@typescript-eslint/no-explicit-any": "warn",
    "prefer-const": "error",
    "no-console": ["warn", { "allow": ["error", "warn"] }]
  }
}
```

**Run linting:**

```bash
npm run lint        # Check for issues
npm run lint:fix    # Auto-fix issues
```

## Debugging Guide

### Local Development Debugging

**1. Enable Debug Logging:**

Debug logging provides detailed information about request processing, OAuth flows, and API interactions. Enable this mode during development or when troubleshooting production issues to see comprehensive execution traces.

```typescript
// Set environment variable
DEBUG=true npm run dev

// Or add to .dev.vars
DEBUG=true
```

**2. Console Debugging:**

Structured logging approach for development debugging that provides clear visibility into request flow and data transformations.

```typescript
// Use structured logging
console.log("=== PROTOCOL DETECTION ===");
console.log("Method:", request.method);
console.log("Headers:", Object.fromEntries(request.headers.entries()));
console.log("URL:", request.url);

// Debug tool execution
console.log("=== TOOL EXECUTION ===");
console.log("Tool name:", toolName);
console.log("Arguments:", JSON.stringify(args, null, 2));
console.log("Access token present:", !!accessToken);
```

**3. KV Storage Debugging:**

KV storage debugging helps diagnose issues with OAuth tokens, client configurations, and cached data. These utilities allow inspection of stored values and verification of encryption/decryption processes during development.

```typescript
// View KV contents during development
async function debugKV(env: Env, namespace: string, key: string) {
  const value = await env[namespace].get(key);
  console.log(`KV[${namespace}][${key}]:`, value);
  return value;
}

// Usage in handlers
await debugKV(env, "OAUTH_KV", `client:${clientId}`);
```

### Production Debugging

**1. Cloudflare Logs:**

```bash
# Stream real-time logs
wrangler tail --name m365-mcp-server

# Filter specific logs
wrangler tail --name m365-mcp-server --grep "ERROR"

# Pretty format logs
wrangler tail --name m365-mcp-server --format pretty
```

**2. Error Tracking:**

```typescript
// Structured error logging for production
function logError(error: any, context: string, metadata?: any) {
  console.error(
    JSON.stringify({
      timestamp: new Date().toISOString(),
      context,
      error: error instanceof Error ? error.message : String(error),
      stack: error instanceof Error ? error.stack : undefined,
      metadata,
    }),
  );
}

// Usage
try {
  await sendEmail(token, params);
} catch (error) {
  logError(error, "sendEmail", { userId, emailParams: params });
  throw error;
}
```

**3. Performance Monitoring:**

```typescript
// Add timing measurements
async function timedOperation<T>(
  name: string,
  operation: () => Promise<T>,
): Promise<T> {
  const start = Date.now();
  try {
    const result = await operation();
    console.log(`${name} completed in ${Date.now() - start}ms`);
    return result;
  } catch (error) {
    console.log(`${name} failed after ${Date.now() - start}ms:`, error);
    throw error;
  }
}

// Usage
const emails = await timedOperation("getEmails", () =>
  graphClient.getEmails(token, { count: 10 }),
);
```

### Common Issues and Solutions

**Issue: WebSocket connections failing**

```typescript
// Debug WebSocket upgrade detection
console.log("=== WEBSOCKET DEBUG ===");
console.log("Upgrade header:", request.headers.get("Upgrade"));
console.log("Connection header:", request.headers.get("Connection"));
console.log("WebSocket-Key:", request.headers.get("Sec-WebSocket-Key"));
console.log("WebSocket-Version:", request.headers.get("Sec-WebSocket-Version"));

// Solution: Check all WebSocket signals
const isWebSocketRequest =
  request.headers.get("Upgrade")?.toLowerCase() === "websocket" ||
  (request.headers.get("Sec-WebSocket-Key") &&
    request.headers.get("Sec-WebSocket-Version"));
```

**Issue: Microsoft Graph 204 responses causing JSON parsing errors**

```typescript
// Debug response parsing
async function debugGraphResponse(response: Response) {
  console.log("Response status:", response.status);
  console.log(
    "Response headers:",
    Object.fromEntries(response.headers.entries()),
  );

  const contentType = response.headers.get("content-type");
  const contentLength = response.headers.get("content-length");

  console.log("Content-Type:", contentType);
  console.log("Content-Length:", contentLength);

  if (response.status === 204 || contentLength === "0") {
    console.log("Empty response - returning {}");
    return {};
  }

  const text = await response.text();
  console.log("Response body:", text);

  return text ? JSON.parse(text) : {};
}
```

**Issue: Authentication token expiry**

```typescript
// Debug token lifecycle
function debugToken(token: any, context: string) {
  if (!token) {
    console.log(`${context}: No token present`);
    return;
  }

  console.log(`${context}: Token present`);

  if (token.exp) {
    const expiry = new Date(token.exp * 1000);
    const now = new Date();
    const remaining = expiry.getTime() - now.getTime();

    console.log(`Token expires: ${expiry.toISOString()}`);
    console.log(`Time remaining: ${Math.round(remaining / 1000)}s`);
    console.log(`Token expired: ${remaining <= 0}`);
  }
}
```

## Contributing Guidelines

### Contribution Process

**1. Fork and Clone**

```bash
git clone https://github.com/your-username/m365-mcp-server.git
cd m365-mcp-server
git remote add upstream https://github.com/nikolanovoselec/m365-mcp-server.git
```

**2. Create Feature Branch**

```bash
git checkout -b feature/your-feature-name
```

**3. Development Process**

```bash
# Make your changes
# Add tests for new functionality
npm test

# Ensure code quality
npm run lint
npm run type-check

# Test in development environment
npm run dev
```

**4. Commit Guidelines**

```bash
# Use conventional commit format
git commit -m "feat: add support for SharePoint documents"
git commit -m "fix: handle Microsoft Graph rate limiting"
git commit -m "docs: update API reference for new tools"
git commit -m "test: add integration tests for calendar tools"
```

**5. Submit Pull Request**

```bash
git push origin feature/your-feature-name
# Create pull request via GitHub interface
```

### Pull Request Requirements

**Required Checks:**

- All tests passing
- Type checking passes
- Linting passes
- No security vulnerabilities
- Documentation updated
- Breaking changes documented

**PR Template:**

```markdown
## Description

Brief description of changes and motivation.

## Type of Change

- [ ] Bug fix (non-breaking change which fixes an issue)
- [ ] New feature (non-breaking change which adds functionality)
- [ ] Breaking change (fix or feature that would cause existing functionality to not work as expected)
- [ ] Documentation update

## Testing

- [ ] Unit tests added/updated
- [ ] Integration tests added/updated
- [ ] Manual testing completed
- [ ] Load testing completed (if applicable)

## Documentation

- [ ] Code comments updated
- [ ] API documentation updated
- [ ] Architecture documentation updated
- [ ] Development guide updated

## Checklist

- [ ] My code follows the project's style guidelines
- [ ] I have performed a self-review of my code
- [ ] I have made corresponding changes to the documentation
- [ ] My changes generate no new warnings
- [ ] I have added tests that prove my fix is effective or that my feature works
- [ ] New and existing unit tests pass locally with my changes
```

### Code Review Process

**Review Criteria:**

1. **Functionality**: Does the code do what it's supposed to do?
2. **Code Quality**: Is the code clean, readable, and maintainable?
3. **Performance**: Are there any performance implications?
4. **Security**: Are there any security vulnerabilities?
5. **Testing**: Is the code adequately tested?
6. **Documentation**: Is the code properly documented?

**Review Timeline:**

- Initial review within 2 business days
- Follow-up reviews within 1 business day
- Final approval within 3 business days

---

_This development guide provides comprehensive information for contributing to and maintaining the Microsoft 365 MCP Server project. Follow these guidelines to ensure high-quality, maintainable code that integrates seamlessly with the existing architecture._

---

## Part II: Deployment Operations

- [Monitoring & Operations](#monitoring--operations)
- [Troubleshooting](#troubleshooting)
- [Maintenance Procedures](#maintenance-procedures)

## Prerequisites

### Required Accounts & Services

**Microsoft 365 Setup:**

- Microsoft 365 Business/Enterprise account with admin access
- Microsoft Entra ID (Azure AD) admin privileges
- Application registration permissions

**Cloudflare Setup:**

- Cloudflare account (Free tier sufficient for development)
- Workers & KV storage enabled
- Custom domain (optional but recommended for production)

**Development Tools:**

- Node.js 18+ with npm
- Git for source control

### Install Wrangler CLI

```bash
# Install globally via npm
npm install -g wrangler

# Verify installation
wrangler --version

# Alternative: Use npx (no global install)
npx wrangler --version
```

## Production Deployment

### Step 1: Cloudflare Setup

**Authentication:**

```bash
# Opens browser for authentication
wrangler auth login

# Verify authentication
wrangler whoami
# Should show your account details
```

**Create KV Namespaces:**

```bash
# Create required KV namespaces
wrangler kv:namespace create "OAUTH_KV" --env production
wrangler kv:namespace create "CONFIG_KV" --env production
wrangler kv:namespace create "CACHE_KV" --env production

# Note the namespace IDs for wrangler.toml configuration
```

### Step 2: Project Setup

**Clone Repository:**

```bash
git clone https://github.com/nikolanovoselec/m365-mcp-server.git
cd m365-mcp-server
npm install
```

**Environment Configuration:**

```bash
# Copy templates
cp .dev.vars.example .dev.vars
cp wrangler.example.toml wrangler.toml

# Configure with your production values
```

### Step 3: Microsoft Entra ID Configuration

**Create Application Registration:**

1. Navigate to Azure Portal → Microsoft Entra ID → App registrations
2. Click "New registration"
3. Configure application:
   - **Name**: "Microsoft 365 MCP Server"
   - **Account types**: "Accounts in this organizational directory only"
   - **Redirect URI**: `https://your-worker-domain.com/callback`

**Configure Redirect URIs:**
Add these redirect URIs to your app registration:

```
https://your-worker-domain.workers.dev/callback
https://your-custom-domain.com/callback (if using custom domain)
```

**Configure API Permissions:**
Add these Microsoft Graph permissions:

- `User.Read` (Delegated)
- `Mail.Read` (Delegated)
- `Mail.ReadWrite` (Delegated)
- `Mail.Send` (Delegated)
- `Calendars.Read` (Delegated)
- `Calendars.ReadWrite` (Delegated)
- `Contacts.ReadWrite` (Delegated)
- `OnlineMeetings.ReadWrite` (Delegated)
- `ChannelMessage.Send` (Delegated)
- `Team.ReadBasic.All` (Delegated)

**Grant Admin Consent** for all configured permissions.

**Create Client Secret:**

1. Go to "Certificates & secrets"
2. Click "New client secret"
3. Set expiration (24 months recommended)
4. Copy the secret value immediately (you won't see it again)

### Step 4: Deploy Secrets

```bash
# Set Microsoft client secret
wrangler secret put MICROSOFT_CLIENT_SECRET

# Set encryption keys (generate with: openssl rand -hex 32)
wrangler secret put ENCRYPTION_KEY
wrangler secret put COOKIE_ENCRYPTION_KEY
wrangler secret put COOKIE_SECRET
```

### Step 5: Configure wrangler.toml

Update your `wrangler.toml` with production values:

```toml
name = "m365-mcp-server-prod"
main = "src/index.ts"
compatibility_date = "2025-01-01"
compatibility_flags = ["nodejs_compat"]

# Your Cloudflare account ID
account_id = "your-account-id-here"

[vars]
GRAPH_API_VERSION = "v1.0"
MICROSOFT_CLIENT_ID = "your-microsoft-client-id"
MICROSOFT_TENANT_ID = "your-microsoft-tenant-id"
WORKER_DOMAIN = "your-worker.your-subdomain.workers.dev"
PROTOCOL = "https"

# Durable Objects
[[durable_objects.bindings]]
name = "MCP_OBJECT"
class_name = "MicrosoftMCPAgent"

# KV Namespaces (use production namespace IDs)
[[kv_namespaces]]
binding = "OAUTH_KV"
id = "your-oauth-kv-namespace-id"

[[kv_namespaces]]
binding = "CONFIG_KV"
id = "your-config-kv-namespace-id"

[[kv_namespaces]]
binding = "CACHE_KV"
id = "your-cache-kv-namespace-id"
```

### Step 6: Deploy to Production

```bash
# Validate configuration
npm run type-check

# Deploy to Cloudflare Workers
wrangler deploy

# Verify deployment
curl https://your-worker-domain.workers.dev/sse \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1.0.0"}}}'
```

## Environment Configuration

### Production Environment Variables

Production configuration separates public variables from secrets for security. Public variables are defined in wrangler.toml and deployed with the code, while secrets are encrypted and stored separately in Cloudflare's secure storage.

**Public Variables (in wrangler.toml [vars]):**

```toml
GRAPH_API_VERSION = "v1.0"                    # Microsoft Graph API version
MICROSOFT_CLIENT_ID = "your-client-id"        # Microsoft app client ID
MICROSOFT_TENANT_ID = "your-tenant-id"        # Microsoft tenant ID
WORKER_DOMAIN = "your-domain.workers.dev"     # Your worker domain
PROTOCOL = "https"                             # Protocol for callbacks
```

**Secrets (set via wrangler secret put):**

These sensitive values must be set using Cloudflare's secret management to ensure they're encrypted at rest and never exposed in logs or source code. Each secret serves a critical security function in the OAuth flow and data protection.

```bash
MICROSOFT_CLIENT_SECRET   # Microsoft app client secret
ENCRYPTION_KEY           # 32-character encryption key for tokens
COOKIE_ENCRYPTION_KEY    # 32-character key for cookie encryption
COOKIE_SECRET           # Secret for HMAC cookie signing
```

### Custom Domain Setup (Optional)

Custom domains provide professional branding and simplified URLs for production deployments. This configuration maps your domain to the Cloudflare Worker and updates all OAuth callbacks to use the custom domain instead of the workers.dev subdomain.

**Configure Custom Domain:**

```bash
# Add custom domain to Cloudflare Workers
wrangler custom-domains add your-custom-domain.com

# Update wrangler.toml
WORKER_DOMAIN = "your-custom-domain.com"
```

**Update Microsoft App Registration:**

- Add `https://your-custom-domain.com/callback` to redirect URIs
- Update application homepage URL

### Development vs Production

Environment-specific configurations ensure proper isolation between development and production systems. Development uses local URLs and test credentials, while production uses secure HTTPS endpoints and real Microsoft 365 tenant configurations.

**Development (.dev.vars):**

```bash
MICROSOFT_CLIENT_ID=dev-client-id
MICROSOFT_TENANT_ID=dev-tenant-id
WORKER_DOMAIN=localhost:8787
PROTOCOL=http
```

**Production (wrangler.toml):**

```toml
MICROSOFT_CLIENT_ID = "prod-client-id"
MICROSOFT_TENANT_ID = "prod-tenant-id"
WORKER_DOMAIN = "your-production-domain.com"
PROTOCOL = "https"
```

## Security Setup

### OAuth Security Configuration

The OAuth implementation follows OAuth 2.1 best practices with PKCE for all public clients and encrypted token storage. These security measures prevent token interception, replay attacks, and unauthorized access even if individual components are compromised.

**PKCE Challenge Method:**

- Uses S256 method for code challenge
- Generates cryptographically secure code verifiers
- All public clients use PKCE (no client secrets)

**Token Security:**

All tokens are encrypted using AES-256-GCM with unique initialization vectors before storage. The encryption keys are stored as Cloudflare secrets, separate from the encrypted data, providing defense-in-depth against token theft.

```bash
# Generate secure encryption keys
openssl rand -hex 32  # For ENCRYPTION_KEY
openssl rand -hex 32  # For COOKIE_ENCRYPTION_KEY
openssl rand -hex 32  # For COOKIE_SECRET

# Set as secrets (never in environment variables)
wrangler secret put ENCRYPTION_KEY
wrangler secret put COOKIE_ENCRYPTION_KEY
wrangler secret put COOKIE_SECRET
```

**Client Registration Security:**

- Dynamic client registration enabled for web applications
- Static client ID aliasing for mcp-remote compatibility
- HMAC-signed approval cookies with 1-year expiration

### Network Security

Network-layer security ensures all communications are encrypted in transit and properly authenticated. These configurations prevent man-in-the-middle attacks, session hijacking, and cross-origin attacks.

**HTTPS Enforcement:**

- All production endpoints use HTTPS
- HTTP requests automatically upgraded
- Secure cookie flags enabled

**CORS Configuration:**

Cross-Origin Resource Sharing headers control which domains can access the API from web browsers. This configuration allows web-based MCP connectors to securely communicate with the server while preventing unauthorized cross-origin requests.

```typescript
// Configured in production for web connectors
headers: {
  'Access-Control-Allow-Origin': 'https://trusted-domain.com',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization'
}
```

## Monitoring & Operations

### Cloudflare Analytics

Cloudflare provides comprehensive monitoring and analytics for deployed Workers. Real-time logs and metrics help identify issues, track usage patterns, and optimize performance in production environments.

**View Real-time Logs:**

```bash
# Stream live logs
wrangler tail --name m365-mcp-server-prod

# Filter error logs
wrangler tail --name m365-mcp-server-prod --grep "ERROR"

# Pretty format logs
wrangler tail --name m365-mcp-server-prod --format pretty
```

**Analytics Dashboard:**

- Navigate to Cloudflare Dashboard → Workers → your-worker → Analytics
- Monitor requests, errors, CPU usage, and response times

### Key Metrics to Monitor

Monitoring these key performance indicators ensures the system operates within acceptable parameters. Regular tracking helps identify degradation before it impacts users and provides data for capacity planning.

**Performance Metrics:**

- **Response Time**: Should be < 500ms for most requests
- **Error Rate**: Should be < 1% under normal conditions
- **CPU Usage**: Should be < 10ms per request
- **Memory Usage**: Typically < 128MB per request

**Business Metrics:**

- **Authentication Success Rate**: > 95%
- **Tool Execution Success Rate**: > 98%
- **Microsoft Graph API Errors**: < 2%

### Health Checks

Automated health checks verify system availability and functionality at regular intervals. These checks detect issues before users report them and can trigger alerts for immediate intervention when problems occur.

**Automated Health Check:**

```bash
#!/bin/bash
# health-check.sh
WORKER_URL="https://your-domain.workers.dev"

# Test tool discovery
response=$(curl -s -X POST "$WORKER_URL/sse" \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}')

if echo "$response" | jq -e '.result.tools | length > 0' > /dev/null; then
  echo "Health check passed"
  exit 0
else
  echo "Health check failed"
  echo "$response"
  exit 1
fi
```

**Set up monitoring:**

```bash
# Run every 5 minutes
*/5 * * * * /path/to/health-check.sh
```

### Performance Benchmarking

Regular performance benchmarking establishes baseline metrics and identifies performance regressions. These tests simulate production traffic patterns to validate that the system meets performance requirements under expected load.

**Load Testing with k6:**

```javascript
// load-test.js
import http from "k6/http";
import { check } from "k6";

export let options = {
  stages: [
    { duration: "2m", target: 10 },
    { duration: "5m", target: 50 },
    { duration: "2m", target: 0 },
  ],
};

export default function () {
  let response = http.post(
    `${__ENV.WORKER_URL}/sse`,
    JSON.stringify({
      jsonrpc: "2.0",
      id: 1,
      method: "tools/list",
      params: {},
    }),
    { headers: { "Content-Type": "application/json" } },
  );

  check(response, {
    "status is 200": (r) => r.status === 200,
    "response time < 500ms": (r) => r.timings.duration < 500,
  });
}
```

**Run load test:**

```bash
k6 run -e WORKER_URL=https://your-domain.workers.dev load-test.js
```

## Troubleshooting

Comprehensive guide for diagnosing and resolving issues across all layers of the MCP server stack. These solutions address the most frequently encountered problems in development and production environments.

### Common Deployment Issues

Typical problems encountered when deploying to Cloudflare Workers and their proven solutions. Most deployment issues stem from configuration mismatches or missing dependencies.

**Issue: "Module not found" during deployment**

```bash
# Solution: Ensure all dependencies are installed
npm install
npm run type-check
wrangler deploy
```

**Issue: KV namespace binding errors**

```bash
# Solution: Verify KV namespace IDs in wrangler.toml
wrangler kv:namespace list
# Update wrangler.toml with correct IDs
```

**Issue: Microsoft Graph API 401 Unauthorized**

```bash
# Check token expiry and refresh
curl -X POST https://your-domain.workers.dev/sse \
  -H "Authorization: Bearer $TOKEN" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"authenticate","arguments":{}}}'

# Solution: Ensure proper scope configuration and admin consent
```

### Authentication Issues

Diagnostic procedures for OAuth and Microsoft authentication problems. Authentication issues are the most common source of failures and require systematic debugging across multiple components.

**OAuth Flow Debugging:**

```bash
# Test client registration
curl -X POST https://your-domain.workers.dev/register \
  -H "Content-Type: application/json" \
  -d '{
    "client_name": "Debug Client",
    "redirect_uris": ["https://debug.example.com/callback"]
  }'

# Test authorization endpoint
curl -v "https://your-domain.workers.dev/authorize?client_id=CLIENT_ID&response_type=code&redirect_uri=REDIRECT_URI&scope=User.Read"
```

**Token Exchange Issues:**

- Verify Microsoft client secret is correctly set
- Check redirect URI matches exactly (including trailing slashes)
- Ensure all required scopes are configured and granted

### Runtime Errors

Resolution strategies for errors that occur during normal operation including protocol failures, API errors, and connection problems. These issues often require real-time log analysis.

**WebSocket Connection Issues:**

```bash
# Test WebSocket upgrade headers
curl -X GET https://your-domain.workers.dev/sse \
  -H "Upgrade: websocket" \
  -H "Connection: Upgrade" \
  -H "Sec-WebSocket-Key: test" \
  -H "Sec-WebSocket-Version: 13" \
  -v

# Should return 101 Switching Protocols
```

**Microsoft Graph API Errors:**

- **403 Forbidden**: Check API permissions and admin consent
- **429 Too Many Requests**: Implement exponential backoff
- **500 Internal Server Error**: Check Microsoft service health

### Logs Analysis

Techniques for extracting meaningful insights from Cloudflare Workers logs to identify error patterns and performance issues. Effective log analysis is crucial for production debugging.

**Error Pattern Detection:**

```bash
# Find authentication errors
wrangler tail --grep "Authentication.*failed"

# Find Microsoft Graph errors
wrangler tail --grep "Microsoft.*Graph.*error"

# Find performance issues
wrangler tail --grep "timeout\|exceeded"
```

## Maintenance Procedures

### Regular Maintenance Tasks

**Weekly Tasks:**

- Review error logs and performance metrics
- Check Microsoft Graph API error rates
- Verify authentication success rates
- Monitor KV storage usage

**Monthly Tasks:**

- Rotate encryption keys (if required by policy)
- Review and update Microsoft app permissions
- Test disaster recovery procedures
- Update dependencies and security patches

**Quarterly Tasks:**

- Performance benchmark testing
- Security audit and penetration testing
- Documentation updates
- Backup and recovery testing

### Key Rotation

Security procedures for rotating cryptographic keys and secrets without service interruption. Regular key rotation limits the impact of potential credential compromise.

**Rotate Encryption Keys:**

```bash
# Generate new keys
NEW_ENCRYPTION_KEY=$(openssl rand -hex 32)
NEW_COOKIE_KEY=$(openssl rand -hex 32)
NEW_COOKIE_SECRET=$(openssl rand -hex 32)

# Update secrets (zero-downtime)
wrangler secret put ENCRYPTION_KEY --env production
wrangler secret put COOKIE_ENCRYPTION_KEY --env production
wrangler secret put COOKIE_SECRET --env production

# Deploy updated configuration
wrangler deploy --env production
```

**Rotate Microsoft Client Secret:**

1. Create new client secret in Microsoft Entra ID
2. Update Cloudflare Workers secret: `wrangler secret put MICROSOFT_CLIENT_SECRET`
3. Test authentication flow
4. Delete old client secret after verification

### Backup & Recovery

Data protection and disaster recovery procedures for KV storage and configuration. Regular backups ensure business continuity in case of data loss or corruption.

**KV Data Backup:**

```bash
# Backup OAuth data
wrangler kv:key list --namespace-id "oauth-kv-id" > oauth-backup.json

# Backup configuration data
wrangler kv:key list --namespace-id "config-kv-id" > config-backup.json

# Store backups securely (encrypt before storage)
```

**Disaster Recovery:**

1. Deploy worker to new Cloudflare account
2. Recreate KV namespaces
3. Restore KV data from encrypted backups
4. Update DNS records (if using custom domain)
5. Test all functionality

### Scaling Considerations

Strategies for handling increased load and optimizing resource usage as the service grows. Cloudflare's edge infrastructure provides automatic scaling with proper configuration.

**Performance Scaling:**

- Cloudflare Workers automatically scale to handle traffic
- KV storage scales automatically
- Durable Objects scale per-object

**Cost Optimization:**

- Monitor KV read/write operations
- Implement caching for frequently accessed data
- Use appropriate KV TTL values for cached responses

**Geographic Distribution:**

- Workers deploy globally by default
- KV data replicates to edge locations
- Consider data residency requirements for compliance

---

_This deployment guide provides complete operational procedures for running the Microsoft 365 MCP Server in production environments with enterprise-grade reliability and security._

---

## Part III: Advanced Operations

### CI/CD Pipeline

#### GitHub Actions Workflow

Automates testing and deployment to Cloudflare Workers using GitHub Actions CI/CD pipeline. The workflow runs tests on pull requests, validates code quality, and deploys to production when changes are merged to the main branch.

Create `.github/workflows/deploy.yml`:

```yaml
name: Deploy to Cloudflare Workers

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-node@v3
        with:
          node-version: "18"
          cache: "npm"
      - run: npm ci
      - run: npm run type-check
      - run: npm run lint
      - run: npm test -- --coverage
      - uses: codecov/codecov-action@v3
        with:
          token: ${{ secrets.CODECOV_TOKEN }}

  deploy:
    needs: test
    if: github.ref == 'refs/heads/main'
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-node@v3
        with:
          node-version: "18"
      - run: npm ci
      - uses: cloudflare/wrangler-action@v3
        with:
          apiToken: ${{ secrets.CF_API_TOKEN }}
          accountId: ${{ secrets.CF_ACCOUNT_ID }}
          command: deploy
          secrets: |
            MICROSOFT_CLIENT_SECRET
            ENCRYPTION_KEY
            COOKIE_SECRET
        env:
          MICROSOFT_CLIENT_SECRET: ${{ secrets.MICROSOFT_CLIENT_SECRET }}
          ENCRYPTION_KEY: ${{ secrets.ENCRYPTION_KEY }}
          COOKIE_SECRET: ${{ secrets.COOKIE_SECRET }}
```

#### Automated Testing Pipeline

Comprehensive test suite that validates code across multiple Node.js versions and generates coverage reports. This pipeline ensures code quality and compatibility before deployment, running all validation checks in parallel for efficiency.

```yaml
name: Test Suite

on: [push, pull_request]

jobs:
  test:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        node: [18, 20]
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-node@v3
        with:
          node-version: ${{ matrix.node }}
      - run: npm ci
      - run: npm run validate
      - run: npm test -- --coverage
      - name: Upload coverage
        uses: actions/upload-artifact@v3
        with:
          name: coverage-${{ matrix.node }}
          path: coverage/
```

### Maintenance Procedures

Scheduled tasks and procedures for keeping the MCP server secure, performant, and reliable over time. Regular maintenance prevents degradation and security vulnerabilities.

Regular maintenance ensures system reliability, security, and performance over time. These procedures establish a routine for monitoring, updating, and optimizing the deployment while preventing degradation and security vulnerabilities.

#### Regular Maintenance Checklist

**Daily Tasks:**

- [ ] Monitor error rates in Cloudflare Analytics
- [ ] Check API response times
- [ ] Review rate limiting metrics
- [ ] Verify OAuth token refresh success rate

**Weekly Tasks:**

- [ ] Review security logs for anomalies
- [ ] Check KV storage usage
- [ ] Analyze performance metrics
- [ ] Update dependencies (patch versions)

**Monthly Tasks:**

- [ ] Rotate API keys if needed
- [ ] Review and update documentation
- [ ] Performance baseline testing
- [ ] Security audit

**Quarterly Tasks:**

- [ ] Update dependencies (minor/major versions)
- [ ] Review Microsoft Graph API changes
- [ ] Update OAuth scopes if needed
- [ ] Disaster recovery drill

#### Token Rotation Schedule

Token rotation is a critical security practice that limits the impact of potential credential compromise. Regular rotation ensures that even if credentials are exposed, their validity window is limited, reducing the risk of unauthorized access.

**Microsoft Client Secret:**

- **Expiration**: 24 months from creation
- **Rotation Procedure**:
  1. Create new client secret in Azure Portal
  2. Test new secret in development environment
  3. Update production secret: `wrangler secret put MICROSOFT_CLIENT_SECRET`
  4. Verify OAuth flow with new secret
  5. Remove old secret from Azure Portal after 24 hours

**Encryption Keys:**

- **Rotation**: Annually or on security incident
- **Procedure**:
  1. Generate new key: `openssl rand -hex 32`
  2. Update in parallel: old and new keys
  3. Migrate existing encrypted data
  4. Remove old key after migration

#### Dependency Updates

Keeping dependencies updated ensures security patches are applied and new features are available. This systematic approach balances stability with security by testing updates in development before production deployment.

```bash
# Check for outdated packages
npm outdated

# Update patch versions
npm update

# Update minor versions
npm install package@latest

# Update major versions (test thoroughly)
npm install package@next

# Security audit
npm audit
npm audit fix
```

## Production Monitoring

### Key Metrics to Track

Tracking these metrics provides insights into system health, usage patterns, and potential issues. Regular monitoring enables proactive optimization and early detection of problems before they impact users.

#### OAuth Flow Metrics

- **Authorization success rate**: Target >95%
- **Token refresh frequency**: Monitor for anomalies
- **Average token TTL**: Should be ~3600s
- **Failed authentications**: Track by client type

```bash
# Monitor OAuth metrics via Cloudflare Analytics
wrangler tail --format json | jq '. | select(.outcome == "ok") | .logs[] | select(contains("OAuth"))'
```

#### MCP Protocol Metrics

- **Tools invocation frequency**: Identify most-used tools
- **Error rates by tool type**: Focus optimization efforts
- **Discovery vs authenticated requests ratio**: Should be <10% discovery
- **Average response time by tool**: Target <500ms

```bash
# Track MCP tool usage
wrangler tail --format json | jq '. | select(.logs[] | contains("tool:")) | .logs[]'
```

#### Microsoft Graph API Metrics

- **API call latency by endpoint**: Monitor p50, p95, p99
- **Rate limit encounters (429 responses)**: Should be <1%
- **Permission denied errors (403 responses)**: Indicates scope issues
- **Token expiration errors (401 responses)**: Should auto-refresh

```bash
# Monitor Graph API errors
wrangler tail --format json | jq '. | select(.logs[] | contains("Graph API error")) | .logs[]'
```

### Alerting Thresholds

Proactive alerting ensures rapid response to issues before they escalate. These thresholds are based on operational experience and should be adjusted based on your specific usage patterns and SLA requirements.

Configure alerts for critical conditions:

| Metric                       | Warning | Critical | Action                         |
| ---------------------------- | ------- | -------- | ------------------------------ |
| Token refresh failures       | >5%     | >10%     | Check Microsoft service health |
| Graph API 429 responses      | >10/min | >50/min  | Implement request throttling   |
| Durable Object evictions     | >0      | >10/min  | Scale to multiple DOs          |
| OAuth authorization failures | >10%    | >25%     | Review client configuration    |
| Average response time        | >1s     | >3s      | Optimize slow endpoints        |
| Error rate                   | >1%     | >5%      | Check logs for root cause      |

### Real-time Monitoring Setup

Real-time monitoring provides immediate visibility into system behavior and performance. Cloudflare Analytics Engine enables custom metrics collection and analysis for deep insights into application-specific patterns.

```bash
# Set up Cloudflare Analytics Engine
wrangler analytics init

# Configure custom metrics
cat > analytics-config.json << EOF
{
  "datasets": [
    {
      "name": "mcp_metrics",
      "dimensions": ["tool", "client_type", "error_code"],
      "measures": ["latency", "count"]
    }
  ]
}
EOF

# Deploy analytics configuration
wrangler analytics push analytics-config.json
```

### Dashboard Configuration

Visualization dashboards consolidate metrics into actionable insights. These configurations provide at-a-glance system health monitoring and enable quick identification of trends or anomalies requiring attention.

Create monitoring dashboard using Grafana or Datadog:

```yaml
# grafana-dashboard.yml
panels:
  - title: "OAuth Success Rate"
    query: "sum(rate(oauth_success[5m])) / sum(rate(oauth_attempts[5m]))"

  - title: "Tool Usage Heatmap"
    query: "sum by(tool) (rate(mcp_tool_invocations[5m]))"

  - title: "Graph API Latency"
    query: "histogram_quantile(0.95, rate(graph_api_latency_bucket[5m]))"

  - title: "Error Rate by Type"
    query: "sum by(error_type) (rate(errors_total[5m]))"
```

### Health Checks

Automated health checks verify system availability and functionality at regular intervals. These checks detect issues before users report them and can trigger alerts for immediate intervention when problems occur.

Implement automated health checks:

```javascript
// health-check.js
async function healthCheck() {
  const checks = {
    oauth: await checkOAuthEndpoint(),
    discovery: await checkDiscovery(),
    graphApi: await checkGraphConnection(),
    durableObjects: await checkDurableObjects(),
  };

  const healthy = Object.values(checks).every((c) => c.status === "ok");
  return {
    healthy,
    checks,
    timestamp: new Date().toISOString(),
  };
}

// Schedule health checks every 5 minutes
addEventListener("scheduled", (event) => {
  event.waitUntil(healthCheck());
});
```

### Log Aggregation

Centralized logging consolidates logs from distributed Workers into a single searchable repository. This enables correlation of events across multiple requests and provides historical data for debugging and compliance requirements.

Set up centralized logging:

```bash
# Export logs to external service
wrangler tail --format json | \
  jq -c '. + {service: "m365-mcp-server"}' | \
  curl -X POST https://logs.example.com/ingest \
    -H "Content-Type: application/json" \
    -d @-

# Create log retention policy
cat > log-retention.toml << EOF
[env.production]
logpush = true
logpush_dataset = "m365_mcp_logs"
logpush_retention = 30  # days
EOF
```

### Performance Profiling

Performance profiling identifies slow operations and resource-intensive code paths. Using browser-standard Performance API markers, you can measure specific operations and optimize based on real production data.

Monitor performance bottlenecks:

```javascript
// Add performance markers
performance.mark("graph-api-start");
const result = await graphClient.getEmails(token, params);
performance.mark("graph-api-end");
performance.measure("graph-api-call", "graph-api-start", "graph-api-end");

// Log performance metrics
const measure = performance.getEntriesByName("graph-api-call")[0];
console.log(`Graph API call took ${measure.duration}ms`);
```

### Scaling & Performance

#### Load Testing

Performance testing using Artillery to simulate production traffic patterns and identify bottlenecks. The load test gradually ramps up from 10 to 50 concurrent users over several minutes, measuring response times and error rates under stress.

Using Artillery for load testing:

```bash
# Install artillery
npm install -g artillery

# Create test scenario
cat > load-test.yml << EOF
config:
  target: "https://m365-mcp-server.workers.dev"
  phases:
    - duration: 60
      arrivalRate: 10
    - duration: 120
      arrivalRate: 50
scenarios:
  - name: "Tool Discovery"
    flow:
      - post:
          url: "/sse"
          json:
            jsonrpc: "2.0"
            id: 1
            method: "tools/list"
            params: {}
```
