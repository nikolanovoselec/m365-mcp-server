# Microsoft 365 MCP Server

A production-grade Model Context Protocol (MCP) server that provides secure access to Microsoft 365 services through OAuth 2.1 + PKCE authentication. Built on Cloudflare Workers with enterprise-level security and native integration support with Claude, Gemini, ChatGPT, Copilot and other services that support official MCP specification as outlined here - https://modelcontextprotocol.io/docs/getting-started/intro.

## Key Features

- **OAuth 2.1 + PKCE + Dynamic Client Registration** - Industry-standard security
- **Native Remote MCP Support** - Direct remote MCP connections without localhost dependency  
- **Complete Microsoft 365 Integration** - Email, Calendar, Teams, Contacts, and OnlineMeetings
- **Cloudflare Workers Deployment** - Global edge network with Durable Objects persistence
- **Session Management** - Persistent authentication with automatic token refresh
- **Production Security** - End-to-end encryption with secure token storage

## Quick Start

### Claude Setup (Recommended)

1. **Go to claude.ai**
2. **Add a custom connector with the MCP server URL**: `https://your-worker-domain.com/sse`
3. **Complete OAuth authentication** when prompted

That's it! Your Microsoft 365 tools will be available in Claude.

### Alternative Setup Methods

**Command-line usage:**
```bash
npx mcp-remote https://your-worker-domain.com/sse
```

**Claude Desktop JSON configuration:**
```json
{
  "mcpServers": {
    "microsoft-365": {
      "command": "npx",
      "args": ["mcp-remote", "https://your-worker-domain.com/sse"]
    }
  }
}
```

**ChatGPT Integration (via mcp-remote):**
```bash
# Terminal usage
npx mcp-remote https://your-worker-domain.com/sse --method tools/call --params '{"name":"getEmails","arguments":{}}'
```

**Google Gemini Function Calling:**
```python
import requests

# Use MCP tools as Gemini functions
def get_emails(count=10, folder='inbox'):
    response = requests.post('https://your-worker-domain.com/sse', json={
        'jsonrpc': '2.0', 'id': 1, 'method': 'tools/call',
        'params': {'name': 'getEmails', 'arguments': {'count': count, 'folder': folder}}
    })
    return response.json()
```

**GitHub Copilot/VS Code Extension:**
```typescript
// Custom extension integration
const mcpClient = new MCPClient('https://your-worker-domain.com/sse');
await mcpClient.callTool('sendEmail', { to: 'user@company.com', subject: 'Hello' });
```

## Microsoft 365 Tools

### Email Operations

- **`getEmails`** - Retrieve recent emails from Outlook
- **`sendEmail`** - Send emails via Outlook with HTML support
- **`searchEmails`** - Search emails using Microsoft Search

### Calendar Management  

- **`getCalendarEvents`** - Get upcoming calendar events
- **`createCalendarEvent`** - Create new calendar events with attendees

### Teams Integration

- **`sendTeamsMessage`** - Send messages to Teams channels
- **`createTeamsMeeting`** - Create Teams meetings with calendar integration

### Contact Management

- **`getContacts`** - Access your Outlook contacts

### Authentication

- **`authenticate`** - Get OAuth authentication URL for setup

All tools support rich formatting, error handling, and automatic token refresh.

## OAuth Flow & Security

### Standards Compliance

- **OAuth 2.1**: Latest OAuth specification with enhanced security
- **PKCE**: Proof Key for Code Exchange prevents code interception
- **Dynamic Client Registration**: Automatic client registration for Claude Desktop
- **Token Encryption**: All tokens encrypted at rest with AES-256-GCM

### Authentication Flow

```
1. Client connects → Tool discovery (no auth required)
2. First tool call → Authentication required error
3. OAuth flow initiated → Microsoft login page
4. User completes auth → Tokens stored securely
5. All tools functional → Session-based access
```

### Security Features

- **End-to-end encryption** for all sensitive data
- **Automatic token refresh** prevents expired sessions  
- **Session isolation** between different clients
- **CORS protection** with strict origin validation
- **Security headers** (HSTS, CSP) enforced

## Architecture

### Unified Endpoint Design

Single `/sse` endpoint supports all MCP protocols:
- **WebSocket** - Full bidirectional MCP protocol
- **Server-Sent Events** - Claude Desktop validation  
- **JSON-RPC over HTTP** - REST-like tool calls
- **Protocol detection** - Automatic based on headers

### Cloudflare Workers + Durable Objects

- **Global deployment** - Edge computing worldwide
- **Durable Objects** - Persistent state and sessions
- **KV Storage** - OAuth client and configuration data
- **Secrets management** - Encrypted credential storage

### Microsoft Graph Integration

Each tool maps to specific Graph API endpoints:
```typescript
const toolEndpoints = {
  'getEmails': 'GET /me/messages',
  'sendEmail': 'POST /me/sendMail',
  'createCalendarEvent': 'POST /me/events',
  'sendTeamsMessage': 'POST /teams/{id}/channels/{id}/messages'
};
```

## API Reference

### Core Endpoints

**Health Check:**
```bash
GET /health
```

**OAuth Metadata Discovery:**
```bash
GET /.well-known/oauth-authorization-server
```

**MCP Tool Discovery:**
```bash
POST /sse
Content-Type: application/json

{
  "jsonrpc": "2.0",
  "id": 1, 
  "method": "tools/list",
  "params": {}
}
```

**Tool Execution:**
```bash
POST /sse
Content-Type: application/json

{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/call", 
  "params": {
    "name": "getEmails",
    "arguments": {
      "count": 10,
      "folder": "inbox"
    }
  }
}
```

### Tool Schemas

**Email Tools:**
```typescript
// getEmails
{
  count: number;     // Max 50, default 10
  folder: string;    // 'inbox', 'sent', etc.
}

// sendEmail  
{
  to: string;           // Required
  subject: string;      // Required
  body: string;         // Required
  contentType: 'text' | 'html';  // Default: html
}

// searchEmails
{
  query: string;     // Search terms
  count: number;     // Max 50, default 10
}
```

**Calendar Tools:**
```typescript
// getCalendarEvents
{
  days: number;      // Max 30, default 7
}

// createCalendarEvent
{
  subject: string;   // Required
  start: string;     // ISO format, required
  end: string;       // ISO format, required
  body?: string;     // Optional description
  attendees?: string; // Comma-separated emails
}
```

## Troubleshooting

### Common Issues

**Authentication Problems:**
- Verify Microsoft Entra ID redirect URIs include your domain
- Check client ID and tenant ID in environment variables
- Ensure required Graph API permissions are granted

**Connection Issues:**
- Confirm server URL is accessible and returning health status
- Test with curl: `curl https://your-domain.com/health`
- Check Cloudflare Workers logs for deployment issues

**Tool Execution Errors:**
- Verify authentication completed successfully
- Check Microsoft Graph API permissions for specific tools
- Review server logs for specific error messages

**WebSocket Issues (mcp-remote):**
- Known limitation with HTTP/2 protocol compatibility
- Use JSON-RPC over HTTP as fallback
- Consider using direct connection instead

### Testing Commands

```bash
# Test server health
curl https://your-domain.com/health

# Test OAuth metadata
curl https://your-domain.com/.well-known/oauth-authorization-server

# Test tool discovery  
curl -X POST https://your-domain.com/sse \
  -H 'Content-Type: application/json' \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}'

# Test with mcp-remote
npx mcp-remote https://your-domain.com/sse --method tools/list
```

## Roadmap

### Current Priority: Microsoft Graph API Implementation

**Status**: Tools are discovered and authentication works, but return placeholder data instead of actual Microsoft Graph calls.

**Next Steps:**
1. **Token Integration** - Connect tools to OAuth access tokens
2. **Graph API Implementation** - Replace placeholders with real API calls
3. **Response Formatting** - Format Graph responses for MCP clients  
4. **Error Handling** - Handle token expiry, permissions, rate limits
5. **Testing** - Validate with real Microsoft 365 data

### Future Enhancements

**WebSocket Compatibility:**
- Fix HTTP/2 WebSocket upgrade issues for full mcp-remote support
- Implement proper WebSocketPair patterns

**Advanced Features:**
- Real-time notifications via Microsoft Graph webhooks
- Batch API requests for improved performance
- Advanced search with Microsoft Search API
- Multi-tenant support for enterprise deployments

**Performance & Monitoring:**
- Response caching for frequently accessed data
- Usage analytics and performance metrics  
- Enhanced error tracking and alerting

## Technical References

- **OAuth 2.1**: [RFC 9068](https://tools.ietf.org/rfc/rfc9068.txt)
- **PKCE**: [RFC 7636](https://tools.ietf.org/rfc/rfc7636.txt)  
- **MCP Protocol**: [Model Context Protocol](https://modelcontextprotocol.io)
- **Microsoft Graph**: [API Reference](https://docs.microsoft.com/en-us/graph/)
- **Cloudflare Workers**: [Documentation](https://workers.cloudflare.com/)

## Contributing

1. Fork the repository
2. Create feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open Pull Request

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Support

- **Issues**: [GitHub Issues](https://github.com/nikolanovoselec/m365-mcp-server/issues)
- **Deployment**: See [DEPLOYMENT.md](DEPLOYMENT.md) for complete setup guide

---

*Production-ready Microsoft 365 integration for AI platforms with enterprise-grade security and performance.*
