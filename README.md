# Remote MCP Server built on Cloudflare Workers with Microsoft 365 Integration

A production-grade Model Context Protocol (MCP) server providing secure remote access to Microsoft 365 services through OAuth 2.1 + PKCE authentication. Features enterprise-grade security, global edge deployment, and native integration support for MCP-compatible AI applications.

## Why did I build this?

I started this project to explore three buzzworthy technologies: MCP (Model Context Protocol), Cloudflare Workers, and this whole "agentic AI" thing everyone keeps talking about. Just wanted to see what the fuss was about, maybe build a simple proof of concept during the evening after the kids fall asleep. Fast forward a weekend and 2 sleepless nights, and I've accidentally become an OAuth 2.1 expert, can explain PKCE flows in my sleep, dream in HMAC-SHA256 signatures, and have strong opinions about JWT token expiration strategies. What started as "just connect [insert your AI assistant of choice] to my calendar" turned into a full-blown production-ready authentication system with more security layers than my bank. But hey, at least now my AI assistant can read my emails securely while I ponder how a simple curiosity led to implementing half of RFC 6749. The good news? You get to benefit from my OAuth rabbit hole adventure with a server that actually works and won't leak your tokens all over the internet.

**Full Disclosure:** Because I'm a lazy cybersecurity architect (the good kind - the one who automates everything), this code is 80% AI-generated and the documentation is 99.7% AI-generated. It took me longer to write these "About" sections with my actual human fingers than it took the AI to generate 5000+ lines of working code and documentation. I'm not ashamed - I'm efficient. Bill Gates would be proud (you know, that quote about hiring lazy people because they find easy ways to do hard things). The 0.3% of documentation I wrote myself? You're reading it right now. You're welcome.

## Table of Contents

- [What is Model Context Protocol (MCP)?](#what-is-model-context-protocol-mcp)
- [Why Microsoft 365 Integration?](#why-microsoft-365-integration)
- [Prerequisites](#prerequisites)
- [Key Features](#key-features)
- [Quick Start Guide](#quick-start-guide)
  - [5-Minute Setup](#5-minute-setup)
  - [Option 1: AI Assistant Integration](#option-1-ai-assistant-integration-recommended)
  - [Option 2: mcp-remote Integration](#option-2-mcp-remote-integration)
  - [Option 3: Web-Based Access](#option-3-web-based-access)
- [Microsoft 365 Tools](#microsoft-365-tools)
- [Common Use Cases](#common-use-cases)
- [System Architecture](#system-architecture)
- [Troubleshooting](#troubleshooting)
- [Frequently Asked Questions](#frequently-asked-questions)
- [Documentation](#documentation)
- [Development](#development)
- [Deployment](#deployment)
- [Security](#security)
- [Contributing](#contributing)
- [Support](#support)

## What is Model Context Protocol (MCP)?

Remember when your apps couldn't talk to each other and you had to copy-paste everything like it's 1999? MCP is basically couples therapy for AI and your APIs - helping them communicate their needs, respect each other's boundaries (rate limits), and work through their authentication issues. And like any good therapist, MCP comes equipped with a Swiss Army knife of tools - except instead of a tiny scissors you'll never use, you get actual useful functions that let your AI read emails, check calendars, and do real work while you grab another coffee and contemplate how you ended up mediating between machines.

Model Context Protocol (MCP) is an open protocol developed by Anthropic that enables seamless integration between Large Language Model (LLM) applications and external data sources and tools. And yes, as a small thank you to Anthropic for actually open-sourcing this protocol (looking at you, other AI companies hoarding your toys), I tested this integration with Claude first and foremost. The other assistants can get in line behind the one that actually shares its homework with the class. It provides a standardized way for AI models to:

- **Access External Tools**: Execute functions and operations in external systems
- **Retrieve Resources**: Fetch data from APIs, databases, and services
- **Maintain Context**: Preserve session state across interactions
- **Ensure Security**: Handle authentication and authorization properly

MCP servers act as bridges between AI applications and external services, exposing tools and resources that the AI can use to perform real-world tasks.

## Why Microsoft 365 Integration?

Honestly? Because I already had a Microsoft 365 Business subscription for my actual business, and it felt wrong to have all those API endpoints just sitting there, unused and unloved. Turns out, building an OAuth integration for Microsoft's enterprise ecosystem is like solving a Rubik's cube blindfolded - technically possible, surprisingly satisfying when it works, and it makes you look way smarter than you actually are at dinner parties. 

The beautiful part is that this could have been ANY service - Google Workspace, Slack, Discord, whatever has OAuth 2.1. The architecture I accidentally over-engineered is completely service-agnostic. Just swap out the Microsoft Graph endpoints for any other OAuth-compatible API, and boom - you've got yourself an MCP server for your favorite service. But since I had M365 lying around and those Exchange endpoints were calling my name...

This MCP server enables AI applications to interact with Microsoft 365 services, providing:

- **Email Management**: Read, send, and search emails through Outlook
- **Calendar Operations**: Create events, check availability, manage meetings
- **Teams Integration**: Send messages, create meetings, collaborate
- **Contact Access**: Search and manage Microsoft 365 contacts
- **Secure Authentication**: Enterprise-grade OAuth 2.1 + PKCE flow
- **Real-time Operations**: Direct API access, not cached data

## Prerequisites

Before setting up the Microsoft 365 MCP Server, ensure you have:

### Required Accounts

- [ ] **Microsoft 365 Account** (Business or Enterprise)
  - Admin access for app registration
  - Active subscription with Exchange Online
  
- [ ] **Microsoft Entra ID Access** (formerly Azure AD)
  - Ability to register applications
  - Permission to grant admin consent
  
- [ ] **Cloudflare Account** (Free tier supported)
  - Workers & KV storage enabled
  - Custom domain (optional for production)

### Required Permissions

Your Microsoft 365 administrator must grant these Graph API permissions:
- `User.Read` - Read user profile
- `Mail.Read`, `Mail.ReadWrite`, `Mail.Send` - Email operations
- `Calendars.Read`, `Calendars.ReadWrite` - Calendar access
- `Contacts.ReadWrite` - Contact management
- `OnlineMeetings.ReadWrite` - Teams meetings
- `ChannelMessage.Send` - Teams messages
- `Team.ReadBasic.All` - Teams information

### Development Tools

- **Node.js 18+** with npm
- **Git** for version control
- **Wrangler CLI** (`npm install -g wrangler`)

## Key Features

- **Enterprise Security**: OAuth 2.1 + PKCE with dynamic client registration
- **Global Edge Network**: Deployed on Cloudflare Workers with 330+ edge locations
- **Hybrid Protocol Support**: Single endpoint supporting WebSocket, SSE, and HTTP JSON-RPC
- **Complete Microsoft 365 Integration**: Email, Calendar, Teams, Contacts via Microsoft Graph API
- **Native MCP Compatibility**: Direct integration with AI assistants and mcp-remote
- **Production Security**: End-to-end encryption with secure token storage
- **Automatic Token Management**: Handles refresh tokens and session persistence
- **Real-time Operations**: Direct API access without caching delays

## Documentation

| Document                                  | Description                                                   |
| ----------------------------------------- | ------------------------------------------------------------- |
| **[Technical Reference](TECHNICAL.md)**   | Complete technical documentation, architecture, and API docs |
| **[Operations Guide](OPERATIONS.md)**     | Development, deployment, and maintenance procedures          |
| **README.md**                             | This file - overview, quick start, and troubleshooting      |

## Quick Start Guide

### 5-Minute Setup

1. **Check Prerequisites**: Ensure you have all required accounts and permissions
2. **Clone Repository**: `git clone https://github.com/nikolanovoselec/m365-mcp-server.git`
3. **Configure Azure**: Register app in Microsoft Entra ID
4. **Deploy to Cloudflare**: Run `wrangler deploy`
5. **Test Connection**: Use one of the three integration methods below

## Quick Diagnostics

If you encounter issues, run these diagnostic commands:

```bash
# Check OAuth Provider health
curl -X GET https://your-worker.workers.dev/.well-known/openid-configuration

# Verify Durable Objects binding
wrangler tail --format json | grep "MCP_OBJECT"

# Test discovery phase (should return tools without auth)
curl -X POST https://your-worker.workers.dev/mcp \
  -H "Authorization: Bearer test-token" \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","method":"tools/list","id":1}'

# Check KV namespaces configuration
wrangler kv namespace list

# Monitor real-time logs
wrangler tail --format pretty

# Test OAuth flow (manual)
open "https://your-worker.workers.dev/authorize?client_id=YOUR_CLIENT_ID&redirect_uri=YOUR_REDIRECT_URI&response_type=code&scope=openid"

# Validate Microsoft Graph token
curl -X GET https://graph.microsoft.com/v1.0/me \
  -H "Authorization: Bearer YOUR_ACCESS_TOKEN"
```

### Common Issues Quick Fixes

| Issue | Quick Fix |
|-------|-----------|
| 404 on /sse endpoint | User not authenticated - complete OAuth flow first |
| 401 from Graph API | Token expired - OAuth Provider should auto-refresh |
| 403 from Graph API | Missing permissions - check Azure AD app scopes |
| 429 rate limiting | Implement exponential backoff in your client |
| WebSocket fails | Expected - Cloudflare uses HTTP/2, fallback to SSE |
| Empty discovery | Check Durable Object binding in wrangler.toml |

### Option 1: AI Assistant Integration (Recommended)

The simplest way to connect your AI assistant to Microsoft 365. Just add the MCP server URL as a custom connector in any MCP-compatible AI assistant.

1. Open your AI assistant's settings and navigate to Connectors or Extensions
2. Add a custom MCP connector with this URL:
   ```
   https://your-worker-domain.com/sse
   ```
3. Authenticate when prompted (will open Microsoft login in your browser)
4. And voilà! Now you can ask your AI to email your mom those meeting notes you keep promising to send, schedule that dentist appointment you've been avoiding, or find that one email from 2023 with the important attachment you swear you didn't delete.

**Compatibility:** Works with any AI assistant that supports the MCP protocol standard. Tested with various MCP-compatible desktop and web clients.

### Option 2: mcp-remote Integration

For command-line access and automation (Note: This integration is not fully implemented yet):

```bash
# 1. Install mcp-remote globally
npm install -g @modelcontextprotocol/remote

# 2. Configure connection
mcp-remote add m365 https://your-worker-domain.com/sse

# 3. Authenticate (will open browser)
mcp-remote auth m365

# 4. Test connection
mcp-remote call m365 tools/list

# 5. Use tools
mcp-remote call m365 tools/call '{"name": "getEmails", "arguments": {"count": 10}}'
```

### Option 3: Web-Based Access

For web applications and services that need programmatic Microsoft 365 access:

```bash
# 1. Register your web application
curl -X POST https://your-worker-domain.com/register \
  -H "Content-Type: application/json" \
  -d '{
    "client_name": "Your Application",
    "redirect_uris": ["https://your-app.com/api/mcp/auth_callback"],
    "client_uri": "https://your-app.com",
    "grant_types": ["authorization_code"],
    "response_types": ["code"]
  }'

# 2. Start OAuth flow - redirect user to this URL
https://your-worker-domain.com/authorize?client_id=YOUR_CLIENT_ID&response_type=code&redirect_uri=YOUR_CALLBACK&scope=User.Read

# 3. Exchange code for token (after user authorizes)
curl -X POST https://your-worker-domain.com/token \
  -d "grant_type=authorization_code&code=AUTH_CODE&client_id=YOUR_CLIENT_ID"

# 4. Use tools with access token
curl -X POST https://your-worker-domain.com/sse \
  -H "Authorization: Bearer YOUR_ACCESS_TOKEN" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"getEmails","arguments":{"count":10}}}'
```

## Microsoft 365 Tools

### Email Operations

- **`sendEmail`** - Send emails via Outlook
  - Supports HTML and plain text
  - Multiple recipients (to, cc, bcc)
  - File attachments support
  
- **`getEmails`** - Retrieve emails from folders
  - Configurable count (max 50)
  - Folder selection (inbox, sent, drafts)
  - Returns sender, subject, body, date
  
- **`searchEmails`** - Search emails with queries
  - Microsoft Graph search syntax
  - Full-text search across all folders
  - Advanced filters support

### Calendar Management

- **`getCalendarEvents`** - List calendar events
  - Date range filtering
  - Returns title, attendees, location
  - Includes online meeting links
  
- **`createCalendarEvent`** - Create new events
  - Set title, description, location
  - Add multiple attendees
  - Configure reminders
  - Create Teams meetings

### Teams Integration

- **`sendTeamsMessage`** - Post to Teams channels
  - Channel and team selection
  - Rich text formatting
  - Mentions support
  
- **`createTeamsMeeting`** - Schedule Teams meetings
  - Set date, time, duration
  - Add attendees
  - Generate meeting links
  - Configure meeting options

### Contact Management

- **`getContacts`** - Access Microsoft 365 contacts
  - Search by name or email
  - Returns full contact details
  - Pagination for large lists

## Common Use Cases

### Email Automation

**Daily Summary Report**
```javascript
// Fetch recent emails
const emails = await mcp.call('getEmails', { count: 20, folder: 'inbox' });

// Generate summary
const summary = emails.map(e => `${e.from}: ${e.subject}`).join('\n');

// Send report
await mcp.call('sendEmail', {
  to: 'manager@company.com',
  subject: 'Daily Email Summary',
  body: `<h2>Today's Emails</h2><pre>${summary}</pre>`,
  contentType: 'html'
});
```

### Calendar Scheduling

**Meeting Coordination**
```javascript
// Check availability
const events = await mcp.call('getCalendarEvents', {
  startDateTime: '2025-01-10T09:00:00',
  endDateTime: '2025-01-10T17:00:00'
});

// Find free slot
const freeSlot = findAvailableTime(events);

// Schedule meeting
await mcp.call('createCalendarEvent', {
  subject: 'Project Review',
  startDateTime: freeSlot.start,
  endDateTime: freeSlot.end,
  attendees: ['team@company.com'],
  isOnlineMeeting: true
});
```

### Teams Notifications

**Automated Status Updates**
```javascript
// Send project update
await mcp.call('sendTeamsMessage', {
  teamId: 'project-team-id',
  channelId: 'general',
  message: 'Deployment completed successfully\n\nVersion 2.0.1 is now live in production.'
});
```

### Contact Search

**Quick Contact Lookup**
```javascript
// Search for contact
const contacts = await mcp.call('getContacts', {
  search: 'John Doe'
});

// Get email and phone
const contact = contacts[0];
console.log(`Email: ${contact.emailAddresses[0].address}`);
console.log(`Phone: ${contact.mobilePhone}`);
```

## System Architecture

### Hybrid Protocol Support

The server implements intelligent protocol detection:

- **WebSocket** - Full bidirectional MCP protocol for mcp-remote clients
- **Server-Sent Events** - Streaming responses for web connectors
- **JSON-RPC over HTTP** - Direct API testing and debugging
- **Discovery Methods** - Unauthenticated tool/resource enumeration

### Infrastructure

Built on **Cloudflare Workers** with:

- **Durable Objects** - Session persistence and WebSocket handling
- **KV Storage** - Three-tier architecture (OAuth, Config, Cache)
- **Edge Computing** - Global distribution with <200ms response times
- **Enterprise Security** - OAuth 2.1 + PKCE compliance

### Microsoft Graph Integration

Direct API mapping to Microsoft Graph endpoints:

- **Real-time Operations** - Send emails, manage calendar, Teams integration
- **Automatic Token Management** - Refresh tokens, scope validation
- **Error Resilience** - Advanced retry logic and fallback handling
- **Response Optimization** - Edge caching and intelligent parsing

## Troubleshooting

### Authentication Issues

**Problem: "401 Unauthorized" errors**
- **Cause**: Token expired or invalid
- **Solution**: 
  1. Check redirect URI matches exactly in Azure app registration
  2. Verify client secret hasn't expired (check Azure portal)
  3. Ensure admin consent granted for all permissions
  4. Try re-authenticating with `mcp-remote auth m365`

**Problem: "403 Forbidden" on specific operations**
- **Cause**: Missing Microsoft Graph permissions
- **Solution**:
  1. Check required scopes in error message
  2. Add permissions in Azure portal
  3. Grant admin consent
  4. Re-authenticate to get new token with added scopes

### Connection Issues

**Problem: WebSocket connection fails**
- **Cause**: HTTP/2 header issues or firewall
- **Solution**:
  1. Check if behind corporate proxy
  2. Try HTTP JSON-RPC endpoint instead
  3. Verify Cloudflare Worker is deployed
  4. Check browser console for specific errors

**Problem: Tools not appearing in AI assistant**
- **Cause**: MCP server not properly configured
- **Solution**:
  1. Verify config file location and syntax
  2. Restart your AI assistant application completely
  3. Check AI assistant logs for errors
  4. Test with `mcp-remote` CLI first

### Microsoft Graph Errors

**Problem: "Insufficient privileges to complete the operation"**
- **Cause**: Missing admin consent or scope
- **Solution**:
  1. Login to Azure portal as admin
  2. Navigate to App registrations → API permissions
  3. Click "Grant admin consent"
  4. Wait 5-10 minutes for propagation

**Problem: "The mailbox is either inactive, soft-deleted, or is hosted on-premise"**
- **Cause**: Exchange Online not configured
- **Solution**:
  1. Verify Microsoft 365 subscription includes Exchange
  2. Check user has Exchange license assigned
  3. Ensure mailbox is cloud-hosted, not on-premise

### Rate Limiting

**Problem: "429 Too Many Requests"**
- **Cause**: Microsoft Graph API rate limit (2000/min)
- **Solution**:
  1. Implement exponential backoff
  2. Batch API requests when possible
  3. Cache responses in KV storage
  4. Spread requests over time

## Frequently Asked Questions

### General Questions

**Q: What is the difference between this and Microsoft's official Graph SDK?**
A: This is an MCP server that bridges AI applications to Microsoft Graph. It provides a protocol layer that AI models can understand and use, whereas the SDK is for direct programmatic access.

**Q: Can I use this with ChatGPT or other AI models?**
A: Currently designed for MCP-compatible clients (AI assistant applications, mcp-remote). Other AI models would need an MCP client implementation.

**Q: Is this free to use?**
A: The server code is open source. You'll need your own Microsoft 365 subscription and Cloudflare account (free tier supported).

### Security Questions

**Q: How are my Microsoft credentials stored?**
A: Credentials are never stored. OAuth tokens are encrypted with AES-256-GCM and stored in Cloudflare KV with automatic expiration.

**Q: Can others access my emails through this server?**
A: No. Each user authenticates separately and can only access their own Microsoft 365 data.

**Q: What data passes through Cloudflare?**
A: Only encrypted tokens and API responses. Cloudflare cannot decrypt your Microsoft data.

### Technical Questions

**Q: Why do I need Cloudflare Workers?**
A: Cloudflare provides the infrastructure for WebSocket handling, global edge distribution, and Durable Objects for session management.

**Q: Can I self-host this?**
A: Not directly. The architecture requires Cloudflare Workers specific features. You could adapt it for other platforms but would need significant changes.

**Q: What's the latency like?**
A: Typically <200ms for cached operations, <500ms for Microsoft Graph API calls, depending on your location.

**Q: Can I add custom tools?**
A: Yes! See OPERATIONS.md for development guide on adding new tools.

### Limitations

**Q: What are the rate limits?**
A: Microsoft Graph: 2000 requests/min per app. Cloudflare Workers: 100,000 requests/day on free tier.

**Q: Maximum email attachment size?**
A: 25MB per email (Microsoft Graph limitation).

**Q: How many concurrent users?**
A: Unlimited via Cloudflare's Durable Objects architecture.

## Quick Links

- **Live API**: `https://your-worker-domain.com`
- **Tool Discovery**: `POST /sse` with `{"method": "tools/list"}`
- **Repository**: [GitHub](https://github.com/nikolanovoselec/m365-mcp-server)
- **Issues**: [Report Issues](https://github.com/nikolanovoselec/m365-mcp-server/issues)

## Development

### Local Setup

```bash
# Clone and setup
git clone https://github.com/nikolanovoselec/m365-mcp-server.git
cd m365-mcp-server
npm install

# Configure environment
cp .dev.vars.example .dev.vars
cp wrangler.example.toml wrangler.toml
# Edit files with your Microsoft 365 and Cloudflare credentials

# Start development server
npm run dev
```

### Testing

```bash
# Run all tests
npm test

# Type checking
npm run type-check

# Linting
npm run lint
```

See [OPERATIONS.md](OPERATIONS.md) for complete development workflow.

## Deployment

The server is production-ready and deployed on Cloudflare Workers. For your own deployment:

```bash
# Deploy to Cloudflare Workers
wrangler deploy

# Set production secrets
wrangler secret put MICROSOFT_CLIENT_SECRET
wrangler secret put ENCRYPTION_KEY
```

See [OPERATIONS.md](OPERATIONS.md) for complete deployment guide.

## Requirements

- **Microsoft 365 Business/Enterprise account** with admin access
- **Microsoft Entra ID app registration** with appropriate permissions
- **Cloudflare Workers account** (Free tier supported)
- **Node.js 18+** for development

## Security

- **OAuth 2.1 + PKCE** compliance with S256 challenge method
- **End-to-end encryption** for all tokens and sensitive data
- **Zero client secrets** for public clients (mcp-remote compatibility)
- **HMAC-signed approval cookies** for session management
- **Input sanitization** and XSS protection
- **Rate limiting** and abuse prevention

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Contributing

We welcome contributions! Please read [OPERATIONS.md](OPERATIONS.md) for:

- Development setup and workflow
- Code standards and testing procedures
- Pull request requirements
- Issue reporting guidelines

## Support

- **Documentation**: Complete guides in this repository
- **Issues**: [GitHub Issues](https://github.com/nikolanovoselec/m365-mcp-server/issues)
- **Discussions**: [GitHub Discussions](https://github.com/nikolanovoselec/m365-mcp-server/discussions)

---

Built with ❤️ for the Model Context Protocol ecosystem. Empowering AI applications with secure, production-grade Microsoft 365 integration.