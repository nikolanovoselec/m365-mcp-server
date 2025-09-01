# Microsoft 365 MCP Server for Claude Desktop

Connect Claude Desktop to your Microsoft 365 MCP Server running on Cloudflare Workers.

## Quick Setup (Web Connector - Recommended)

### Option 1: Use Claude Desktop Web Connectors

1. **Open Claude Desktop**
2. **Go to Settings → Connectors**
3. **Click "Add Connector"**
4. **Enter connector details:**
   - **Name**: Microsoft 365
   - **URL**: `https://your-worker-domain.com/sse`
5. **Click "Connect"** 
6. **Complete OAuth authentication** when prompted

That's it! Your Microsoft 365 tools will be available in Claude.

## Alternative Setup (JSON Configuration)

### Option 2: Legacy JSON Configuration

If web connectors are not available, you can use the traditional JSON configuration method:

#### Step 1: Install
```bash
git clone /path/to/this/repository
cd m365-mcp-desktop
npm install
npm run build
```

#### Step 2: Get OAuth Access Token

1. Register OAuth client:
```bash
curl -X POST https://your-worker-domain.com/register \
  -H "Content-Type: application/json" \
  -d '{"client_name":"claude-desktop","redirect_uris":["http://localhost:8080/callback"]}'
```

2. Complete OAuth flow using the returned `client_id` and `client_secret`

3. Get your OAuth access token

#### Step 3: Configure Claude Desktop

**For direct connection (recommended):**

1. **Locate your Claude Desktop config file:**
   - **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`  
   - **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

2. **Add this configuration:**
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

**For local connector (advanced):**
```json
{
  "mcpServers": {
    "microsoft-365": {
      "command": "node",
      "args": ["/path/to/m365-mcp-desktop/dist/index.js"],
      "env": {
        "M365_ACCESS_TOKEN": "your_oauth_access_token_here",
        "M365_SERVER_URL": "https://your-worker-domain.com"
      }
    }
  }
}
```

3. **Restart Claude Desktop** for changes to take effect

## Available Tools

Once connected, you'll have access to:

### Email Tools
- **sendEmail** - Send emails via Outlook
- **getEmails** - Get recent emails from inbox
- **searchEmails** - Search emails by content

### Calendar Tools  
- **getCalendarEvents** - View upcoming calendar events
- **createCalendarEvent** - Create new calendar events

### Teams Tools
- **sendTeamsMessage** - Send messages to Teams channels
- **createTeamsMeeting** - Create Teams meetings

### Contact Tools
- **getContacts** - Access your Outlook contacts

## Available Resources

- **microsoft://profile** - Your Microsoft 365 profile
- **microsoft://calendars** - Available calendars
- **microsoft://teams** - Joined Teams information

## Example Usage

After setup, you can ask Claude:

- *"Send an email to john@company.com about the meeting tomorrow"*
- *"What's on my calendar this week?"*  
- *"Create a meeting for next Tuesday at 2 PM"*
- *"Search my emails for messages about the project deadline"*
- *"Show me my Microsoft 365 profile information"*

## Troubleshooting

### Web Connector Issues
- Make sure you're on Claude Pro/Team/Enterprise plan
- Try refreshing Claude Desktop after adding the connector
- Check that the server URL is correct: `https://your-worker-domain.com/sse`

### JSON Configuration Issues
- Ensure the `M365_ACCESS_TOKEN` environment variable is set
- Check that your OAuth access token hasn't expired
- Verify the path to the index.js file is correct
- Restart Claude Desktop after configuration changes

## How It Works

This connector bridges Claude Desktop with the Microsoft 365 MCP Server:

1. **Authentication**: Uses OAuth 2.1 with Microsoft Graph API
2. **Security**: All tokens encrypted and stored securely  
3. **Transport**: Server-Sent Events (SSE) for real-time communication
4. **Tools**: 8 Microsoft 365 tools for email, calendar, Teams, and contacts
5. **Resources**: 3 resource endpoints for profile and metadata

The server runs on Cloudflare Workers with enterprise-grade security and global edge performance.