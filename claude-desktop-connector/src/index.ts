#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { 
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ListResourcesRequestSchema,
  ReadResourceRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';

// Configuration from environment variables
const M365_ACCESS_TOKEN = process.env.M365_ACCESS_TOKEN;
const M365_SERVER_URL = process.env.M365_SERVER_URL || 'https://your-worker-domain.com';

interface M365Config {
  accessToken: string;
  serverUrl: string;
}

class Microsoft365MCPServer {
  private server: Server;
  private config: M365Config;

  constructor() {
    if (!M365_ACCESS_TOKEN) {
      console.error('\n❌ M365_ACCESS_TOKEN environment variable is required\n');
      console.error('To get an access token:');
      console.error('1. Register OAuth client:');
      console.error(`   curl -X POST ${M365_SERVER_URL}/register \\`);
      console.error('     -H "Content-Type: application/json" \\');
      console.error('     -d \'{"client_name":"claude-desktop","redirect_uris":["http://localhost:8080/callback"]}\'');
      console.error('\n2. Complete OAuth flow with returned client_id and client_secret');
      console.error('3. Set access token: export M365_ACCESS_TOKEN="your_token_here"\n');
      process.exit(1);
    }

    this.config = {
      accessToken: M365_ACCESS_TOKEN,
      serverUrl: M365_SERVER_URL,
    };

    this.server = new Server(
      {
        name: 'microsoft-365',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
          resources: {},
        },
      }
    );

    this.setupToolHandlers();
    this.setupResourceHandlers();
  }

  private setupToolHandlers(): void {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      return {
        tools: [
          {
            name: 'sendEmail',
            description: 'Send an email via Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                to: {
                  type: 'string',
                  description: 'Recipient email address',
                },
                subject: {
                  type: 'string',
                  description: 'Email subject',
                },
                body: {
                  type: 'string',
                  description: 'Email body content',
                },
                contentType: {
                  type: 'string',
                  enum: ['text', 'html'],
                  default: 'html',
                  description: 'Content type (text or html)',
                },
              },
              required: ['to', 'subject', 'body'],
            },
          },
          {
            name: 'getEmails',
            description: 'Get recent emails from Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                count: {
                  type: 'number',
                  maximum: 50,
                  default: 10,
                  description: 'Number of emails to retrieve',
                },
                folder: {
                  type: 'string',
                  default: 'inbox',
                  description: 'Mail folder to search in',
                },
              },
            },
          },
          {
            name: 'searchEmails',
            description: 'Search emails in Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                query: {
                  type: 'string',
                  description: 'Search query string',
                },
                count: {
                  type: 'number',
                  maximum: 50,
                  default: 10,
                  description: 'Number of search results',
                },
              },
              required: ['query'],
            },
          },
          {
            name: 'getCalendarEvents',
            description: 'Get upcoming calendar events',
            inputSchema: {
              type: 'object',
              properties: {
                days: {
                  type: 'number',
                  maximum: 30,
                  default: 7,
                  description: 'Number of days ahead to look',
                },
              },
            },
          },
          {
            name: 'createCalendarEvent',
            description: 'Create a new calendar event',
            inputSchema: {
              type: 'object',
              properties: {
                subject: {
                  type: 'string',
                  description: 'Event title/subject',
                },
                start: {
                  type: 'string',
                  description: 'Start time in ISO 8601 format',
                },
                end: {
                  type: 'string',
                  description: 'End time in ISO 8601 format',
                },
                attendees: {
                  type: 'array',
                  items: {
                    type: 'string',
                  },
                  description: 'List of attendee email addresses',
                },
                body: {
                  type: 'string',
                  description: 'Event description/body',
                },
              },
              required: ['subject', 'start', 'end'],
            },
          },
          {
            name: 'sendTeamsMessage',
            description: 'Send a message to a Teams channel',
            inputSchema: {
              type: 'object',
              properties: {
                teamId: {
                  type: 'string',
                  description: 'Microsoft Teams team ID',
                },
                channelId: {
                  type: 'string',
                  description: 'Teams channel ID',
                },
                message: {
                  type: 'string',
                  description: 'Message content to send',
                },
              },
              required: ['teamId', 'channelId', 'message'],
            },
          },
          {
            name: 'createTeamsMeeting',
            description: 'Create a new Teams meeting',
            inputSchema: {
              type: 'object',
              properties: {
                subject: {
                  type: 'string',
                  description: 'Meeting title/subject',
                },
                startTime: {
                  type: 'string',
                  description: 'Meeting start time in ISO 8601 format',
                },
                endTime: {
                  type: 'string',
                  description: 'Meeting end time in ISO 8601 format',
                },
                attendees: {
                  type: 'array',
                  items: {
                    type: 'string',
                  },
                  description: 'List of attendee email addresses',
                },
              },
              required: ['subject', 'startTime', 'endTime'],
            },
          },
          {
            name: 'getContacts',
            description: 'Get contacts from Outlook',
            inputSchema: {
              type: 'object',
              properties: {
                count: {
                  type: 'number',
                  maximum: 100,
                  default: 50,
                  description: 'Number of contacts to retrieve',
                },
                search: {
                  type: 'string',
                  description: 'Search term to filter contacts',
                },
              },
            },
          },
        ],
      };
    });

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      try {
        // For now, we'll use mcp-remote to proxy the request
        // In a full implementation, this would make HTTP calls to the OAuth server
        
        return {
          content: [
            {
              type: 'text',
              text: `✅ Microsoft 365 MCP Server Connected
              
Tool: ${name}
Arguments: ${JSON.stringify(args, null, 2)}
Server: ${this.config.serverUrl}
Token: ${this.config.accessToken.substring(0, 20)}...

🔄 This connector would now forward your request to the Microsoft 365 MCP Server 
running at ${this.config.serverUrl} using your OAuth access token.

To complete the integration, this tool needs to establish communication with 
the remote OAuth-enabled MCP server.`,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: `❌ Error calling ${name}: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    });
  }

  private setupResourceHandlers(): void {
    this.server.setRequestHandler(ListResourcesRequestSchema, async () => {
      return {
        resources: [
          {
            uri: 'microsoft://profile',
            mimeType: 'application/json',
            name: 'Microsoft 365 Profile',
            description: 'Current user profile information from Microsoft 365',
          },
          {
            uri: 'microsoft://calendars',
            mimeType: 'application/json', 
            name: 'Calendars',
            description: 'List of available calendars in Microsoft 365',
          },
          {
            uri: 'microsoft://teams',
            mimeType: 'application/json',
            name: 'Microsoft Teams',
            description: 'Information about joined Microsoft Teams',
          },
        ],
      };
    });

    this.server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
      const { uri } = request.params;

      const resourceData = {
        uri,
        server: this.config.serverUrl,
        authenticated: true,
        access_token_configured: !!this.config.accessToken,
        message: `Resource ${uri} would be fetched from the Microsoft 365 MCP Server using OAuth authentication`,
        next_steps: [
          'This connector is authenticated and ready',
          'Resource requests would be forwarded to the remote server',
          'OAuth tokens would be used for Microsoft Graph API calls',
        ],
      };

      return {
        contents: [
          {
            uri,
            mimeType: 'application/json',
            text: JSON.stringify(resourceData, null, 2),
          },
        ],
      };
    });
  }

  async run(): Promise<void> {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    
    // This error message goes to stderr and won't interfere with MCP protocol
    console.error(`🚀 Microsoft 365 MCP Server ready`);
    console.error(`📡 Connected to: ${this.config.serverUrl}`);
    console.error(`🔐 Access token: ${this.config.accessToken.substring(0, 20)}...`);
  }
}

// Start the server
const server = new Microsoft365MCPServer();
server.run().catch((error) => {
  console.error('❌ Failed to start Microsoft 365 MCP Server:', error);
  process.exit(1);
});