/**
 * Microsoft 365 MCP Agent - Using Cloudflare OAuthProvider pattern
 */

import { McpAgent } from 'agents/mcp';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';
import { MicrosoftGraphClient } from './microsoft-graph';
import { Env } from './index';

// Props from Microsoft OAuth flow, encrypted & stored in the auth token
// and provided to the MCP Agent as this.props
export type Props = {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  accessToken: string;
  // Microsoft OAuth tokens from tokenExchangeCallback
  microsoftAccessToken?: string;
  microsoftTokenType?: string;
  microsoftScope?: string;
  microsoftRefreshToken?: string;
};

interface State {
  lastActivity?: number;
}

export class MicrosoftMCPAgent extends McpAgent<Env, State, Props> {
  // MCP Server configuration
  server = new McpServer({
    name: 'microsoft-365-mcp',
    version: '0.0.3',
  });

  // Initial state
  initialState: State = {
    lastActivity: Date.now(),
  };

  private graphClient: MicrosoftGraphClient;

  constructor(ctx: DurableObjectState, env: Env) {
    super(ctx, env);
    this.graphClient = new MicrosoftGraphClient(env);
  }

  private getAuthErrorResponse(): CallToolResult {
    return {
      content: [
        {
          type: 'text',
          text: 'Microsoft 365 authentication required. Please ensure you have completed the OAuth flow and have a valid access token.',
        },
      ],
      isError: true,
    };
  }

  async init() {
    // Email tools
    this.server.tool(
      'sendEmail',
      'Send an email via Outlook',
      {
        to: z.string().describe('Recipient email address'),
        subject: z.string().describe('Email subject'),
        body: z.string().describe('Email body content'),
        contentType: z.enum(['text', 'html']).default('html').describe('Content type'),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          await this.graphClient.sendEmail(accessToken, args);
          return { content: [{ type: 'text', text: `Email sent successfully to ${args.to}` }] };
        } catch (error: any) {
          return {
            content: [
              { type: 'text', text: `Failed to send email: ${error?.message || String(error)}` },
            ],
            isError: true,
          };
        }
      }
    );

    this.server.tool(
      'getEmails',
      'Get recent emails',
      {
        count: z.number().max(50).default(10).describe('Number of emails'),
        folder: z.string().default('inbox').describe('Mail folder'),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          const emails = await this.graphClient.getEmails(accessToken, {
            count: args.count,
            folder: args.folder,
          });
          return { content: [{ type: 'text', text: JSON.stringify(emails, null, 2) }] };
        } catch (error: any) {
          return {
            content: [
              { type: 'text', text: `Failed to get emails: ${error?.message || String(error)}` },
            ],
            isError: true,
          };
        }
      }
    );

    this.server.tool(
      'searchEmails',
      'Search emails',
      {
        query: z.string().describe('Search query'),
        count: z.number().max(50).default(10).describe('Number of results'),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          const results = await this.graphClient.searchEmails(accessToken, {
            query: args.query,
            count: args.count,
          });
          return { content: [{ type: 'text', text: JSON.stringify(results, null, 2) }] };
        } catch (error: any) {
          return {
            content: [
              { type: 'text', text: `Failed to search emails: ${error?.message || String(error)}` },
            ],
            isError: true,
          };
        }
      }
    );

    // Calendar tools
    this.server.tool(
      'getCalendarEvents',
      'Get calendar events',
      {
        days: z.number().max(30).default(7).describe('Days ahead'),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          const events = await this.graphClient.getCalendarEvents(accessToken, {
            days: args.days,
          });
          return { content: [{ type: 'text', text: JSON.stringify(events, null, 2) }] };
        } catch (error: any) {
          return {
            content: [
              {
                type: 'text',
                text: `Failed to get calendar events: ${error?.message || String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    this.server.tool(
      'createCalendarEvent',
      'Create calendar event',
      {
        subject: z.string().describe('Event title'),
        start: z.string().describe('Start time (ISO 8601)'),
        end: z.string().describe('End time (ISO 8601)'),
        attendees: z.array(z.string()).optional().describe('Attendee emails'),
        body: z.string().optional().describe('Event description'),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          const event = await this.graphClient.createCalendarEvent(accessToken, args);
          return { content: [{ type: 'text', text: `Event created: ${event.id}` }] };
        } catch (error: any) {
          return {
            content: [
              {
                type: 'text',
                text: `Failed to create calendar event: ${error?.message || String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    // Teams tools
    this.server.tool(
      'sendTeamsMessage',
      'Send Teams message',
      {
        teamId: z.string().describe('Team ID'),
        channelId: z.string().describe('Channel ID'),
        message: z.string().describe('Message content'),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          await this.graphClient.sendTeamsMessage(accessToken, args);
          return { content: [{ type: 'text', text: 'Teams message sent' }] };
        } catch (error: any) {
          return {
            content: [
              {
                type: 'text',
                text: `Failed to send Teams message: ${error?.message || String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    this.server.tool(
      'createTeamsMeeting',
      'Create Teams meeting',
      {
        subject: z.string().describe('Meeting title'),
        startTime: z.string().describe('Start time (ISO 8601)'),
        endTime: z.string().describe('End time (ISO 8601)'),
        attendees: z.array(z.string()).optional().describe('Attendee emails'),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          const meeting = await this.graphClient.createTeamsMeeting(accessToken, args);
          return { content: [{ type: 'text', text: `Meeting created: ${meeting.joinWebUrl}` }] };
        } catch (error: any) {
          return {
            content: [
              {
                type: 'text',
                text: `Failed to create Teams meeting: ${error?.message || String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    this.server.tool(
      'getContacts',
      'Get contacts',
      {
        count: z.number().max(100).default(50).describe('Number of contacts'),
        search: z.string().optional().describe('Search term'),
      },
      async (args): Promise<CallToolResult> => {
        const accessToken = this.props?.microsoftAccessToken;
        if (!accessToken) {
          return this.getAuthErrorResponse();
        }

        try {
          const contacts = await this.graphClient.getContacts(accessToken, {
            count: args.count,
            search: args.search,
          });
          return { content: [{ type: 'text', text: JSON.stringify(contacts, null, 2) }] };
        } catch (error: any) {
          return {
            content: [
              { type: 'text', text: `Failed to get contacts: ${error?.message || String(error)}` },
            ],
            isError: true,
          };
        }
      }
    );

    // Authentication tool for Claude Desktop
    this.server.tool(
      'authenticate',
      'Get authentication URL for Microsoft 365',
      {},
      async (): Promise<CallToolResult> => {
        return {
          content: [
            {
              type: 'text',
              text: 'Authentication is handled automatically by the OAuth provider. If you are seeing this message, please check your OAuth client configuration and ensure you have completed the authorization flow properly.',
            },
          ],
        };
      }
    );

    // Resources
    this.server.resource('profile', 'microsoft://profile', async () => {
      const accessToken = this.props?.microsoftAccessToken;
      if (!accessToken) {
        return {
          contents: [
            {
              uri: 'microsoft://profile',
              mimeType: 'application/json',
              text: JSON.stringify({ error: 'Authentication required', authenticated: false }, null, 2),
            },
          ],
        };
      }

      try {
        const profile = await this.graphClient.getUserProfile(accessToken);
        return {
          contents: [
            {
              uri: 'microsoft://profile',
              mimeType: 'application/json',
              text: JSON.stringify(profile, null, 2),
            },
          ],
        };
      } catch (error: any) {
        return {
          contents: [
            {
              uri: 'microsoft://profile',
              mimeType: 'application/json',
              text: JSON.stringify({ error: error.message || 'Failed to fetch profile', authenticated: true }, null, 2),
            },
          ],
        };
      }
    });

    this.server.resource('calendars', 'microsoft://calendars', async () => {
      const accessToken = this.props?.microsoftAccessToken;
      if (!accessToken) {
        return {
          contents: [
            {
              uri: 'microsoft://calendars',
              mimeType: 'application/json',
              text: JSON.stringify({ error: 'Authentication required', authenticated: false, calendars: [] }, null, 2),
            },
          ],
        };
      }

      try {
        const calendars = await this.graphClient.getCalendars(accessToken);
        return {
          contents: [
            {
              uri: 'microsoft://calendars',
              mimeType: 'application/json',
              text: JSON.stringify(calendars, null, 2),
            },
          ],
        };
      } catch (error: any) {
        return {
          contents: [
            {
              uri: 'microsoft://calendars',
              mimeType: 'application/json',
              text: JSON.stringify({ error: error.message || 'Failed to fetch calendars', authenticated: true, calendars: [] }, null, 2),
            },
          ],
        };
      }
    });

    this.server.resource('teams', 'microsoft://teams', async () => {
      const accessToken = this.props?.microsoftAccessToken;
      if (!accessToken) {
        return {
          contents: [
            {
              uri: 'microsoft://teams',
              mimeType: 'application/json',
              text: JSON.stringify({ error: 'Authentication required', authenticated: false, teams: [] }, null, 2),
            },
          ],
        };
      }

      try {
        const teams = await this.graphClient.getTeams(accessToken);
        return {
          contents: [
            {
              uri: 'microsoft://teams',
              mimeType: 'application/json',
              text: JSON.stringify(teams, null, 2),
            },
          ],
        };
      } catch (error: any) {
        return {
          contents: [
            {
              uri: 'microsoft://teams',
              mimeType: 'application/json',
              text: JSON.stringify({ error: error.message || 'Failed to fetch teams', authenticated: true, teams: [] }, null, 2),
            },
          ],
        };
      }
    });
  }

  // Handle conditional authentication for Claude Desktop + mcp-remote
  async fetch(request: Request): Promise<Response> {
    const mcpMode = request.headers.get('X-MCP-Mode');
    const webSocketSession = request.headers.get('X-WebSocket-Session');
    
    // Handle WebSocket upgrade requests
    if (mcpMode === 'websocket' && webSocketSession) {
      console.log(`MCP Agent handling WebSocket upgrade for session: ${webSocketSession}`);
      
      // WebSocket upgrades require proper upgrade response
      const upgradeHeader = request.headers.get('Upgrade');
      const webSocketKey = request.headers.get('Sec-WebSocket-Key');
      
      if (upgradeHeader?.toLowerCase() === 'websocket' && webSocketKey) {
        try {
          // Implement proper WebSocket upgrade using Cloudflare's WebSocket API
          console.log('Processing WebSocket upgrade for Claude Desktop with Cloudflare WebSocket API');
          
          // Create a WebSocket pair
          const webSocketPair = new WebSocketPair();
          const [client, server] = Object.values(webSocketPair);
          
          // Accept the WebSocket connection
          server.accept();
          
          // Set up basic MCP WebSocket handling
          server.addEventListener('message', (event) => {
            console.log('WebSocket message received:', event.data);
            // Echo back for now - full MCP protocol implementation would go here
            server.send(JSON.stringify({
              jsonrpc: '2.0',
              error: { code: -32601, message: 'WebSocket MCP protocol not fully implemented' }
            }));
          });
          
          server.addEventListener('close', () => {
            console.log('WebSocket connection closed');
          });
          
          // Return the client WebSocket to the browser
          return new Response(null, {
            status: 101,
            webSocket: client,
          });
        } catch (error) {
          console.error('WebSocket upgrade processing error:', error);
          return new Response(`WebSocket processing error: ${error}`, { status: 500 });
        }
      }
    }
    
    // Handle unprotected handshake and other methods for Claude Desktop
    if (mcpMode === 'handshake' || mcpMode === 'other') {
      console.log(`MCP Agent handling ${mcpMode} mode request without authentication`);
      
      // Temporarily clear props to allow unauthenticated access
      const originalProps = this.props;
      (this as any).props = null;
      
      try {
        const response = await super.fetch(request);
        return response;
      } finally {
        // Restore original props
        (this as any).props = originalProps;
      }
    }
    
    // For OAuth-protected requests, use normal authentication
    console.log('MCP Agent handling authenticated request');
    return super.fetch(request);
  }

  // Generate WebSocket accept key per RFC 6455
  private async generateWebSocketAccept(webSocketKey: string): Promise<string> {
    const webSocketMagicString = '258EAFA5-E914-47DA-95CA-C5AB0DC85B11';
    const concatenated = webSocketKey + webSocketMagicString;
    
    // Hash with SHA-1 and encode as base64
    const encoder = new TextEncoder();
    const data = encoder.encode(concatenated);
    const hashBuffer = await crypto.subtle.digest('SHA-1', data);
    const hashArray = new Uint8Array(hashBuffer);
    
    // Convert to base64
    let binary = '';
    for (let i = 0; i < hashArray.length; i++) {
      binary += String.fromCharCode(hashArray[i]);
    }
    return btoa(binary);
  }

  // State update handler
  onStateUpdate(state: State) {
    console.log('State updated:', { lastActivity: state.lastActivity });
  }
}
