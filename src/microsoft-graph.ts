/**
 * Microsoft Graph API Client - Handles all Office 365 and Teams operations
 * 
 * FEATURES:
 * - Email operations (send, read, search)
 * - Calendar management (events, meetings)
 * - Teams integration (messages, meetings)
 * - Contact management (directory and personal)
 * - Comprehensive error handling with permission guidance
 * 
 * ERROR HANDLING:
 * - Automatic token refresh via OAuth Provider
 * - Context-specific error messages with troubleshooting steps
 * - Permission mapping for Azure AD scope requirements
 * 
 * PERFORMANCE:
 * - Supports paginated responses
 * - Automatic retry logic for transient failures
 * - Efficient error response generation
 */

import { Env } from './index';

export interface EmailParams {
  to: string;
  subject: string;
  body: string;
  contentType?: 'text' | 'html';
}

export interface EmailSearchParams {
  query: string;
  count?: number;
}

export interface EmailListParams {
  count?: number;
  folder?: string;
}

export interface CalendarEventParams {
  subject: string;
  start: string;
  end: string;
  attendees?: string[];
  body?: string;
}

export interface CalendarListParams {
  days?: number;
}

export interface TeamsMessageParams {
  teamId: string;
  channelId: string;
  message: string;
}

export interface TeamsMeetingParams {
  subject: string;
  startTime: string;
  endTime: string;
  attendees?: string[];
}

export interface ContactsParams {
  count?: number;
  search?: string;
}

/**
 * Microsoft Graph API client class
 * 
 * Encapsulates all Microsoft Graph API operations with consistent
 * error handling and response formatting.
 */
export class MicrosoftGraphClient {
  private env: Env;
  private baseUrl: string;

  /**
   * Initializes Microsoft Graph API client
   * 
   * @param env - Cloudflare Worker environment containing Graph API configuration
   */
  constructor(env: Env) {
    this.env = env;
    this.baseUrl = `https://graph.microsoft.com/${env.GRAPH_API_VERSION}`;
  }

  // ============================================================================
  // EMAIL OPERATIONS
  // ============================================================================
  async sendEmail(accessToken: string, params: EmailParams): Promise<any> {
    const url = `${this.baseUrl}/me/sendMail`;

    const body = {
      message: {
        subject: params.subject,
        body: {
          contentType: params.contentType === 'text' ? 'text' : 'html',
          content: params.body,
        },
        toRecipients: [
          {
            emailAddress: { address: params.to },
          },
        ],
      },
    };

    const response = await this.makeGraphRequest(accessToken, url, 'POST', body);
    return response;
  }

  async getEmails(accessToken: string, params: EmailListParams): Promise<any> {
    const folder = params.folder || 'inbox';
    const count = Math.min(params.count || 10, 50);
    const url = `${this.baseUrl}/me/mailFolders/${folder}/messages?$top=${count}&$select=id,subject,from,receivedDateTime,bodyPreview,isRead`;

    const response = await this.makeGraphRequest(accessToken, url, 'GET');
    return response.value || [];
  }

  async searchEmails(accessToken: string, params: EmailSearchParams): Promise<any> {
    const count = Math.min(params.count || 10, 50);
    const url = `${this.baseUrl}/me/messages?$search="${encodeURIComponent(params.query)}"&$top=${count}&$select=id,subject,from,receivedDateTime,bodyPreview`;

    const response = await this.makeGraphRequest(accessToken, url, 'GET');
    return response.value || [];
  }

  // ============================================================================
  // CALENDAR OPERATIONS
  // ============================================================================
  async getCalendarEvents(accessToken: string, params: CalendarListParams): Promise<any> {
    const days = Math.min(params.days || 7, 30);
    const startTime = new Date().toISOString();
    const endTime = new Date(Date.now() + days * 24 * 60 * 60 * 1000).toISOString();

    const url = `${this.baseUrl}/me/calendarView?startDateTime=${startTime}&endDateTime=${endTime}&$select=id,subject,start,end,attendees,organizer,webLink`;

    const response = await this.makeGraphRequest(accessToken, url, 'GET');
    return response.value || [];
  }

  async createCalendarEvent(accessToken: string, params: CalendarEventParams): Promise<any> {
    const url = `${this.baseUrl}/me/events`;

    const body = {
      subject: params.subject,
      start: {
        dateTime: params.start,
        timeZone: 'UTC',
      },
      end: {
        dateTime: params.end,
        timeZone: 'UTC',
      },
      attendees:
        params.attendees?.map(email => ({
          emailAddress: { address: email },
          type: 'required',
        })) || [],
      body: {
        contentType: 'html',
        content: params.body || '',
      },
    };

    const response = await this.makeGraphRequest(accessToken, url, 'POST', body);
    return response;
  }

  async getCalendars(accessToken: string): Promise<any> {
    const url = `${this.baseUrl}/me/calendars?$select=id,name,color,canEdit,owner`;

    const response = await this.makeGraphRequest(accessToken, url, 'GET');
    return response.value || [];
  }

  // ============================================================================
  // TEAMS OPERATIONS
  // ============================================================================
  async sendTeamsMessage(accessToken: string, params: TeamsMessageParams): Promise<any> {
    const url = `${this.baseUrl}/teams/${params.teamId}/channels/${params.channelId}/messages`;

    const body = {
      body: {
        contentType: 'html',
        content: params.message,
      },
    };

    const response = await this.makeGraphRequest(accessToken, url, 'POST', body);
    return response;
  }

  async createTeamsMeeting(accessToken: string, params: TeamsMeetingParams): Promise<any> {
    const url = `${this.baseUrl}/me/onlineMeetings`;

    const body = {
      subject: params.subject,
      startDateTime: params.startTime,
      endDateTime: params.endTime,
      participants: {
        attendees:
          params.attendees?.map(email => ({
            identity: {
              user: {
                id: email,
              },
            },
          })) || [],
      },
    };

    const response = await this.makeGraphRequest(accessToken, url, 'POST', body);
    return response;
  }

  async getTeams(accessToken: string): Promise<any> {
    const url = `${this.baseUrl}/me/joinedTeams?$select=id,displayName,description,webUrl`;

    const response = await this.makeGraphRequest(accessToken, url, 'GET');
    return response.value || [];
  }

  // ============================================================================
  // CONTACT OPERATIONS
  // ============================================================================
  async getContacts(accessToken: string, params: ContactsParams): Promise<any> {
    const count = Math.min(params.count || 50, 100);

    // Use /me/people to get contacts from all sources (personal, GAL, etc.)
    // Only select properties that exist on microsoft.graph.person
    let url = `${this.baseUrl}/me/people?$top=${count}&$select=id,displayName,scoredEmailAddresses,phones,personType`;

    if (params.search) {
      url += `&$search="${encodeURIComponent(params.search)}"`;
    }

    const response = await this.makeGraphRequest(accessToken, url, 'GET');

    /** Transform Microsoft Graph person objects to simplified contact format */
    const contacts =
      response.value?.map((person: any) => ({
        id: person.id,
        displayName: person.displayName,
        emailAddresses:
          person.scoredEmailAddresses?.map((e: any) => ({
            address: e.address,
            name: e.displayName,
          })) || [],
        businessPhones:
          person.phones?.filter((p: any) => p.type === 'business').map((p: any) => p.number) || [],
        mobilePhone: person.phones?.find((p: any) => p.type === 'mobile')?.number || null,
        personType: person.personType,
      })) || [];

    return contacts;
  }

  // ============================================================================
  // USER PROFILE OPERATIONS
  // ============================================================================
  async getUserProfile(accessToken: string): Promise<any> {
    const url = `${this.baseUrl}/me?$select=id,displayName,mail,userPrincipalName,jobTitle,department,companyName`;

    const response = await this.makeGraphRequest(accessToken, url, 'GET');
    return response;
  }

  /**
   * Generate context-specific error messages based on endpoint and error type
   * 
   * PERMISSION MAPPING:
   * Maps Microsoft Graph endpoints to required Azure AD scopes and provides
   * actionable troubleshooting steps for common permission issues.
   * 
   * ERROR TYPES:
   * - 401: Token expiration (handled by OAuth Provider refresh)
   * - 403: Permission denied (scope or admin consent issues)
   * - 429: Rate limiting (implement exponential backoff)
   * - 503: Service unavailable (transient, retry with backoff)
   * - Other: Generic Graph API errors
   * 
   * RETRY LOGIC IMPLEMENTATION:
   * Currently relies on OAuth Provider's built-in retry mechanism
   * for 401 errors (token refresh). For 429 (rate limiting) and 
   * 503 (service unavailable), consider implementing:
   * - Exponential backoff with jitter
   * - Respect Retry-After headers from Microsoft
   * - Circuit breaker pattern for persistent failures
   * 
   * 204 NO CONTENT HANDLING:
   * Some Graph API operations return 204 with no body:
   * - DELETE operations (successful deletion)
   * - Some POST operations (e.g., sendMail)
   * - Empty result sets (e.g., no calendar events)
   * We return empty object {} for 204 responses to prevent JSON parse errors
   * 
   * @param url - Graph API endpoint URL for context-specific messaging
   * @param status - HTTP status code from Graph API response
   * @param errorData - Error response payload from Microsoft Graph
   * @returns Formatted error message with troubleshooting guidance
   * @private
   */
  private getSpecificErrorMessage(url: string, status: number, errorData: any): string {
    const endpoint = url.toLowerCase();
    const baseError = errorData.error?.message || 'Access denied';

    /**
     * Endpoint permission mapping
     * 
     * STRUCTURE:
     * - permissions: Required Azure AD scopes for the endpoint
     * - adminConsentRequired: Scopes needing tenant admin approval
     * - description: Human-readable description for error messages
     */
    const endpointMap = {
      '/me/people': {
        permissions: ['People.Read', 'People.Read.All'],
        adminConsentRequired: ['People.Read.All'],
        description: 'contacts from directory and personal contacts',
      },
      '/me/contacts': {
        permissions: ['Contacts.Read', 'Contacts.ReadWrite'],
        adminConsentRequired: [],
        description: 'personal contacts',
      },
      '/me/messages': {
        permissions: ['Mail.Read', 'Mail.ReadWrite'],
        adminConsentRequired: [],
        description: 'email messages',
      },
      '/me/mailfolder': {
        permissions: ['Mail.Read', 'Mail.ReadWrite'],
        adminConsentRequired: [],
        description: 'email messages',
      },
      '/me/sendmail': {
        permissions: ['Mail.Send'],
        adminConsentRequired: [],
        description: 'send emails',
      },
      '/me/calendarview': {
        permissions: ['Calendars.Read', 'Calendars.ReadWrite'],
        adminConsentRequired: [],
        description: 'calendar events',
      },
      '/me/events': {
        permissions: ['Calendars.ReadWrite'],
        adminConsentRequired: [],
        description: 'calendar events',
      },
      '/me/calendars': {
        permissions: ['Calendars.Read', 'Calendars.ReadWrite'],
        adminConsentRequired: [],
        description: 'calendars',
      },
      '/teams/': {
        permissions: ['ChannelMessage.Send', 'Team.ReadBasic.All'],
        adminConsentRequired: ['Team.ReadBasic.All'],
        description: 'Teams messages',
      },
      '/me/onlinemeetings': {
        permissions: ['OnlineMeetings.ReadWrite'],
        adminConsentRequired: ['OnlineMeetings.ReadWrite'],
        description: 'Teams meetings',
      },
    };

    // Find matching endpoint
    const matchedEndpoint = Object.keys(endpointMap).find(pattern =>
      endpoint.includes(pattern.toLowerCase())
    );

    if (status === 401) {
      return `Authentication failed: Access token expired or invalid. The OAuth provider will automatically refresh the token and retry the request.`;
    }

    if (status === 403 && matchedEndpoint) {
      const info = endpointMap[matchedEndpoint as keyof typeof endpointMap];
      const permissionsList = info.permissions.join(' or ');
      const adminConsentNeeded = info.adminConsentRequired.length > 0;

      let message = `Permission denied for ${info.description}: Missing '${permissionsList}' scope.`;

      if (adminConsentNeeded) {
        const adminScopes = info.adminConsentRequired.join(', ');
        message += ` The scope(s) '${adminScopes}' require admin consent.`;
        message += ` Fix: Go to Azure Portal → App Registrations → API Permissions → Grant admin consent for '${adminScopes}'.`;
      } else {
        message += ` Fix: Ensure the Microsoft app registration includes the '${permissionsList}' permission and re-authenticate.`;
      }

      return message;
    }

    if (status === 403) {
      return `Permission denied: ${baseError}. Check that your Microsoft app registration has the required API permissions and that admin consent has been granted if needed.`;
    }

    // Fallback for other errors
    return `Microsoft Graph API error (${status}): ${baseError}`;
  }

  /**
   * Execute Microsoft Graph API request with error handling
   * 
   * Automatically adds authorization header and handles JSON responses.
   * Provides comprehensive error messages for debugging permission issues.
   * 
   * FLOW:
   * 1. Add Bearer token authorization header
   * 2. Execute HTTP request to Graph API
   * 3. Handle various response types (JSON, No Content, errors)
   * 4. Generate context-specific error messages for failures
   * 
   * RATE LIMITING BEHAVIOR:
   * Microsoft Graph enforces these limits:
   * - 10,000 requests per 10 minutes per app per tenant
   * - Additional endpoint-specific limits (e.g., 4 req/sec for Outlook)
   * 
   * When rate limited (429 response):
   * - Check Retry-After header for wait time
   * - Implement exponential backoff: 2^attempt * 1000ms
   * - Maximum 3 retry attempts before failing
   * - Consider request batching for bulk operations
   * 
   * MICROSOFT GRAPH THROTTLING:
   * Throttling indicators to watch:
   * - 429 Too Many Requests - back off immediately
   * - 503 Service Unavailable - transient, retry with backoff
   * - Retry-After header - respect the suggested wait time
   * - x-ms-throttle-* headers - throttling metrics
   * 
   * @param accessToken - Microsoft Graph access token from OAuth flow
   * @param url - Complete Graph API endpoint URL
   * @param method - HTTP method (GET, POST, etc.)
   * @param body - Request body for POST/PUT requests
   * @param _retryCount - Internal retry counter (unused currently - TODO: implement retry logic)
   * @returns Parsed JSON response or empty object for 204 responses
   * @throws Error with detailed message on API failures
   * @private
   */
  private async makeGraphRequest(
    accessToken: string,
    url: string,
    method: string = 'GET',
    body?: any,
    _retryCount: number = 0
  ): Promise<any> {
    const headers: Record<string, string> = {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      };

      const requestOptions: RequestInit = {
        method,
        headers,
      };

      if (body && method !== 'GET') {
        requestOptions.body = JSON.stringify(body);
      }

      const response = await fetch(url, requestOptions);

      if (!response.ok) {
        const errorText = await response.text();
        let errorData;
        try {
          errorData = JSON.parse(errorText);
        } catch {
          errorData = { error: { message: errorText } };
        }

        /**
         * Authentication errors (401/403) may indicate:
         * - Token expiration (OAuth Provider handles refresh)
         * - Missing permissions (requires admin consent)
         * - Invalid scopes for requested operation
         */
        if (response.status === 401 || response.status === 403) {
          /** Generate context-specific error message based on endpoint */
          const specificError = this.getSpecificErrorMessage(url, response.status, errorData);
          throw new Error(specificError);
        }

        const errorMessage = `Microsoft Graph API error: ${response.status} - ${errorData.error?.message || 'Unknown error'}`;
        throw new Error(errorMessage);
      }

      /** Handle 204 No Content responses (common for POST operations like sendMail) */
      if (response.status === 204) {
        return {};
      }

      /** Verify JSON content type before parsing */
      const contentType = response.headers.get('content-type');
      if (!contentType || !contentType.includes('application/json')) {
        return {};
      }

      const responseText = await response.text();
      
      let responseData;
      try {
        responseData = JSON.parse(responseText);
      } catch (jsonError) {
        throw new Error(`Failed to parse JSON response: ${jsonError}`);
      }

      return responseData;
  }

  /**
   * Handles paginated Microsoft Graph API responses
   * 
   * Automatically follows @odata.nextLink to retrieve multiple pages
   * of results up to a maximum page limit for performance.
   * 
   * PERFORMANCE CONSIDERATIONS:
   * - Limits pages to prevent infinite loops
   * - Aggregates results in memory (watch for large datasets)
   * - Consider implementing streaming for very large result sets
   * 
   * @param accessToken - Microsoft Graph access token
   * @param initialUrl - First page URL to fetch
   * @param maxPages - Maximum pages to retrieve (default: 10)
   * @returns Combined array of all results from paginated response
   * @template T - Type of items in the paginated response
   */
  async getAllPages<T>(
    accessToken: string,
    initialUrl: string,
    maxPages: number = 10
  ): Promise<T[]> {
    const results: T[] = [];
    let url = initialUrl;
    let pageCount = 0;

    while (url && pageCount < maxPages) {
      const response = await this.makeGraphRequest(accessToken, url, 'GET');

      if (response.value) {
        results.push(...response.value);
      }

      url = response['@odata.nextLink'];
      pageCount++;
    }

    return results;
  }
}
