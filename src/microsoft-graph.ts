/**
 * Microsoft Graph API Client
 * Handles all interactions with Microsoft Graph API for Office 365 and Teams
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

export class MicrosoftGraphClient {
  private env: Env;
  private baseUrl: string;

  constructor(env: Env) {
    this.env = env;
    this.baseUrl = `https://graph.microsoft.com/${env.GRAPH_API_VERSION}`;
  }

  // Email Operations
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

  // Calendar Operations
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

  // Teams Operations
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

  // Contact Operations
  async getContacts(accessToken: string, params: ContactsParams): Promise<any> {
    const count = Math.min(params.count || 50, 100);
    let url = `${this.baseUrl}/me/contacts?$top=${count}&$select=id,displayName,emailAddresses,businessPhones,mobilePhone`;

    if (params.search) {
      url += `&$filter=startswith(displayName,'${encodeURIComponent(params.search)}') or startswith(givenName,'${encodeURIComponent(params.search)}')`;
    }

    const response = await this.makeGraphRequest(accessToken, url, 'GET');
    return response.value || [];
  }

  // User Profile
  async getUserProfile(accessToken: string): Promise<any> {
    const url = `${this.baseUrl}/me?$select=id,displayName,mail,userPrincipalName,jobTitle,department,companyName`;

    const response = await this.makeGraphRequest(accessToken, url, 'GET');
    return response;
  }

  // Generic Graph API request handler
  private async makeGraphRequest(
    accessToken: string,
    url: string,
    method: string = 'GET',
    body?: any
  ): Promise<any> {
    try {
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

        throw new Error(
          `Microsoft Graph API error: ${response.status} - ${errorData.error?.message || 'Unknown error'}`
        );
      }

      // Handle 204 No Content responses
      if (response.status === 204) {
        return {};
      }

      const responseData = await response.json();
      return responseData;
    } catch (error) {
      console.error('Graph API request failed:', error);

      if (error instanceof Error) {
        throw error;
      }

      throw new Error('Unknown error occurred during Graph API request');
    }
  }

  // Helper method to handle paginated responses
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
