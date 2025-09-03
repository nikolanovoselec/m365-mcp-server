/**
 * Utility Functions for Microsoft Identity Platform OAuth Integration
 * 
 * SCOPE:
 * - OAuth 2.1 URL construction for Microsoft authorization endpoints
 * - Token exchange implementation for authorization code flow
 * - Error handling for Microsoft Identity Platform responses
 * 
 * ARCHITECTURE:
 * These utilities bridge the OAuth Provider with Microsoft's OAuth implementation,
 * handling the specifics of Microsoft's OAuth 2.1 + PKCE requirements.
 */

import { Props } from './microsoft-mcp-agent';

export { Props };

/**
 * Build OAuth 2.1 authorization URL for Microsoft Identity Platform
 * 
 * Constructs a properly formatted authorization URL following Microsoft's
 * OAuth 2.1 implementation with all required parameters.
 * 
 * PARAMETERS EXPLAINED:
 * - response_type: 'code' for authorization code flow
 * - response_mode: 'query' returns code as query parameter
 * - state: Preserves request context through OAuth flow
 * - scope: Space-separated list of Microsoft Graph permissions
 * 
 * @param params - Authorization URL parameters
 * @param params.client_id - Microsoft application (client) ID from Azure AD
 * @param params.redirect_uri - Callback URL registered in Azure AD
 * @param params.scope - Space-separated Microsoft Graph API scopes
 * @param params.state - Opaque value to maintain state between request and callback
 * @param params.upstream_url - Microsoft authorization endpoint URL
 * @returns Fully constructed authorization URL for redirect
 */
export function getUpstreamAuthorizeUrl(params: {
  client_id: string;
  redirect_uri: string;
  scope: string;
  state: string;
  upstream_url: string;
}) {
  const url = new URL(params.upstream_url);
  url.searchParams.set('client_id', params.client_id);
  url.searchParams.set('response_type', 'code');
  url.searchParams.set('redirect_uri', params.redirect_uri);
  url.searchParams.set('scope', params.scope);
  url.searchParams.set('state', params.state);
  url.searchParams.set('response_mode', 'query');

  return url.toString();
}

/**
 * Exchange authorization code for access tokens with Microsoft Identity Platform
 * 
 * Implements OAuth 2.1 token exchange endpoint call with proper error handling
 * and response validation. Used by OAuth Provider tokenExchangeCallback.
 * 
 * ERROR HANDLING:
 * Returns tuple pattern for explicit error handling without exceptions.
 * Caller must check for error response before using access token.
 * 
 * SECURITY:
 * - Uses application/x-www-form-urlencoded to prevent JSON injection
 * - Validates response before extracting tokens
 * - Returns detailed error messages for debugging
 * 
 * @param params - Token exchange parameters
 * @param params.client_id - Microsoft application (client) ID
 * @param params.client_secret - Microsoft application secret (keep secure!)
 * @param params.code - Authorization code from Microsoft callback
 * @param params.redirect_uri - Must match original authorization request
 * @param params.upstream_url - Microsoft token endpoint URL
 * @returns Tuple of [access_token, error_response] - check error before using token
 */
export async function fetchUpstreamAuthToken(params: {
  client_id: string;
  client_secret: string;
  code: string | null;
  redirect_uri: string;
  upstream_url: string;
}): Promise<[string, Response | null]> {
  /**
   * Validate authorization code presence
   * Missing code indicates failed authorization or invalid callback
   */
  if (!params.code) {
    return ['', new Response('Missing authorization code', { status: 400 })];
  }

  /**
   * Construct token exchange request body
   * 
   * Microsoft requires application/x-www-form-urlencoded format.
   * grant_type must be 'authorization_code' for initial token request.
   */
  const body = new URLSearchParams({
    client_id: params.client_id,
    client_secret: params.client_secret,
    code: params.code,
    grant_type: 'authorization_code',
    redirect_uri: params.redirect_uri,
  });

  const response = await fetch(params.upstream_url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      Accept: 'application/json',
    },
    body: body.toString(),
  });

  const data = (await response.json()) as any;

  /**
   * Handle token exchange errors
   * 
   * Common errors:
   * - invalid_grant: Authorization code expired or already used
   * - invalid_client: Client ID/secret mismatch
   * - invalid_request: Missing or malformed parameters
   */
  if (!response.ok) {
    const errorMsg = data.error_description || data.error || 'Token exchange failed';
    return [
      '',
      new Response(JSON.stringify({ error: errorMsg }), {
        status: response.status,
        headers: { 'Content-Type': 'application/json' },
      }),
    ];
  }

  /** Successfully obtained access token - return for storage in OAuth Provider props */
  return [data.access_token, null];
}
