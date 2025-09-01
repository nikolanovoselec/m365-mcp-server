import { Props } from './microsoft-mcp-agent';

export { Props };

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

export async function fetchUpstreamAuthToken(params: {
  client_id: string;
  client_secret: string;
  code: string | null;
  redirect_uri: string;
  upstream_url: string;
}): Promise<[string, Response | null]> {
  if (!params.code) {
    return ['', new Response('Missing authorization code', { status: 400 })];
  }

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

  return [data.access_token, null];
}
