async function getGatewayAccessToken() {
  const scopeBase = process.env.GATEWAY_API_APP_ID_URI || '';
  if (!scopeBase) return null;

  const tenantId = process.env.AAD_TENANT_ID || process.env.MicrosoftAppTenantId;
  const clientId = process.env.BOT_CALLER_CLIENT_ID || process.env.MicrosoftAppId;
  const clientSecret = process.env.BOT_CALLER_CLIENT_SECRET || process.env.MicrosoftAppPassword;

  if (!tenantId || !clientId || !clientSecret) return null;

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: clientId,
    client_secret: clientSecret,
    scope: `${scopeBase}/.default`
  });

  try {
    const res = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body
    });
    if (!res.ok) return null;
    const data = await res.json();
    return data && data.access_token ? data.access_token : null;
  } catch (_) {
    return null;
  }
}

module.exports = { getGatewayAccessToken };
