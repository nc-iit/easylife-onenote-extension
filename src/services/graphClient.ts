import { ConfidentialClientApplication } from "@azure/msal-node";

const GRAPH_SCOPE = "https://graph.microsoft.com/.default";

let cca: ConfidentialClientApplication | undefined;

function getConfidentialClient(): ConfidentialClientApplication {
  if (cca) {
    return cca;
  }

  const tenantId = process.env.GRAPH_TENANT_ID;
  const clientId = process.env.GRAPH_CLIENT_ID;
  const clientSecret = process.env.GRAPH_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET must be configured.");
  }

  cca = new ConfidentialClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      clientSecret,
    },
  });
  return cca;
}

/** Acquires an app-only (client credentials) Graph access token. */
export async function getGraphAccessToken(): Promise<string> {
  const client = getConfidentialClient();
  const result = await client.acquireTokenByClientCredential({ scopes: [GRAPH_SCOPE] });
  if (!result?.accessToken) {
    throw new Error("Failed to acquire Graph access token.");
  }
  return result.accessToken;
}

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const MAX_THROTTLE_RETRIES = 4;

export async function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export async function graphFetch(path: string, token: string, init?: RequestInit): Promise<Response> {
  for (let attempt = 0; attempt < MAX_THROTTLE_RETRIES; attempt++) {
    const response = await fetch(`${GRAPH_BASE_URL}${path}`, {
      ...init,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(init?.headers ?? {}),
      },
    });

    if (response.status !== 429 && response.status !== 503) {
      return response;
    }

    const retryAfter = response.headers.get("Retry-After");
    await sleep(retryAfter ? parseFloat(retryAfter) * 1000 : 2000 * (attempt + 1));
  }

  throw new Error(`Graph request ${path} is still throttled after ${MAX_THROTTLE_RETRIES} retries.`);
}
