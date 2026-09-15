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

/** Acquires an app-only (client credentials) Graph access token using Notes.ReadWrite.All / Group.ReadWrite.All. */
export async function getGraphAccessToken(): Promise<string> {
  const client = getConfidentialClient();
  const result = await client.acquireTokenByClientCredential({ scopes: [GRAPH_SCOPE] });
  if (!result?.accessToken) {
    throw new Error("Failed to acquire Graph access token.");
  }
  return result.accessToken;
}

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

export async function graphFetch(
  path: string,
  token: string,
  init?: RequestInit
): Promise<Response> {
  const response = await fetch(`${GRAPH_BASE_URL}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(init?.headers ?? {}),
    },
  });
  return response;
}
