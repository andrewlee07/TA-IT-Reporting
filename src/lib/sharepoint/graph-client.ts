import { getEnv, requireGraphClientId, requireGraphClientSecret, requireGraphTenantId, requireSharePointDriveId } from "@/lib/env";

const GRAPH_SCOPE = "https://graph.microsoft.com/.default";
const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

type GraphListChildrenResponse = {
  value?: Array<{
    name?: string;
    file?: Record<string, unknown>;
    folder?: Record<string, unknown>;
  }>;
};

function normalizePath(path: string): string {
  return path.replace(/^\/+/, "").replace(/\/{2,}/g, "/");
}

function toGraphPath(path: string): string {
  const normalized = normalizePath(path);
  return normalized.length > 0 ? normalized.split("/").map(encodeURIComponent).join("/") : "";
}

async function requestManagedIdentityToken(clientId: string): Promise<string> {
  const env = getEnv();
  const identityEndpoint = process.env.IDENTITY_ENDPOINT;
  const identityHeader = process.env.IDENTITY_HEADER;

  if (identityEndpoint && identityHeader) {
    const url = new URL(identityEndpoint);
    url.searchParams.set("api-version", "2019-08-01");
    url.searchParams.set("resource", "https://graph.microsoft.com");
    url.searchParams.set("client_id", clientId);

    const response = await fetch(url, {
      headers: {
        "x-identity-header": identityHeader,
        Metadata: "true",
      },
      cache: "no-store",
    });

    if (!response.ok) {
      throw new Error(`Managed identity token request failed with ${response.status}.`);
    }

    const payload = (await response.json()) as { access_token?: string };
    if (!payload.access_token) {
      throw new Error("Managed identity token response did not include an access token.");
    }

    return payload.access_token;
  }

  throw new Error(`Managed identity is configured but IDENTITY_ENDPOINT/IDENTITY_HEADER are unavailable in ${env.AUTH_MODE} mode.`);
}

async function requestClientCredentialsToken(): Promise<string> {
  const tenantId = requireGraphTenantId();
  const clientId = requireGraphClientId();
  const clientSecret = requireGraphClientSecret();
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: GRAPH_SCOPE,
  });

  const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: {
      "content-type": "application/x-www-form-urlencoded",
    },
    body,
    cache: "no-store",
  });

  if (!response.ok) {
    throw new Error(`Client credentials token request failed with ${response.status}.`);
  }

  const payload = (await response.json()) as { access_token?: string };
  if (!payload.access_token) {
    throw new Error("Client credentials token response did not include an access token.");
  }

  return payload.access_token;
}

async function getAccessToken(): Promise<string> {
  const env = getEnv();

  if (env.GRAPH_AUTH_MODE === "disabled") {
    throw new Error("SharePoint storage is configured but GRAPH_AUTH_MODE is disabled.");
  }

  if (env.GRAPH_AUTH_MODE === "managed-identity") {
    return requestManagedIdentityToken(requireGraphClientId());
  }

  return requestClientCredentialsToken();
}

export class SharePointGraphClient {
  private readonly driveId = requireSharePointDriveId();
  private readonly rootPath = normalizePath(getEnv().SHAREPOINT_ROOT_PATH);

  private resolveFullPath(path: string): string {
    const normalized = normalizePath(path);
    return [this.rootPath, normalized].filter(Boolean).join("/");
  }

  private async request(path: string, init?: RequestInit): Promise<Response> {
    const token = await getAccessToken();
    return fetch(`${GRAPH_BASE_URL}${path}`, {
      ...init,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(init?.headers ?? {}),
      },
      cache: "no-store",
    });
  }

  private async ensureFolder(path: string): Promise<void> {
    const fullPath = this.resolveFullPath(path);
    const segments = fullPath.split("/").filter(Boolean);

    let currentPath = "";

    for (const segment of segments) {
      const parentPath = currentPath;
      currentPath = [currentPath, segment].filter(Boolean).join("/");

      const existsResponse = await this.request(`/drives/${this.driveId}/root:/${toGraphPath(currentPath)}`);
      if (existsResponse.ok) {
        continue;
      }

      if (existsResponse.status !== 404) {
        throw new Error(`SharePoint folder lookup failed for ${currentPath} with ${existsResponse.status}.`);
      }

      const createResponse = await this.request(
        `/drives/${this.driveId}/root:/${toGraphPath(parentPath)}:/children`,
        {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            name: segment,
            folder: {},
            "@microsoft.graph.conflictBehavior": "replace",
          }),
        },
      );

      if (!createResponse.ok) {
        throw new Error(`SharePoint folder creation failed for ${currentPath} with ${createResponse.status}.`);
      }
    }
  }

  async putBuffer(path: string, buffer: Buffer, contentType: string): Promise<void> {
    const fullPath = this.resolveFullPath(path);
    const parent = fullPath.split("/").slice(0, -1).join("/");
    if (parent) {
      await this.ensureFolder(parent.replace(`${this.rootPath}/`, ""));
    }

    const response = await this.request(`/drives/${this.driveId}/root:/${toGraphPath(fullPath)}:/content`, {
      method: "PUT",
      headers: {
        "content-type": contentType,
      },
      body: new Uint8Array(buffer),
    });

    if (!response.ok) {
      throw new Error(`SharePoint upload failed for ${path} with ${response.status}.`);
    }
  }

  async getBuffer(path: string): Promise<Buffer> {
    const fullPath = this.resolveFullPath(path);
    const response = await this.request(`/drives/${this.driveId}/root:/${toGraphPath(fullPath)}:/content`);
    if (!response.ok) {
      throw new Error(`SharePoint download failed for ${path} with ${response.status}.`);
    }

    return Buffer.from(await response.arrayBuffer());
  }

  async exists(path: string): Promise<boolean> {
    const fullPath = this.resolveFullPath(path);
    const response = await this.request(`/drives/${this.driveId}/root:/${toGraphPath(fullPath)}`);
    return response.ok;
  }

  async putJson(path: string, value: unknown): Promise<void> {
    await this.putBuffer(path, Buffer.from(`${JSON.stringify(value, null, 2)}\n`, "utf8"), "application/json");
  }

  async getJson<T>(path: string): Promise<T> {
    const buffer = await this.getBuffer(path);
    return JSON.parse(buffer.toString("utf8")) as T;
  }

  async listChildren(path: string): Promise<string[]> {
    const fullPath = this.resolveFullPath(path);
    const response = await this.request(`/drives/${this.driveId}/root:/${toGraphPath(fullPath)}:/children`);

    if (response.status === 404) {
      return [];
    }

    if (!response.ok) {
      throw new Error(`SharePoint list children failed for ${path} with ${response.status}.`);
    }

    const payload = (await response.json()) as GraphListChildrenResponse;
    return (payload.value ?? []).map((entry) => entry.name).filter((name): name is string => typeof name === "string");
  }
}
