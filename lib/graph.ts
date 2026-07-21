/**
 * Microsoft Graph client (application permissions / client credentials).
 *
 * Reuses the same Entra app-registration pattern as packages/invoicing:
 * TENANT_ID + CLIENT_ID + CLIENT_SECRET, token via client_credentials,
 * scope https://graph.microsoft.com/.default.
 */

const GRAPH = "https://graph.microsoft.com/v1.0";
const LOGIN = "https://login.microsoftonline.com";

let cachedToken: { token: string; expiresAt: number } | null = null;

function env(name: string, fallback?: string): string {
  const v = process.env[name] ?? fallback;
  if (v === undefined || v === "") {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return v;
}

export function optionalEnv(name: string, fallback: string): string {
  return process.env[name] || fallback;
}

export async function getToken(): Promise<string> {
  // Reuse token until 2 minutes before expiry.
  if (cachedToken && Date.now() < cachedToken.expiresAt - 120_000) {
    return cachedToken.token;
  }
  const tenant = env("TENANT_ID");
  const body = new URLSearchParams({
    client_id: env("CLIENT_ID"),
    client_secret: env("CLIENT_SECRET"),
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const res = await fetch(`${LOGIN}/${tenant}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Token request failed (${res.status}): ${text}`);
  }
  const json = (await res.json()) as { access_token: string; expires_in: number };
  cachedToken = {
    token: json.access_token,
    expiresAt: Date.now() + json.expires_in * 1000,
  };
  return cachedToken.token;
}

/**
 * fetch wrapper with auth + retry on throttling (429) and transient
 * server errors (502/503/504). Honors Retry-After when present.
 */
export async function graphFetch(
  path: string,
  init: RequestInit = {},
  extraHeaders: Record<string, string> = {},
  maxRetries = 4
): Promise<Response> {
  const url = path.startsWith("http") ? path : `${GRAPH}${path}`;
  let attempt = 0;
  // eslint-disable-next-line no-constant-condition
  while (true) {
    const token = await getToken();
    const res = await fetch(url, {
      ...init,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(init.body && !(init.body instanceof Uint8Array)
          ? { "Content-Type": "application/json" }
          : {}),
        ...extraHeaders,
        ...(init.headers as Record<string, string> | undefined),
      },
    });
    if (res.ok) return res;
    const retryable = [429, 502, 503, 504].includes(res.status);
    if (!retryable || attempt >= maxRetries) {
      const text = await res.text().catch(() => "");
      throw new Error(`Graph ${init.method || "GET"} ${path} failed (${res.status}): ${text.slice(0, 500)}`);
    }
    const retryAfter = Number(res.headers.get("Retry-After")) || 2 ** attempt;
    await new Promise((r) => setTimeout(r, retryAfter * 1000));
    attempt++;
  }
}

export async function graphJson<T>(
  path: string,
  init: RequestInit = {},
  extraHeaders: Record<string, string> = {}
): Promise<T> {
  const res = await graphFetch(path, init, extraHeaders);
  return (await res.json()) as T;
}

/**
 * Resolve the drive that holds the pipeline folder.
 * Precedence:
 *   1. DRIVE_ID env (use directly — fastest, no extra calls)
 *   2. SITE_ID env  -> pick drive by DRIVE_NAME (default "Documents")
 *   3. SITE_HOSTNAME + SITE_PATH -> resolve site, then drive by name
 */
export async function resolveDriveId(): Promise<string> {
  if (process.env.DRIVE_ID) return process.env.DRIVE_ID;

  let siteId = process.env.SITE_ID;
  if (!siteId) {
    const hostname = env("SITE_HOSTNAME"); // e.g. eggersco.sharepoint.com
    const sitePath = env("SITE_PATH"); // e.g. /sites/Eggers_GTM_Collission
    const site = await graphJson<{ id: string }>(
      `/sites/${hostname}:${sitePath.startsWith("/") ? sitePath : `/${sitePath}`}`
    );
    siteId = site.id;
  }

  const driveName = optionalEnv("DRIVE_NAME", "Documents");
  const drives = await graphJson<{ value: Array<{ id: string; name: string }> }>(
    `/sites/${siteId}/drives?$select=id,name`
  );
  const drive =
    drives.value.find((d) => d.name === driveName) ??
    drives.value.find((d) => d.name.toLowerCase() === driveName.toLowerCase());
  if (!drive) {
    throw new Error(
      `Drive "${driveName}" not found on site. Available: ${drives.value.map((d) => d.name).join(", ")}`
    );
  }
  return drive.id;
}

export interface DriveItem {
  id: string;
  name: string;
  size: number;
  file?: { mimeType: string };
  lastModifiedDateTime?: string;
}

/** List all files in a drive folder (handles pagination). */
export async function listFolder(driveId: string, folderPath: string): Promise<DriveItem[]> {
  const clean = folderPath.replace(/^\/+|\/+$/g, "");
  let url: string | null =
    `/drives/${driveId}/root:/${clean.split("/").map(encodeURIComponent).join("/")}:/children` +
    `?$select=id,name,size,file,lastModifiedDateTime&$top=200`;
  const items: DriveItem[] = [];
  while (url) {
    const page: { value: DriveItem[]; "@odata.nextLink"?: string } = await graphJson(url);
    items.push(...page.value);
    url = page["@odata.nextLink"] ?? null;
  }
  return items;
}

const SIMPLE_UPLOAD_LIMIT = 3_500_000; // stay safely under Graph's 4 MB simple-upload cap

/** Upload (create or replace) a file in a drive folder. Handles large files via upload session. */
export async function uploadFile(
  driveId: string,
  folderPath: string,
  fileName: string,
  content: Uint8Array
): Promise<void> {
  const clean = folderPath.replace(/^\/+|\/+$/g, "");
  const itemPath = `${clean}/${fileName}`.split("/").map(encodeURIComponent).join("/");

  if (content.byteLength <= SIMPLE_UPLOAD_LIMIT) {
    await graphFetch(
      `/drives/${driveId}/root:/${itemPath}:/content`,
      { method: "PUT", body: content },
      { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
    );
    return;
  }

  // Large file: resumable upload session in 5 MB chunks.
  const session = await graphJson<{ uploadUrl: string }>(
    `/drives/${driveId}/root:/${itemPath}:/createUploadSession`,
    {
      method: "POST",
      body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } }),
    }
  );
  const CHUNK = 5 * 1024 * 1024;
  for (let start = 0; start < content.byteLength; start += CHUNK) {
    const end = Math.min(start + CHUNK, content.byteLength);
    const res = await fetch(session.uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(end - start),
        "Content-Range": `bytes ${start}-${end - 1}/${content.byteLength}`,
      },
      body: content.slice(start, end),
    });
    if (!res.ok && res.status !== 201 && res.status !== 200 && res.status !== 202) {
      throw new Error(`Chunk upload failed (${res.status}): ${await res.text()}`);
    }
  }
}
