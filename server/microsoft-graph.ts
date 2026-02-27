const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const GRAPH_BETA = "https://graph.microsoft.com/beta";

const SCOPES = [
  "User.Read",
  "User.Read.All",
  "Reports.Read.All",
  "Organization.Read.All",
  "offline_access",
];

function getClientId(): string {
  const id = process.env.MICROSOFT_CLIENT_ID;
  if (!id) throw new Error("MICROSOFT_CLIENT_ID not configured");
  return id;
}

function getClientSecret(): string {
  const secret = process.env.MICROSOFT_CLIENT_SECRET;
  if (!secret) throw new Error("MICROSOFT_CLIENT_SECRET not configured");
  return secret;
}

function getTenantId(): string {
  return process.env.MICROSOFT_TENANT_ID || "common";
}

export function isOAuthConfigured(): boolean {
  return !!(process.env.MICROSOFT_CLIENT_ID && process.env.MICROSOFT_CLIENT_SECRET);
}

export function getAuthUrl(redirectUri: string, state: string): string {
  const scope = SCOPES.join(" ");
  const clientId = getClientId();
  return `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${encodeURIComponent(redirectUri)}&scope=${encodeURIComponent(scope)}&state=${state}&response_mode=query&prompt=consent`;
}

function decodeJwtPayload(token: string): Record<string, any> {
  try {
    const parts = token.split(".");
    if (parts.length !== 3) return {};
    const payload = Buffer.from(parts[1], "base64url").toString("utf-8");
    return JSON.parse(payload);
  } catch {
    return {};
  }
}

export async function exchangeCodeForTokens(
  code: string,
  redirectUri: string
): Promise<{
  accessToken: string;
  refreshToken: string | undefined;
  expiresAt: Date;
  tenantId: string;
}> {
  const params = new URLSearchParams({
    client_id: getClientId(),
    client_secret: getClientSecret(),
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    scope: SCOPES.join(" "),
  });

  const response = await fetch(
    `https://login.microsoftonline.com/common/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: params.toString() }
  );

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Token exchange failed: ${errText}`);
  }

  const data = await response.json();
  const claims = decodeJwtPayload(data.access_token);
  const tenantId = claims.tid || "unknown";

  if (claims.aud && claims.aud !== "https://graph.microsoft.com" && claims.aud !== "00000003-0000-0000-c000-000000000000") {
    throw new Error(`Unexpected token audience: ${claims.aud}`);
  }
  if (claims.iss && !claims.iss.startsWith("https://sts.windows.net/") && !claims.iss.startsWith("https://login.microsoftonline.com/")) {
    throw new Error(`Unexpected token issuer: ${claims.iss}`);
  }

  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    expiresAt: new Date(Date.now() + data.expires_in * 1000),
    tenantId,
  };
}

export async function refreshAccessToken(refreshToken: string): Promise<{
  accessToken: string;
  refreshToken: string | undefined;
  expiresAt: Date;
}> {
  const params = new URLSearchParams({
    client_id: getClientId(),
    client_secret: getClientSecret(),
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: SCOPES.join(" "),
  });

  const response = await fetch(
    `https://login.microsoftonline.com/common/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: params.toString() }
  );

  if (!response.ok) throw new Error("Token refresh failed");
  const data = await response.json();
  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    expiresAt: new Date(Date.now() + data.expires_in * 1000),
  };
}

async function graphFetch(accessToken: string, url: string, isCSV = false): Promise<any> {
  const headers: Record<string, string> = { Authorization: `Bearer ${accessToken}` };
  if (!isCSV) headers["Content-Type"] = "application/json";
  headers["ConsistencyLevel"] = "eventual";

  const response = await fetch(url, { headers });
  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Graph API error (${response.status}): ${errText}`);
  }
  return isCSV ? response.text() : response.json();
}

export async function getCurrentUser(accessToken: string): Promise<{ displayName: string; mail: string }> {
  const data = await graphFetch(accessToken, `${GRAPH_BASE}/me?$select=displayName,mail,userPrincipalName`);
  return { displayName: data.displayName, mail: data.mail || data.userPrincipalName };
}

const SKU_MAP: Record<string, { name: string; cost: number }> = {
  "SPE_E5": { name: "Microsoft 365 E5", cost: 57.00 },
  "SPE_E3": { name: "Microsoft 365 E3", cost: 36.00 },
  "STANDARDPACK": { name: "Office 365 E1", cost: 10.00 },
  "SPE_F1": { name: "Microsoft 365 F1", cost: 2.25 },
  "ENTERPRISEPREMIUM": { name: "Office 365 E5", cost: 38.00 },
  "ENTERPRISEPACK": { name: "Office 365 E3", cost: 23.00 },
  "VISIOCLIENT": { name: "Visio Plan 2", cost: 15.00 },
  "PROJECTPREMIUM": { name: "Project Plan 5", cost: 55.00 },
  "PROJECTPROFESSIONAL": { name: "Project Plan 3", cost: 30.00 },
  "POWER_BI_PRO": { name: "Power BI Pro", cost: 10.00 },
  "POWER_BI_PREMIUM_PER_USER": { name: "Power BI Premium Per User", cost: 20.00 },
  "Microsoft_365_Copilot": { name: "Microsoft 365 Copilot", cost: 30.00 },
  "EXCHANGESTANDARD": { name: "Exchange Online Plan 1", cost: 4.00 },
  "EXCHANGEENTERPRISE": { name: "Exchange Online Plan 2", cost: 8.00 },
  "O365_BUSINESS_ESSENTIALS": { name: "Microsoft 365 Business Basic", cost: 6.00 },
  "O365_BUSINESS_PREMIUM": { name: "Microsoft 365 Business Standard", cost: 12.50 },
  "SPB": { name: "Microsoft 365 Business Premium", cost: 22.00 },
  "TEAMS_EXPLORATORY": { name: "Teams Exploratory", cost: 0 },
  "FLOW_FREE": { name: "Power Automate Free", cost: 0 },
  "POWERAPPS_VIRAL": { name: "Power Apps Trial", cost: 0 },
  "STREAM": { name: "Microsoft Stream", cost: 0 },
};

export async function fetchLicensedUsers(accessToken: string): Promise<any[]> {
  const skuData = await graphFetch(accessToken, `${GRAPH_BASE}/subscribedSkus?$select=skuId,skuPartNumber,prepaidUnits,consumedUnits`);
  const skus = skuData.value || [];

  const skuIdToInfo: Record<string, { name: string; cost: number }> = {};
  for (const sku of skus) {
    const mapped = SKU_MAP[sku.skuPartNumber];
    skuIdToInfo[sku.skuId] = mapped || { name: sku.skuPartNumber, cost: 0 };
  }

  let allUsers: any[] = [];
  let nextLink: string | null = `${GRAPH_BASE}/users?$select=id,displayName,userPrincipalName,department,assignedLicenses,accountEnabled&$filter=assignedLicenses/$count ne 0&$count=true&$top=999`;

  while (nextLink) {
    const data = await graphFetch(accessToken, nextLink);
    allUsers = allUsers.concat(data.value || []);
    nextLink = data["@odata.nextLink"] || null;
  }

  return allUsers
    .filter((u) => u.accountEnabled && u.assignedLicenses?.length > 0)
    .map((u) => {
      const licenses: string[] = [];
      let cost = 0;
      for (const al of u.assignedLicenses) {
        const info = skuIdToInfo[al.skuId];
        if (info) { licenses.push(info.name); cost += info.cost; }
      }
      return {
        id: u.id,
        displayName: u.displayName,
        upn: u.userPrincipalName,
        department: u.department || "Unassigned",
        licenses,
        cost,
        status: "Active",
      };
    });
}

function parseCSVLine(line: string): string[] {
  const result: string[] = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    if (char === '"') { inQuotes = !inQuotes; }
    else if (char === "," && !inQuotes) { result.push(current.trim()); current = ""; }
    else { current += char; }
  }
  result.push(current.trim());
  return result;
}

export async function fetchMailboxUsage(accessToken: string): Promise<Map<string, { usageGB: number; maxGB: number }>> {
  let csvData: string;
  try {
    csvData = await graphFetch(accessToken, `${GRAPH_BETA}/reports/getMailboxUsageDetail(period='D7')?$format=text/csv`, true);
  } catch {
    return new Map();
  }

  const lines = (csvData as string).split("\n").filter((l: string) => l.trim());
  if (lines.length < 2) return new Map();

  const headers = parseCSVLine(lines[0]);
  const upnIdx = headers.findIndex((h) => h.toLowerCase().includes("user principal name"));
  const storageIdx = headers.findIndex((h) => h.toLowerCase().includes("storage used"));
  const quotaIdx = headers.findIndex((h) => h.toLowerCase().includes("prohibit send/receive quota") || h.toLowerCase().includes("issue warning quota"));

  if (upnIdx === -1 || storageIdx === -1) return new Map();

  const result = new Map<string, { usageGB: number; maxGB: number }>();
  for (let i = 1; i < lines.length; i++) {
    const fields = parseCSVLine(lines[i]);
    if (fields.length <= Math.max(upnIdx, storageIdx)) continue;
    const upn = fields[upnIdx];
    const storageBytes = parseFloat(fields[storageIdx]) || 0;
    const quotaBytes = quotaIdx >= 0 ? parseFloat(fields[quotaIdx]) || 0 : 0;
    const usageGB = storageBytes / (1024 * 1024 * 1024);
    const maxGB = quotaBytes > 0 ? quotaBytes / (1024 * 1024 * 1024) : 50;
    result.set(upn.toLowerCase(), { usageGB: Math.round(usageGB * 10) / 10, maxGB: Math.round(maxGB) });
  }
  return result;
}

export async function fetchActiveUserDetailReport(accessToken: string): Promise<any[]> {
  let csvData: string;
  try {
    csvData = await graphFetch(accessToken, `${GRAPH_BETA}/reports/getOffice365ActiveUserDetail(period='D30')?$format=text/csv`, true);
  } catch (err: any) {
    throw new Error(`Failed to fetch Active User Detail report: ${err.message}`);
  }

  const lines = (csvData as string).split("\n").filter((l: string) => l.trim());
  if (lines.length < 2) return [];

  const headers = parseCSVLine(lines[0]);
  const result: any[] = [];

  for (let i = 1; i < lines.length; i++) {
    const fields = parseCSVLine(lines[i]);
    if (fields.length < headers.length) continue;

    const row: Record<string, string> = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = fields[j] || "";
    }
    result.push(row);
  }

  return result;
}

export async function fetchM365Data(accessToken: string): Promise<any[]> {
  const [users, mailboxMap] = await Promise.all([
    fetchLicensedUsers(accessToken),
    fetchMailboxUsage(accessToken),
  ]);

  return users.map((user) => {
    const mailbox = mailboxMap.get(user.upn.toLowerCase());
    return {
      ...user,
      usageGB: mailbox?.usageGB ?? 0,
      maxGB: mailbox?.maxGB ?? 50,
      status: mailbox
        ? mailbox.usageGB / mailbox.maxGB > 0.9 ? "Critical"
          : mailbox.usageGB / mailbox.maxGB > 0.7 ? "Warning" : "Active"
        : "Active",
    };
  });
}
