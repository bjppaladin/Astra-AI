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

import { SKU_COST_MAP, findLicenseInfo } from "./sku-map";
const SKU_MAP = SKU_COST_MAP;

export async function fetchLicensedUsers(accessToken: string): Promise<any[]> {
  const skuData = await graphFetch(accessToken, `${GRAPH_BASE}/subscribedSkus?$select=skuId,skuPartNumber,prepaidUnits,consumedUnits`);
  const skus = skuData.value || [];

  const skuIdToInfo: Record<string, { name: string; cost: number }> = {};
  for (const sku of skus) {
    skuIdToInfo[sku.skuId] = findLicenseInfo(sku.skuPartNumber);
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

export interface UserActivity {
  exchangeActive: boolean;
  oneDriveActive: boolean;
  sharePointActive: boolean;
  teamsActive: boolean;
  yammerActive: boolean;
  skypeActive: boolean;
  exchangeLastDate: string | null;
  oneDriveLastDate: string | null;
  sharePointLastDate: string | null;
  teamsLastDate: string | null;
  yammerLastDate: string | null;
  skypeLastDate: string | null;
  activeServiceCount: number;
  totalServiceCount: number;
  daysSinceLastActivity: number | null;
}

export async function fetchUserActivityMap(accessToken: string): Promise<Map<string, UserActivity>> {
  let csvData: string;
  try {
    csvData = await graphFetch(accessToken, `${GRAPH_BETA}/reports/getOffice365ActiveUserDetail(period='D30')?$format=text/csv`, true);
  } catch {
    return new Map();
  }

  const lines = (csvData as string).split("\n").filter((l: string) => l.trim());
  if (lines.length < 2) return new Map();

  const headers = parseCSVLine(lines[0]);
  const findCol = (keyword: string) => headers.findIndex((h) => h.toLowerCase().includes(keyword.toLowerCase()));

  const upnIdx = findCol("user principal name");
  if (upnIdx === -1) return new Map();

  const exchangeDateIdx = findCol("exchange last activity date");
  const oneDriveDateIdx = findCol("onedrive last activity date");
  const sharePointDateIdx = findCol("sharepoint last activity date");
  const teamsDateIdx = findCol("teams last activity date");
  const yammerDateIdx = findCol("yammer last activity date");
  const skypeDateIdx = findCol("skype for business last activity date");

  const hasExchangeLicIdx = findCol("has exchange license");
  const hasOneDriveLicIdx = findCol("has onedrive license");
  const hasSharePointLicIdx = findCol("has sharepoint license");
  const hasTeamsLicIdx = findCol("has teams license");
  const hasYammerLicIdx = findCol("has yammer license");
  const hasSkypeLicIdx = findCol("has skype for business license");

  const result = new Map<string, UserActivity>();
  const now = new Date();

  for (let i = 1; i < lines.length; i++) {
    const fields = parseCSVLine(lines[i]);
    if (fields.length <= upnIdx) continue;

    const upn = fields[upnIdx]?.trim();
    if (!upn) continue;

    const getDate = (idx: number): string | null => {
      if (idx === -1 || idx >= fields.length) return null;
      const val = fields[idx]?.trim();
      return val && val !== "" ? val : null;
    };

    const isActive = (dateStr: string | null): boolean => {
      return dateStr !== null && dateStr !== "";
    };

    const hasLicense = (idx: number): boolean => {
      if (idx === -1 || idx >= fields.length) return false;
      return fields[idx]?.trim().toLowerCase() === "true";
    };

    const exchangeLastDate = getDate(exchangeDateIdx);
    const oneDriveLastDate = getDate(oneDriveDateIdx);
    const sharePointLastDate = getDate(sharePointDateIdx);
    const teamsLastDate = getDate(teamsDateIdx);
    const yammerLastDate = getDate(yammerDateIdx);
    const skypeLastDate = getDate(skypeDateIdx);

    const exchangeActive = isActive(exchangeLastDate);
    const oneDriveActive = isActive(oneDriveLastDate);
    const sharePointActive = isActive(sharePointLastDate);
    const teamsActive = isActive(teamsLastDate);
    const yammerActive = isActive(yammerLastDate);
    const skypeActive = isActive(skypeLastDate);

    const activeServices = [exchangeActive, oneDriveActive, sharePointActive, teamsActive, yammerActive, skypeActive];
    const activeServiceCount = activeServices.filter(Boolean).length;

    const licensedServices = [
      hasLicense(hasExchangeLicIdx),
      hasLicense(hasOneDriveLicIdx),
      hasLicense(hasSharePointLicIdx),
      hasLicense(hasTeamsLicIdx),
      hasLicense(hasYammerLicIdx),
      hasLicense(hasSkypeLicIdx),
    ];
    const totalServiceCount = Math.max(licensedServices.filter(Boolean).length, activeServiceCount, 1);

    const allDates = [exchangeLastDate, oneDriveLastDate, sharePointLastDate, teamsLastDate, yammerLastDate, skypeLastDate]
      .filter((d): d is string => d !== null)
      .map((d) => new Date(d).getTime())
      .filter((t) => !isNaN(t));

    let daysSinceLastActivity: number | null = null;
    if (allDates.length > 0) {
      const mostRecent = Math.max(...allDates);
      daysSinceLastActivity = Math.floor((now.getTime() - mostRecent) / (1000 * 60 * 60 * 24));
    }

    result.set(upn.toLowerCase(), {
      exchangeActive,
      oneDriveActive,
      sharePointActive,
      teamsActive,
      yammerActive,
      skypeActive,
      exchangeLastDate,
      oneDriveLastDate,
      sharePointLastDate,
      teamsLastDate,
      yammerLastDate,
      skypeLastDate,
      activeServiceCount,
      totalServiceCount,
      daysSinceLastActivity,
    });
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

export async function fetchSubscribedSkus(accessToken: string): Promise<any[]> {
  const data = await graphFetch(accessToken, `${GRAPH_BASE}/subscribedSkus?$select=skuId,skuPartNumber,prepaidUnits,consumedUnits,capabilityStatus,appliesTo`);
  const skus = data.value || [];
  return skus.map((sku: any) => {
    const mapped = findLicenseInfo(sku.skuPartNumber);
    const enabled = sku.prepaidUnits?.enabled ?? 0;
    const consumed = sku.consumedUnits ?? 0;
    return {
      skuId: sku.skuId,
      skuPartNumber: sku.skuPartNumber,
      displayName: mapped.name,
      costPerUser: mapped.cost,
      enabled,
      consumed,
      available: enabled - consumed,
      capabilityStatus: sku.capabilityStatus || "Enabled",
      appliesTo: sku.appliesTo || "User",
    };
  });
}

export async function fetchM365Data(accessToken: string): Promise<any[]> {
  const [users, mailboxMap, activityMap] = await Promise.all([
    fetchLicensedUsers(accessToken),
    fetchMailboxUsage(accessToken),
    fetchUserActivityMap(accessToken),
  ]);

  return users.map((user) => {
    const mailbox = mailboxMap.get(user.upn.toLowerCase());
    const activity = activityMap.get(user.upn.toLowerCase()) ?? null;
    return {
      ...user,
      usageGB: mailbox?.usageGB ?? 0,
      maxGB: mailbox?.maxGB ?? 50,
      status: mailbox
        ? mailbox.usageGB / mailbox.maxGB > 0.9 ? "Critical"
          : mailbox.usageGB / mailbox.maxGB > 0.7 ? "Warning" : "Active"
        : "Active",
      activity,
    };
  });
}
