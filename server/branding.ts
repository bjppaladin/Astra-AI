import { storage } from "./storage";
import type { OrganizationBranding } from "@shared/schema";

export interface BrandingConfig {
  companyName: string;
  logoUrl: string | null;
  logoWidthPx: number;
  primaryColor: string;
  secondaryColor: string;
  accentColor: string;
  reportHeaderText: string;
  reportFooterText: string;
  confidentialityNotice: string;
  contactName: string | null;
  contactEmail: string | null;
  contactPhone: string | null;
  website: string | null;
  showMeridianBadge: boolean;
  customCoverPage: boolean;
}

const MERIDIAN_DEFAULTS: BrandingConfig = {
  companyName: "MERIDIAN",
  logoUrl: null,
  logoWidthPx: 200,
  primaryColor: "#1a56db",
  secondaryColor: "#6b7280",
  accentColor: "#059669",
  reportHeaderText: "IT Due Diligence Assessment",
  reportFooterText: "Prepared by MERIDIAN",
  confidentialityNotice: "CONFIDENTIAL — For intended recipients only.",
  contactName: null,
  contactEmail: null,
  contactPhone: null,
  website: null,
  showMeridianBadge: true,
  customCoverPage: false,
};

export async function getBrandingForReport(tenantId: number): Promise<BrandingConfig> {
  const branding = await storage.getBranding(tenantId);

  if (!branding) {
    return { ...MERIDIAN_DEFAULTS };
  }

  return {
    companyName: branding.companyName || MERIDIAN_DEFAULTS.companyName,
    logoUrl: branding.logoUrl || MERIDIAN_DEFAULTS.logoUrl,
    logoWidthPx: branding.logoWidthPx ?? MERIDIAN_DEFAULTS.logoWidthPx,
    primaryColor: branding.primaryColor || MERIDIAN_DEFAULTS.primaryColor,
    secondaryColor: branding.secondaryColor || MERIDIAN_DEFAULTS.secondaryColor,
    accentColor: branding.accentColor || MERIDIAN_DEFAULTS.accentColor,
    reportHeaderText: branding.reportHeaderText || MERIDIAN_DEFAULTS.reportHeaderText,
    reportFooterText: branding.reportFooterText || MERIDIAN_DEFAULTS.reportFooterText,
    confidentialityNotice: branding.confidentialityNotice || MERIDIAN_DEFAULTS.confidentialityNotice,
    contactName: branding.contactName || MERIDIAN_DEFAULTS.contactName,
    contactEmail: branding.contactEmail || MERIDIAN_DEFAULTS.contactEmail,
    contactPhone: branding.contactPhone || MERIDIAN_DEFAULTS.contactPhone,
    website: branding.website || MERIDIAN_DEFAULTS.website,
    showMeridianBadge: branding.showMeridianBadge ?? MERIDIAN_DEFAULTS.showMeridianBadge,
    customCoverPage: branding.customCoverPage ?? MERIDIAN_DEFAULTS.customCoverPage,
  };
}

export { MERIDIAN_DEFAULTS };
