export type LicensePlan = "free" | "viewer-commercial" | "sdk-commercial" | "enterprise" | "evaluation";
export type LicenseStatus = "open-source" | "valid" | "grace" | "local" | "missing" | "invalid" | "origin-mismatch" | "expired";
export type LicenseBranding = "required" | "hidden" | "evaluation" | "unauthorized";

export interface LicenseClaims {
  readonly licenseId: string;
  readonly customerId: string;
  readonly applicationId: string;
  readonly plan: LicensePlan;
  readonly origins: readonly string[];
  readonly features: readonly string[];
  readonly versionBefore?: string;
  readonly issuedAt: number;
  readonly expiresAt: number;
}

export interface LicenseOptions {
  readonly token: string;
  /** @deprecated Runtime licenses are ignored in the Apache-2.0 release. */
  readonly publicKey: string;
  /** @deprecated No origin restriction applies. */
  readonly origin?: string;
}

export interface LicenseEntitlements {
  readonly viewer: true;
  readonly engine: boolean;
  readonly businessSlots: boolean;
  readonly advancedCustomization: boolean;
}

export interface LicenseState {
  readonly status: LicenseStatus;
  readonly plan?: LicensePlan;
  readonly branding: LicenseBranding;
  readonly entitlements: LicenseEntitlements;
  readonly claims?: LicenseClaims;
}

/** Compatibility state for applications migrating from the proprietary SDK. */
export const OPEN_SOURCE_LICENSE: LicenseState = Object.freeze({
  status: "open-source",
  branding: "hidden",
  entitlements: Object.freeze({
    viewer: true,
    engine: true,
    businessSlots: true,
    advancedCustomization: true,
  }),
});

/** @deprecated All capabilities are available under Apache-2.0; options are ignored. */
export async function evaluateLicense(_options?: LicenseOptions): Promise<LicenseState> {
  return OPEN_SOURCE_LICENSE;
}
