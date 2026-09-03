// Office SSO (Entra ID) token resolver & cache

export function parseJwtExpiry(token: string): number | null {
  try {
    const parts = token.split(".");
    if (parts.length >= 2) {
      const base64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
      const jsonStr = atob(base64);
      const payload = JSON.parse(jsonStr);
      if (typeof payload.exp === "number") {
        return payload.exp * 1000;
      }
    }
  } catch {
    // Non-JWT token or decoding error
  }
  return null;
}

let cachedToken: string | null = null;
let cachedExpiresAt = 0;

export async function resolveOfficeSsoToken(
  forceRefresh = false,
): Promise<string> {
  const now = Date.now();
  if (!forceRefresh && cachedToken && cachedExpiresAt > now + 60_000) {
    return cachedToken;
  }

  const runtimeAuth =
    (globalThis as any).OfficeRuntime?.auth || (globalThis as any).Office?.auth;

  if (!runtimeAuth?.getAccessToken) {
    throw new Error(
      "Office SSO getAccessToken is not available in current environment",
    );
  }

  const token = await runtimeAuth.getAccessToken({ allowSignInPrompt: true });
  cachedToken = token;
  const exp = parseJwtExpiry(token);
  cachedExpiresAt = exp ?? now + 5 * 60 * 1000;
  return token;
}

export function clearCachedSsoToken(): void {
  cachedToken = null;
  cachedExpiresAt = 0;
}
