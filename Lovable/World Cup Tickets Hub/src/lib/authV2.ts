// =============================================================================
// Story 2.11 / Quartas — Identidade do CLIENTE via Microsoft Entra External ID (CIAM)
// e Identidade ADMIN via Entra Workforce ("dois mundos").
// =============================================================================

import {
  PublicClientApplication,
  type Configuration,
  type AccountInfo,
  type RedirectRequest,
  InteractionRequiredAuthError,
} from '@azure/msal-browser';

// =============================================================================
// 1. CONFIGURAÇÕES CIAM (CLIENTE FINAL) - Usadas no Frontend
// =============================================================================
const ciamClientId = import.meta.env.VITE_CIAM_CLIENT_ID ?? '';
const ciamAuthority = import.meta.env.VITE_CIAM_AUTHORITY ?? '';

// O MSAL exige que o host da authority CIAM/B2C esteja em knownAuthorities.
function deriveKnownAuthorityHost(authority: string): string {
  if (!authority) return '';
  try {
    return new URL(authority).host;
  } catch {
    return '';
  }
}
const knownAuthorityHost = deriveKnownAuthorityHost(ciamAuthority);

// =============================================================================
// 2. CONFIGURAÇÕES ADMIN (WORKFORCE) - Para referência e uso no lab
// =============================================================================
const adminClientId = import.meta.env.VITE_ADMIN_CLIENT_ID ?? '';
const adminTenantId = import.meta.env.VITE_ADMIN_TENANT_ID ?? '';
const adminScope = import.meta.env.VITE_ADMIN_SCOPE ?? '';

// =============================================================================
// 3. VARIÁVEIS COMPARTILHADAS
// =============================================================================
const redirectUri = window.location.origin;

// Scope da API: Tenta o Admin Scope primeiro, com fallback para o CIAM.
const apiScope = adminScope || (ciamClientId ? `api://${ciamClientId}/purchase.write` : 'openid');

/** True quando as variáveis mínimas de identidade do cliente (CIAM) estão configuradas. */
export const isEntraConfigured = (): boolean =>
  Boolean(ciamClientId && ciamAuthority && knownAuthorityHost);

// =============================================================================
// 4. CONFIGURAÇÃO DO MSAL (Apontado para o CIAM / ciamlogin.com)
// =============================================================================
const msalConfig: Configuration = {
  auth: {
    clientId: ciamClientId,
    authority: ciamAuthority,
    knownAuthorities: knownAuthorityHost ? [knownAuthorityHost] : [],
    redirectUri,
    navigateToLoginRequestUrl: false,
  },
  cache: {
    // sessionStorage: token não persiste entre abas/fechamento — mais seguro p/ SPA.
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
};

/**
 * Instância única do MSAL.
 */
export const msalInstance = new PublicClientApplication(msalConfig);

export const loginRequest: RedirectRequest = {
  scopes: [apiScope],
};

/**
 * Obtém um access token v2 silenciosamente; se a sessão exigir interação, cai para popup.
 */
export async function getV2AccessToken(): Promise<string | null> {
  const account: AccountInfo | undefined = msalInstance.getAllAccounts()[0];
  if (!account) {
    return null;
  }

  try {
    const result = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account,
    });
    return result.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const result = await msalInstance.acquireTokenPopup(loginRequest);
      return result.accessToken;
    }
    throw error;
  }
}
