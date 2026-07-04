// =============================================================================
// Story 2.3 / F3 — Identidade v2 com MSAL.js (Authorization Code Flow + PKCE).
//
// Valores de configuração vêm de variáveis Vite (VITE_ADMIN_*) — nunca hardcoded
// (ADE-005 Inv 5 / AC-14).
// =============================================================================

import {
  PublicClientApplication,
  type Configuration,
  type AccountInfo,
  type RedirectRequest,
  InteractionRequiredAuthError,
} from '@azure/msal-browser';

// =============================================================================
// CONFIGURAÇÕES DE IDENTIDADE (ADMIN/WORKFORCE)
// =============================================================================
const clientId = import.meta.env.VITE_ADMIN_CLIENT_ID ?? '';
const tenantId = import.meta.env.VITE_ADMIN_TENANT_ID ?? '';

// Authority do tenant workforce (NÃO 'common' — alinhado ao gateway fail-closed AC-6).
const authority = tenantId
  ? `https://login.microsoftonline.com/${tenantId}`
  : 'https://login.microsoftonline.com/organizations';

const redirectUri = window.location.origin;

// Scope da API exposta pela App Registration (ex.: api://<client-id>/purchase.write).
const apiScope =
  import.meta.env.VITE_ADMIN_SCOPE ??
  (clientId ? `api://${clientId}/purchase.write` : 'openid');

/** True quando as variáveis mínimas de identidade v2 estão configuradas. */
export const isEntraConfigured = (): boolean =>
  Boolean(clientId && tenantId);

// =============================================================================
// CONFIGURAÇÃO DO MSAL
// =============================================================================
const msalConfig: Configuration = {
  auth: {
    clientId,
    authority,
    redirectUri,
    // Não pede consentimento de novo a cada navegação.
    navigateToLoginRequestUrl: false,
  },
  cache: {
    // sessionStorage: token não persiste entre abas/fechamento — mais seguro p/ SPA.
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
};

/**
 * Instância única do MSAL. Deve ser inicializada (await msalInstance.initialize())
 * antes do primeiro uso — feito no bootstrap (main.tsx / MsalProvider).
 */
export const msalInstance = new PublicClientApplication(msalConfig);

/** Scopes solicitados no login v2 (escopo da API + OIDC básico). */
export const loginRequest: RedirectRequest = {
  scopes: [apiScope],
};

/**
 * Obtém um access token v2 silenciosamente (acquireTokenSilent); se a
 * sessão exigir interação (token expirado sem refresh, consent), cai para popup.
 * Retorna null se não houver conta logada.
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
    // Token expirado/sem refresh válido → interação explícita.
    if (error instanceof InteractionRequiredAuthError) {
      const result = await msalInstance.acquireTokenPopup(loginRequest);
      return result.accessToken;
    }
    throw error;
  }
}
