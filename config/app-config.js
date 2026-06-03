/**
 * Asterwix Unified Web Portal — Configuration File
 * =====================================================================
 * IMPORTANT: Ensure this file is listed in .gitignore for security.
 * =====================================================================
 */

window.ASTERWIX_CONFIG = {

  // Live Microsoft 365 authentication is now ENABLED
  MOCK_MODE_ENABLED: false,

  // Azure AD (Entra ID) — Asterwix Portal App Registration
  auth: {
    clientId: 'bf5114ed-db8d-4137-b2a3-2d4ea4cd25e8',
    authority: 'https://login.microsoftonline.com/16522c0a-f9da-4e67-ac59-ed6adea4e80e',
    redirectUri: window.location.origin + window.location.pathname,
    productionRedirectUri: 'https://hub.asterwix.com/'
  },

  // SharePoint Online Site Configuration (optional)
  sharePoint: {
    hostname: 'asterwix.sharepoint.com',
    siteName: 'AsterwixHub',
    siteId: '' // Resolved dynamically if left empty
  }
};
